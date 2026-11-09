#!/usr/bin/env python3
"""Replace {{MATH:id}} runs with native PowerPoint OMML plus a picture fallback.

Does not author a deck by itself. A formula spec resolves in this order:
``runs``, then ``omml``, then ``latex``. LaTeX is compiled with the vendored
ppt-master compiler in block mode, so summation limits use ``undOvr``. This
module keeps its own ``m:oMathPara`` wrapper. Pandoc is used only when that
compiler raises and pandoc is installed. Otherwise the command fails.

The equation is written as ``mc:AlternateContent`` / ``mc:Choice Requires="a14"``
/ ``a14:m`` / ``m:oMathPara`` / ``m:oMath`` (MS-ODRAWXML). ``mc:Fallback`` is a
high-DPI transparent PNG of the same formula, fitted to the box without stretching.
When the spec has LaTeX and a TeX engine plus a PNG converter are installed, that
PNG is a real TeX render (see ``tex_fallback.py``). Otherwise it is the mathtext
picture: ``\\frac`` is rewritten to ``\\dfrac`` for the picture only. The string
sent to the OMML compiler is not rewritten. A missing TeX install does not fail
the page. LaTeX source is never written onto a slide.

The mathtext rung needs matplotlib. The TeX rung does not, when it succeeds.
"""
import argparse
import importlib.util
import io
import json
import os
import posixpath
import re
import shutil
import struct
import subprocess
import sys
import tempfile
import types
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from xml.sax.saxutils import escape

# Keep aligned with scripts/audit_pptx.py RAW_LATEX.
# Mathtext has no style switches. Dropping these leaves the same symbols;
# commands it still cannot draw fail the render instead of leaking source.
_MATHTEXT_DROP = re.compile(
    r"\\(?:displaystyle|textstyle|scriptstyle|scriptscriptstyle|limits|nolimits)\b"
)
# ``\frac`` is text style in mathtext and draws small beside 18–20pt type.
# ``\dfrac`` / ``\tfrac`` do not contain the command ``\frac``.
_FRAC_CMD = re.compile(r"\\frac\b")
# Display fractions and roots need a larger nominal size than a single letter.
_TALL_MATH = re.compile(r"\\(?:dfrac|sqrt)\b")
RAW_LATEX = re.compile(
    r"\\(?:frac|dfrac|tfrac|sum|prod|int|iint|iiint|lambda|alpha|beta|gamma|theta|sigma|"
    r"mu|nu|omega|phi|psi|delta|epsilon|left|right|cdot|times|sqrt|mathbb|mathrm|mathbf|"
    r"text|begin|end|overline|hat|bar|vec|partial|infty|leq|geq|neq|approx|equiv|"
    r"subset|rightarrow|leftarrow)\b"
    r"|\\[{}]"
    r"|\^\{"
    r"|_\{"
)
PLACEHOLDER = re.compile(r"\{\{MATH:([A-Za-z0-9_-]+)\}\}")
RUN_BLOCK = re.compile(r"<a:r\b[^>]*>.*?</a:r>", re.DOTALL)
TEXT_NODE = re.compile(r"<a:t(?:\s[^>]*)?>(.*?)</a:t>", re.DOTALL)
ID_RE = re.compile(r"[A-Za-z0-9_-]+")
MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
A14_NS = "http://schemas.microsoft.com/office/drawing/2010/main"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS = {"m": MATH_NS}
EMU_IN = 914400
TAG_RE = re.compile(r"<(/?)([A-Za-z0-9]+:[A-Za-z0-9]+)([^>]*)>")
_GREEK = {
    "α": r"\alpha", "β": r"\beta", "γ": r"\gamma", "δ": r"\delta",
    "ε": r"\epsilon", "θ": r"\theta", "λ": r"\lambda", "μ": r"\mu",
    "ν": r"\nu", "π": r"\pi", "ρ": r"\rho", "σ": r"\sigma",
    "τ": r"\tau", "φ": r"\phi", "ψ": r"\psi", "ω": r"\omega",
    "Σ": r"\Sigma", "Π": r"\Pi", "Ω": r"\Omega", "∞": r"\infty",
    "≤": r"\leq", "≥": r"\geq", "≠": r"\neq", "≈": r"\approx",
    "×": r"\times", "·": r"\cdot", "−": "-", "–": "-", "—": "-",
}


def assert_no_raw_latex(text, where):
    if text and RAW_LATEX.search(text):
        raise ValueError(f"Refusing to write LaTeX source into {where}")


def _piece(item):
    if not isinstance(item, dict) or len(item) != 1:
        raise ValueError("Each formula run must be an object with one key")
    kind, body = next(iter(item.items()))
    if kind == "t":
        if not isinstance(body, str) or not body:
            raise ValueError("Text run must be a non-empty string")
        assert_no_raw_latex(body, "a formula run")
        return f"<m:r><m:t>{escape(body)}</m:t></m:r>"
    if kind not in {"sub", "sup", "frac"}:
        raise ValueError("Unknown formula run kind: " + kind)
    if not isinstance(body, dict):
        raise ValueError(kind + " run needs an object")
    if kind == "frac":
        num, den = body.get("num"), body.get("den")
        if not isinstance(num, str) or not isinstance(den, str):
            raise ValueError("frac needs string num and den")
        assert_no_raw_latex(num, "a numerator")
        assert_no_raw_latex(den, "a denominator")
        return (
            "<m:f><m:num>" + _piece({"t": num}) + "</m:num><m:den>"
            + _piece({"t": den}) + "</m:den></m:f>"
        )
    base, script = body.get("base"), body.get(kind)
    if not isinstance(base, str) or not isinstance(script, str):
        raise ValueError(kind + " needs string base and " + kind)
    assert_no_raw_latex(base, "a formula base")
    assert_no_raw_latex(script, "a formula script")
    outer = "sSub" if kind == "sub" else "sSup"
    inner = "sub" if kind == "sub" else "sup"
    return (
        f"<m:{outer}><m:e>{_piece({'t': base})}</m:e>"
        f"<m:{inner}>{_piece({'t': script})}</m:{inner}></m:{outer}>"
    )


def omml_from_runs(runs):
    if not isinstance(runs, list) or not runs:
        raise ValueError("runs must be a non-empty list")
    inner = "".join(_piece(item) for item in runs)
    return (
        '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
        + inner
        + "</m:oMath>"
    )


def _local(tag):
    if tag.startswith("{"):
        ns, name = tag[1:].split("}", 1)
        if ns == MATH_NS:
            return "m:" + name
        if ns == "http://www.w3.org/XML/1998/namespace":
            return "xml:" + name
    return tag


def serialize_omath(elem):
    """Emit an m:oMath element with a stable prefix."""
    if _local(elem.tag) != "m:oMath":
        raise ValueError("OMML root must be m:oMath")

    def rec(node, root=False):
        attrs = []
        if root:
            attrs.append('xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"')
        for key, value in node.attrib.items():
            attrs.append(f'{_local(key)}="{escape(value, {chr(34): "&quot;"})}"')
        body = escape(node.text or "")
        for child in list(node):
            body += rec(child)
            body += escape(child.tail or "")
        attr = (" " + " ".join(attrs)) if attrs else ""
        return f"<{_local(node.tag)}{attr}>{body}</{_local(node.tag)}>"

    xml = rec(elem, True)
    assert_no_raw_latex(xml, "OMML")
    return xml


def _load_formula_compiler():
    """Load the vendored ppt-master compiler. The five modules stay unchanged."""
    package_name = "pptmaster_formula"
    loaded = sys.modules.get(package_name + ".formula_compiler")
    if loaded is not None and hasattr(loaded, "compile_latex_to_omml"):
        return loaded
    root = Path(__file__).resolve().parent / "vendor" / "pptmaster_formula"
    compiler_path = root / "formula_compiler.py"
    if not compiler_path.is_file():
        raise ValueError("Vendored formula compiler is missing")
    package = sys.modules.get(package_name)
    if package is None:
        package = types.ModuleType(package_name)
        package.__path__ = [str(root)]
        package.__package__ = package_name
        sys.modules[package_name] = package
    spec = importlib.util.spec_from_file_location(
        package_name + ".formula_compiler", compiler_path
    )
    module = importlib.util.module_from_spec(spec)
    module.__package__ = package_name
    sys.modules[package_name + ".formula_compiler"] = module
    spec.loader.exec_module(module)
    return module


def _omath_from_block(block_xml):
    """Take the inner m:oMath from a block m:oMathPara compile."""
    root = ET.fromstring(block_xml)
    found = None
    for child in root.iter():
        if _local(child.tag) == "m:oMath":
            found = child
            break
    if found is None:
        raise ValueError("Formula compiler did not produce m:oMath")
    return serialize_omath(found)


def omml_from_vendor(latex):
    """Block-compile LaTeX. Summation limits are undOvr. Unsupported commands raise."""
    compiler = _load_formula_compiler()
    try:
        block = compiler.compile_latex_to_omml(latex)
    except compiler.FormulaCompileError as exc:
        raise ValueError(str(exc)) from exc
    return _omath_from_block(block)


def omml_from_latex(latex):
    """Vendored compiler, then pandoc if that raises and pandoc exists, else fail."""
    source = _unwrap_math_dollars(latex)
    if not source:
        raise ValueError("latex must be a non-empty string")
    try:
        return omml_from_vendor(source)
    except ValueError as compiler_error:
        compiled = None
        try:
            compiled = omml_from_pandoc(source)
        except ValueError:
            compiled = None
        if compiled:
            return compiled
        raise compiler_error


def omml_from_pandoc(latex):
    if shutil.which("pandoc") is None:
        return None
    with tempfile.TemporaryDirectory() as folder:
        source = Path(folder) / "formula.md"
        docx = Path(folder) / "formula.docx"
        source.write_text("$$\n" + latex.strip() + "\n$$\n", encoding="utf-8")
        completed = subprocess.run(
            ["pandoc", str(source), "-o", str(docx)],
            capture_output=True,
            text=True,
        )
        if completed.returncode != 0 or not docx.is_file():
            raise ValueError("pandoc could not convert LaTeX to OMML")
        with zipfile.ZipFile(docx) as package:
            document = ET.fromstring(package.read("word/document.xml"))
    found = document.find(".//m:oMath", NS)
    if found is None:
        raise ValueError("pandoc did not produce OMML")
    return serialize_omath(found)


def omml_for_spec(spec):
    if not isinstance(spec, dict):
        raise ValueError("Each formula spec must be an object")
    if spec.get("runs"):
        return omml_from_runs(spec["runs"])
    if spec.get("omml"):
        if not isinstance(spec["omml"], str) or "oMath" not in spec["omml"]:
            raise ValueError("omml must be an m:oMath XML string")
        root = ET.fromstring(spec["omml"])
        return serialize_omath(root)
    if spec.get("latex"):
        if not isinstance(spec["latex"], str) or not spec["latex"].strip():
            raise ValueError("latex must be a non-empty string")
        return omml_from_latex(spec["latex"])
    raise ValueError("Formula spec needs runs, omml, or latex")


def load_formulas(path):
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    if not isinstance(data, dict) or not data:
        raise ValueError("formulas file must be a non-empty object")
    for key in data:
        if not isinstance(key, str) or not ID_RE.fullmatch(key):
            raise ValueError("Bad formula id: " + str(key))
    return data


def _math_fragment(text):
    parts = []
    for char in text:
        if char in _GREEK:
            parts.append(_GREEK[char] + " ")
        elif char == " ":
            parts.append(r"\ ")
        elif char in "\\{}$%#_&^~":
            parts.append("\\" + char)
        else:
            parts.append(char)
    return "".join(parts)


def _runs_mathtext(runs):
    pieces = []
    for item in runs:
        kind, body = next(iter(item.items()))
        if kind == "t":
            pieces.append(_math_fragment(body))
        elif kind == "frac":
            pieces.append(r"\dfrac{" + _math_fragment(body["num"]) + "}{" + _math_fragment(body["den"]) + "}")
        else:
            pieces.append(_math_fragment(body["base"]) + ("_{" if kind == "sub" else "^{") + _math_fragment(body[kind]) + "}")
    return "".join(pieces)


def _omml_plain(omml_xml):
    chunks = re.findall(r"<m:t(?:\s[^>]*)?>(.*?)</m:t>", omml_xml, re.DOTALL)
    visible = "".join(chunks)
    return (
        visible.replace("&lt;", "<").replace("&gt;", ">")
        .replace("&quot;", '"').replace("&amp;", "&")
    )


def _unwrap_math_dollars(latex):
    text = latex.strip()
    if text.startswith("$$") and text.endswith("$$") and len(text) >= 4:
        return text[2:-2].strip()
    if len(text) >= 2 and text.startswith("$") and text.endswith("$"):
        return text[1:-1].strip()
    return text


def _display_fractions(text):
    """Rewrite ``\\frac`` to ``\\dfrac`` for the fallback picture only."""
    return _FRAC_CMD.sub(lambda _match: r"\dfrac", text)


def _mathtext_latex(latex):
    """Same LaTeX string mathtext will draw. Style commands it rejects are removed.

    ``\\frac`` becomes ``\\dfrac`` here. ``omml_from_latex`` still compiles the
    original string, so the OMML fraction stays ``m:f``.
    """
    text = _MATHTEXT_DROP.sub("", _unwrap_math_dollars(latex))
    text = text.replace(r"\operatorname*{", r"\operatorname{")
    text = _display_fractions(text)
    return text.strip()


def mathtext_for_spec(spec, omml_xml):
    """Mathtext for the fallback image only. This string is not written into the pptx.

    A ``latex`` spec is drawn from that string, not rebuilt from ``runs``.
    """
    if isinstance(spec, dict) and isinstance(spec.get("latex"), str) and spec["latex"].strip():
        body = _mathtext_latex(spec["latex"])
        if not body:
            raise ValueError("latex did not produce a mathtext expression")
        return "$" + body + "$"
    if isinstance(spec, dict) and spec.get("runs"):
        return "$" + _runs_mathtext(spec["runs"]) + "$"
    return _omml_plain(omml_xml) or "?"


def _png_figure(draw, width_in, height_in):
    import matplotlib
    matplotlib.use("Agg", force=False)
    import matplotlib.pyplot as plt
    figure = plt.figure(figsize=(width_in, height_in), dpi=400)
    figure.patch.set_alpha(0)
    figure.patch.set_facecolor("none")
    axis = figure.add_axes([0, 0, 1, 1])
    axis.set_axis_off()
    axis.patch.set_alpha(0)
    axis.set_facecolor("none")
    axis.set_xlim(0, 1)
    axis.set_ylim(0, 1)
    try:
        draw(axis, height_in)
        buffer = io.BytesIO()
        figure.savefig(
            buffer, format="png", dpi=400,
            transparent=True, facecolor="none",
            bbox_inches="tight", pad_inches=0.06,
        )
    finally:
        plt.close(figure)
    return buffer.getvalue()


def _fallback_fontsize(height_in, count, expression, inline=False):
    """Point size for the fallback drawing, before the PNG is fitted to the box.

    The floor sits with 18–20pt Chinese beside the formula. Display fractions
    and roots are a bit larger so the numerator is not a size below the body.
    """
    per_line = 32 if inline else 30
    size = float(height_in) / max(1, count) * per_line
    if _TALL_MATH.search(expression or ""):
        size *= 1.2
    return max(22, min(48, size))


def render_equation_png(expressions, width_in, height_in, plain, strict=False, inline=False):
    """High-DPI PNG of the formula. ``strict`` refuses a plain-text stand-in.

    The font is stixsans so the fallback sits next to the sans body. When
    ``strict`` is set, a command mathtext cannot draw aborts the inject.
    Runs-only specs may still fall back to the OMML's visible text.
    """
    width_in = max(1.2, min(12.0, float(width_in)))
    height_in = max(0.45, min(3.5, float(height_in)))
    try:
        import matplotlib
        matplotlib.use("Agg", force=False)
        matplotlib.rcParams["mathtext.fontset"] = "stixsans"
    except ImportError as exc:
        raise ValueError(
            "matplotlib is required to render the OMML fallback image"
        ) from exc

    def draw_math(axis, height):
        count = max(1, len(expressions))
        for index, expr in enumerate(expressions):
            size = _fallback_fontsize(height, count, expr, inline=inline)
            axis.text(
                0.03, 1 - (index + 0.5) / count, expr,
                fontsize=size, color="#222222", va="center", ha="left",
            )

    if strict:
        from matplotlib import mathtext
        parser = mathtext.MathTextParser("agg")
        for expr in expressions:
            try:
                parser.parse(expr, dpi=72)
            except Exception as exc:
                raise ValueError(
                    "matplotlib mathtext could not render this formula fallback. "
                    "Refusing to place LaTeX source on the slide."
                ) from exc
    try:
        return _png_figure(draw_math, width_in, height_in)
    except Exception as exc:
        if strict:
            raise ValueError(
                "matplotlib mathtext could not render this formula fallback. "
                "Refusing to place LaTeX source on the slide."
            ) from exc
        def draw_plain(axis, height):
            axis.text(
                0.03, 0.5, plain or "?",
                fontsize=max(20, min(40, height * 24)),
                color="#222222", va="center", ha="left", fontfamily="DejaVu Sans",
            )
        return _png_figure(draw_plain, width_in, height_in)


def _a14_math(omath_xml):
    return (
        '<a14:m xmlns:a14="' + A14_NS + '">'
        '<m:oMathPara xmlns:m="' + MATH_NS + '">'
        '<m:oMathParaPr><m:jc m:val="left"/></m:oMathParaPr>'
        + omath_xml
        + "</m:oMathPara></a14:m>"
    )


def _shape_spans(xml):
    spans = []
    stack = []
    for match in TAG_RE.finditer(xml):
        closing, name, rest = match.group(1), match.group(2), match.group(3)
        if name != "p:sp":
            continue
        if closing:
            if stack:
                spans.append((stack.pop(), match.end()))
        elif not rest.rstrip().endswith("/"):
            stack.append(match.start())
    return spans


def _attrs(fragment, tag):
    match = re.search(r"<" + tag + r"\b([^>/]*)/?>", fragment)
    if not match:
        return {}
    return dict(re.findall(r'([A-Za-z0-9:]+)="([^"]*)"', match.group(1)))


def _xfrm_box(shape):
    match = re.search(r"<a:xfrm\b[^>]*>.*?</a:xfrm>", shape, re.DOTALL)
    if not match:
        return (
            '<a:xfrm><a:off x="640080" y="1600200"/><a:ext cx="5486400" cy="731520"/></a:xfrm>',
            6.0,
            0.8,
        )
    xfrm = match.group(0)
    ext = _attrs(xfrm, "a:ext")
    try:
        width = int(ext.get("cx", "5486400")) / EMU_IN
        height = int(ext.get("cy", "731520")) / EMU_IN
    except ValueError:
        width, height = 6.0, 0.8
    return xfrm, width, height


def _png_size(blob):
    if len(blob) < 24 or blob[:8] != b"\x89PNG\r\n\x1a\n":
        raise ValueError("Equation fallback is not a PNG")
    width, height = struct.unpack(">II", blob[16:24])
    if width <= 0 or height <= 0:
        raise ValueError("Equation fallback PNG has no pixels")
    return width, height


def _set_xml_attr(fragment, tag, name, value):
    pattern = r"(<" + tag + r"\b[^>]*\b" + name + r'=")[^"]*"'
    updated, count = re.subn(
        pattern, lambda match: match.group(1) + str(int(value)) + '"', fragment, count=1
    )
    if count != 1:
        raise ValueError("Missing " + tag + " @" + name)
    return updated


def _fit_xfrm(xfrm, png_w, png_h):
    """Fit the fallback box inside the placeholder at the PNG's aspect ratio.

    The picture is left-aligned and vertically centered. It is not stretched
    to the placeholder.
    """
    ext = _attrs(xfrm, "a:ext")
    off = _attrs(xfrm, "a:off")
    try:
        cx = int(ext.get("cx", "0"))
        cy = int(ext.get("cy", "0"))
        ox = int(off.get("x", "0"))
        oy = int(off.get("y", "0"))
    except ValueError:
        return xfrm
    if cx <= 0 or cy <= 0 or png_w <= 0 or png_h <= 0:
        return xfrm
    if png_w * cy >= png_h * cx:
        new_cx = cx
        new_cy = max(1, int(round(cx * png_h / float(png_w))))
        if new_cy > cy:
            new_cy = cy
    else:
        new_cy = cy
        new_cx = max(1, int(round(cy * png_w / float(png_h))))
        if new_cx > cx:
            new_cx = cx
    new_x = ox
    new_y = oy + max(0, (cy - new_cy) // 2)
    fitted = _set_xml_attr(xfrm, "a:ext", "cx", new_cx)
    fitted = _set_xml_attr(fitted, "a:ext", "cy", new_cy)
    fitted = _set_xml_attr(fitted, "a:off", "x", new_x)
    fitted = _set_xml_attr(fitted, "a:off", "y", new_y)
    return fitted


def _fallback_shape(xfrm, rel_id, shape_id):
    return (
        '<p:sp><p:nvSpPr><p:cNvPr id="%d" name="Equation fallback"/>'
        "<p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>%s"
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '<a:blipFill><a:blip r:embed="%s"/></a:blipFill>'
        "<a:ln><a:noFill/></a:ln>"
        "</p:spPr></p:sp>"
    ) % (shape_id, xfrm, rel_id)


def _wrap_alternate(choice_shape, fallback_shape):
    return (
        "<mc:AlternateContent>"
        '<mc:Choice Requires="a14">' + choice_shape + "</mc:Choice>"
        "<mc:Fallback>" + fallback_shape + "</mc:Fallback>"
        "</mc:AlternateContent>"
    )


def _ensure_namespaces(xml):
    match = re.search(r"<p:sld\b([^>]*)>", xml)
    if not match:
        raise ValueError("Slide XML has no p:sld root")
    attrs = match.group(1)
    prefix = ""
    if "xmlns:mc=" not in attrs:
        prefix += ' xmlns:mc="' + MC_NS + '"'
    if "xmlns:a14=" not in attrs:
        prefix += ' xmlns:a14="' + A14_NS + '"'
    if re.search(r"\bmc:Ignorable=", attrs):
        def add(found):
            parts = found.group(1).split()
            if "a14" not in parts:
                parts.append("a14")
            return 'mc:Ignorable="' + " ".join(parts) + '"'
        attrs = re.sub(r'\bmc:Ignorable="([^"]*)"', add, attrs, count=1)
    else:
        attrs += ' mc:Ignorable="a14"'
    return xml[:match.start()] + "<p:sld" + prefix + attrs + ">" + xml[match.end():]


def _fresh_rid(rels_text):
    used = set(re.findall(r'\bId="([^"]+)"', rels_text or ""))
    number = 1
    while ("rId%d" % number) in used:
        number += 1
    return "rId%d" % number


def _add_image_rel(rels_text, rel_id, target):
    if not rels_text:
        rels_text = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            "</Relationships>"
        )
    if ('Id="%s"' % rel_id) in rels_text:
        return rels_text
    relation = (
        '<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="%s"/>'
        % (rel_id, target)
    )
    if "</Relationships>" not in rels_text:
        raise ValueError("Slide relationships part is missing </Relationships>")
    return rels_text.replace("</Relationships>", relation + "</Relationships>", 1)


def _next_shape_id(xml):
    found = [int(value) for value in re.findall(r'\bid="(\d+)"', xml)]
    return (max(found) if found else 1) + 1


def _load_tex_fallback():
    """Load the optional TeX picture renderer. It does not touch the OMML compiler."""
    name = "tex_fallback"
    loaded = sys.modules.get(name)
    if loaded is not None and hasattr(loaded, "select_degradation"):
        return loaded
    path = Path(__file__).resolve().parent / "tex_fallback.py"
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


def _spec_latex(spec):
    if isinstance(spec, dict) and isinstance(spec.get("latex"), str) and spec["latex"].strip():
        return spec["latex"].strip()
    return None


def _note_tex_degrade(formula_ids, exc):
    detail = str(exc).splitlines()[0] if str(exc) else type(exc).__name__
    sys.stderr.write(
        "TeX fallback unavailable for %s; using mathtext. %s\n"
        % (",".join(formula_ids), detail)
    )


def _write_requested_sidecars(tex, formulas, used_ids):
    """Write ``{id}.tex`` when ``STANFORD_FORMULA_TEX_DIR`` is set. Never into the pptx."""
    folder = os.environ.get("STANFORD_FORMULA_TEX_DIR")
    if not folder:
        return
    root = Path(folder)
    root.mkdir(parents=True, exist_ok=True)
    for item in used_ids:
        spec = formulas.get(item)
        latex = _spec_latex(spec)
        if not latex:
            continue
        document = tex.build_latex_document(latex, display=not spec.get("inline"))
        tex.write_tex_sidecar(root / ("%s.tex" % item), document)


def _fallback_png(formulas, used_ids, expressions, width, height, plain, strict, inline):
    """Picture rung: TeX PNG when the toolchain and a LaTeX source exist, else mathtext.

    OMML has already been compiled. Pandoc is not involved here. A TeX failure
    degrades to mathtext instead of failing the page.
    """
    tex = _load_tex_fallback()
    sources = []
    has_latex = True
    for item in used_ids:
        spec = formulas.get(item)
        latex = _spec_latex(spec)
        if not latex:
            has_latex = False
            break
        sources.append((latex, not spec.get("inline")))
    status = tex.toolchain_status()
    plan = tex.select_degradation(
        compiler_ok=True,
        pandoc_present=shutil.which("pandoc") is not None,
        has_latex=has_latex,
        engine_present=status["engine"] is not None,
        png_converter_present=status["png_converter"] is not None,
        svg_converter_present=status["svg_converter"] is not None,
        mathtext_present=True,
    )
    _write_requested_sidecars(tex, formulas, used_ids)
    if plan["picture"] == "tex-png":
        try:
            blobs = [
                tex.render_tex_png(latex, display=display, dpi=tex.DEFAULT_DPI)
                for latex, display in sources
            ]
            return tex.stack_pngs(blobs)
        except (tex.TexFallbackError, OSError, subprocess.TimeoutExpired) as exc:
            _note_tex_degrade(used_ids, exc)
    return render_equation_png(
        expressions,
        width,
        max(height, 0.45 * len(expressions)),
        plain,
        strict=strict,
        inline=inline,
    )


def transform_slide(xml_text, formulas, rels_text, media_stem):
    """Return slide XML, slide rels XML, and ``{part_name: png_bytes}``."""
    cache = {}
    xml = xml_text
    rels = rels_text or ""
    parts = {}
    spans = [span for span in _shape_spans(xml) if "{{MATH:" in xml[span[0]:span[1]]]
    if not spans:
        raise ValueError("{{MATH: placeholder is not inside a text shape")
    for start, end in reversed(spans):
        shape = xml[start:end]
        used_ids = []

        def repl(match, bucket=used_ids):
            run = match.group(0)
            visible = "".join(TEXT_NODE.findall(run))
            found = PLACEHOLDER.findall(visible)
            if not found:
                return run
            if len(found) != 1 or visible.strip() != "{{MATH:%s}}" % found[0]:
                raise ValueError("Placeholder must be the entire run, found: " + visible[:80])
            formula_id = found[0]
            if formula_id not in formulas:
                raise ValueError("Unknown formula id: " + formula_id)
            if formula_id not in cache:
                cache[formula_id] = omml_for_spec(formulas[formula_id])
            bucket.append(formula_id)
            return _a14_math(cache[formula_id])

        choice = RUN_BLOCK.sub(repl, shape)
        if not used_ids or PLACEHOLDER.search(choice):
            raise ValueError("Unresolved {{MATH: placeholder")
        expressions = [mathtext_for_spec(formulas[item], cache[item]) for item in used_ids]
        plain = "  ".join(_omml_plain(cache[item]) for item in used_ids)
        xfrm, width, height = _xfrm_box(shape)
        strict = any(
            isinstance(formulas.get(item), dict)
            and isinstance(formulas[item].get("latex"), str)
            and formulas[item]["latex"].strip()
            for item in used_ids
        )
        inline = any(
            isinstance(formulas.get(item), dict) and formulas[item].get("inline")
            for item in used_ids
        )
        png = _fallback_png(
            formulas, used_ids, expressions, width, height, plain, strict, inline,
        )
        png_w, png_h = _png_size(png)
        xfrm = _fit_xfrm(xfrm, png_w, png_h)
        rel_id = _fresh_rid(rels)
        slug = "-".join(used_ids)
        part = "ppt/media/%s-%s.png" % (media_stem, slug)
        target = "../media/%s-%s.png" % (media_stem, slug)
        parts[part] = png
        rels = _add_image_rel(rels, rel_id, target)
        fallback = _fallback_shape(xfrm, rel_id, _next_shape_id(xml))
        xml = xml[:start] + _wrap_alternate(choice, fallback) + xml[end:]
    xml = _ensure_namespaces(xml)
    ET.fromstring(xml)
    if RAW_LATEX.search(xml):
        raise ValueError("Refusing to write LaTeX source into the slide")
    return xml, rels, parts


def _ensure_png_content_type(xml_text):
    if 'Extension="png"' in xml_text:
        return xml_text
    if "</Types>" not in xml_text:
        raise ValueError("[Content_Types].xml has no Types element")
    return xml_text.replace(
        "</Types>",
        '<Default Extension="png" ContentType="image/png"/></Types>',
        1,
    )


def inject_pptx(src, formulas, dest):
    """Write a new pptx with placeholders replaced. Does not modify src."""
    source = Path(src).resolve()
    target = Path(dest).resolve()
    if not source.is_file():
        raise ValueError("Missing pptx: " + str(source))
    if target == source:
        raise ValueError("Refusing to overwrite the source pptx")
    if target.exists():
        raise ValueError("Refusing to overwrite destination: " + str(target))
    if not isinstance(formulas, dict):
        raise ValueError("formulas must be a dict")

    with zipfile.ZipFile(source) as package:
        broken = package.testzip()
        if broken:
            raise ValueError("Corrupt ZIP member: " + broken)
        order = [info.filename for info in package.infolist()]
        files = {name: package.read(name) for name in order}
    for name in list(order):
        if not (name.startswith("ppt/slides/slide") and name.endswith(".xml") and "/_rels/" not in name):
            continue
        text = files[name].decode("utf-8")
        if "{{MATH:" not in text:
            continue
        rels_name = posixpath.join(posixpath.dirname(name), "_rels", posixpath.basename(name) + ".rels")
        rels_text = files[rels_name].decode("utf-8") if rels_name in files else ""
        stem = Path(name).stem
        text, rels_text, parts = transform_slide(text, formulas, rels_text, "omml-" + stem)
        files[name] = text.encode("utf-8")
        files[rels_name] = rels_text.encode("utf-8")
        if rels_name not in order:
            order.append(rels_name)
        for part, blob in parts.items():
            files[part] = blob
            if part not in order:
                order.append(part)
        if "[Content_Types].xml" in files:
            files["[Content_Types].xml"] = _ensure_png_content_type(
                files["[Content_Types].xml"].decode("utf-8")
            ).encode("utf-8")
    target.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(target, "w") as package:
        for name in order:
            package.writestr(name, files[name])
    return target


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pptx", help="Source pptx containing {{MATH:id}} runs")
    parser.add_argument("formulas", help="formulas.json")
    parser.add_argument("-o", "--output", required=True, help="New pptx path; never overwrites")
    args = parser.parse_args(argv)
    try:
        inject_pptx(args.pptx, load_formulas(args.formulas), args.output)
    except (OSError, ValueError, ET.ParseError, zipfile.BadZipFile, json.JSONDecodeError) as exc:
        print(type(exc).__name__ + ": " + str(exc))
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
