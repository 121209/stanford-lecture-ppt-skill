#!/usr/bin/env python3
"""Replace {{MATH:id}} runs with OMML. Does not author a deck by itself.

Stdlib only. Pandoc is optional and used only when a formula spec has ``latex``
and neither ``runs`` nor ``omml``. LaTeX source is never written onto a slide.
"""
import argparse
import json
import re
import shutil
import subprocess
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from xml.sax.saxutils import escape

# Keep aligned with scripts/audit_pptx.py RAW_LATEX.
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
NS = {"m": MATH_NS}


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
        omml = omml_from_pandoc(spec["latex"])
        if omml is None:
            raise ValueError(
                "pandoc is not available; use a high-DPI math image or provide runs/omml. "
                "Refusing to place LaTeX source on the slide."
            )
        return omml
    raise ValueError("Formula spec needs runs, omml, or latex")


def load_formulas(path):
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    if not isinstance(data, dict) or not data:
        raise ValueError("formulas file must be a non-empty object")
    for key in data:
        if not isinstance(key, str) or not ID_RE.fullmatch(key):
            raise ValueError("Bad formula id: " + str(key))
    return data


def replace_placeholders(xml_text, formulas):
    cache = {}

    def repl(match):
        run = match.group(0)
        visible = "".join(TEXT_NODE.findall(run))
        ids = PLACEHOLDER.findall(visible)
        if not ids:
            return run
        if len(ids) != 1 or visible.strip() != "{{MATH:%s}}" % ids[0]:
            raise ValueError("Placeholder must be the entire run, found: " + visible[:80])
        formula_id = ids[0]
        if formula_id not in formulas:
            raise ValueError("Unknown formula id: " + formula_id)
        if formula_id not in cache:
            cache[formula_id] = omml_for_spec(formulas[formula_id])
        return cache[formula_id]

    updated = RUN_BLOCK.sub(repl, xml_text)
    if PLACEHOLDER.search(updated):
        raise ValueError("Unresolved {{MATH: placeholder")
    return updated


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

    prepared = []
    with zipfile.ZipFile(source) as package:
        broken = package.testzip()
        if broken:
            raise ValueError("Corrupt ZIP member: " + broken)
        for info in package.infolist():
            data = package.read(info.filename)
            name = info.filename
            if name.startswith("ppt/slides/slide") and name.endswith(".xml") and "/_rels/" not in name:
                text = data.decode("utf-8")
                if "{{MATH:" in text:
                    text = replace_placeholders(text, formulas)
                    if RAW_LATEX.search(text):
                        raise ValueError("Refusing to write LaTeX source into " + name)
                    data = text.encode("utf-8")
            prepared.append((name, data))
    target.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(target, "w") as package:
        for name, data in prepared:
            package.writestr(name, data)
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
