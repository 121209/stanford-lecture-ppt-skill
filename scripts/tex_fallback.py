#!/usr/bin/env python3
"""Optional TeX raster for the OMML picture fallback.

The editable equation stays OMML from the vendored compiler. This module only
draws the ``mc:Fallback`` image when a real TeX engine and a PNG converter are
both on ``PATH``. It follows the nature-image2ppt formula renderer: engine
order ``xelatex`` / ``lualatex`` / ``pdflatex``, a ``standalone`` document with
``amsmath,amssymb,mathtools,bm`` and ``xcolor``, display ``\\[...\\]`` or inline
``$...$``, then PDF to a high-DPI PNG or an SVG sidecar.

It does not adopt that renderer's product rules. A missing engine is not a
page failure, the picture is never marked ``editable: false``, and SVG is not
the LibreOffice-visible fallback. ``shell-escape`` stays off.
"""
from __future__ import annotations

import io
import os
import shutil
import subprocess
from pathlib import Path

DEFAULT_ENGINE_CANDIDATES = ("xelatex", "lualatex", "pdflatex")
DEFAULT_TIMEOUT = 120
MIN_DPI = 300
DEFAULT_DPI = 400
# First match wins for the picture. Pandoc is not a picture rung.
PICTURE_RUNGS = ("tex-png", "mathtext")
# The vendored compiler is always attempted. Pandoc runs only if it raises.
OMML_RUNGS = ("vendor", "pandoc")
INK = (0x22, 0x22, 0x22)
# Shared invisible box so separately cropped inline symbols keep one baseline.
INLINE_STRUT = r"\vphantom{\mathrm{softmax}d_{k}}"


class TexFallbackError(RuntimeError):
    pass


def select_degradation(
    *,
    compiler_ok=True,
    pandoc_present=False,
    has_latex=False,
    engine_present=False,
    png_converter_present=False,
    svg_converter_present=False,
    mathtext_present=True,
):
    """Choose rungs from probes. Does not spawn TeX, pandoc, or matplotlib.

    ``omml_attempt`` is always the vendored compiler. ``omml`` is the rung that
    can succeed given these probes: ``vendor``, else ``pandoc``, else ``fail``.
    ``picture`` is ``tex-png`` only when the spec has LaTeX and both an engine
    and a PNG converter exist; otherwise ``mathtext`` when that rung is
    available. ``svg_sidecar`` never replaces the picture inside the pptx.
    """
    if compiler_ok:
        omml = "vendor"
    elif pandoc_present:
        omml = "pandoc"
    else:
        omml = "fail"
    if has_latex and engine_present and png_converter_present:
        picture = "tex-png"
    elif mathtext_present:
        picture = "mathtext"
    else:
        picture = "fail"
    if has_latex and engine_present and svg_converter_present:
        svg = "tex-svg"
    else:
        svg = None
    return {
        "omml_attempt": "vendor",
        "omml": omml,
        "omml_if_compiler_raises": "pandoc" if pandoc_present else "fail",
        "picture": picture,
        "svg_sidecar": svg,
    }


def find_engine(engine=None):
    """Return ``(name, path)`` or ``None``. Does not raise when nothing is installed."""
    if engine and engine != "auto":
        resolved = shutil.which(engine)
        if not resolved:
            return None
        return Path(resolved).name, resolved
    for candidate in DEFAULT_ENGINE_CANDIDATES:
        resolved = shutil.which(candidate)
        if resolved:
            return candidate, resolved
    return None


def find_png_converter():
    """First available PNG tool: pdftocairo, ghostscript, then ImageMagick."""
    if shutil.which("pdftocairo"):
        return "pdftocairo"
    if shutil.which("gs"):
        return "ghostscript"
    if shutil.which("magick"):
        return "magick"
    if shutil.which("convert"):
        return "convert"
    return None


def find_svg_converter():
    """Optional sidecar converter. Not used as the pptx fallback image."""
    if shutil.which("dvisvgm"):
        return "dvisvgm"
    if shutil.which("pdf2svg"):
        return "pdf2svg"
    return None


def toolchain_status(engine=None):
    """Report TeX tools. ``STANFORD_FORMULA_DISABLE_TEX=1`` forces the mathtext rung."""
    if os.environ.get("STANFORD_FORMULA_DISABLE_TEX") == "1":
        return {
            "engine": None,
            "engine_path": None,
            "png_converter": None,
            "svg_converter": None,
            "tex_png": False,
            "tex_svg": False,
        }
    found = find_engine(engine)
    png = find_png_converter()
    svg = find_svg_converter()
    return {
        "engine": None if found is None else found[0],
        "engine_path": None if found is None else found[1],
        "png_converter": png,
        "svg_converter": svg,
        "tex_png": found is not None and png is not None,
        "tex_svg": found is not None and svg is not None,
    }


def require_engine(engine=None):
    found = find_engine(engine)
    if found is None:
        if engine and engine != "auto":
            raise TexFallbackError("LaTeX engine not found: " + engine)
        raise TexFallbackError(
            "No LaTeX engine found. Install xelatex, lualatex, or pdflatex "
            "to use the TeX picture fallback."
        )
    return found


def latex_command(engine, filename, shell_escape=False):
    command = [
        engine,
        "-interaction=nonstopmode",
        "-halt-on-error",
        "-file-line-error",
    ]
    if shell_escape:
        command.append("-shell-escape")
    command.append(filename)
    return command


def imagemagick_png_command(magick, pdf, out, dpi):
    """Nature's PNG invocation, with ``PNG32:`` so the alpha channel can survive."""
    return [
        str(magick),
        "-density",
        str(dpi),
        str(pdf),
        "-trim",
        "+repage",
        "PNG32:%s" % out,
    ]


def dvisvgm_svg_command(dvisvgm, pdf, out):
    return [
        str(dvisvgm),
        "--pdf",
        "--no-fonts",
        "--exact",
        "--output",
        str(out),
        str(pdf),
    ]


def _strip_math_delimiters(tex):
    text = tex.strip()
    if text.startswith("$$") and text.endswith("$$") and len(text) >= 4:
        return text[2:-2].strip()
    if (
        len(text) >= 2
        and text.startswith("$")
        and text.endswith("$")
        and text.count("$") == 2
    ):
        return text[1:-1].strip()
    return text


def build_latex_document(
    tex, *, preamble="", full_document=False, display=True, font_size=None
):
    """Standalone wrapper. Ink is ``#222222`` so the fallback matches the slide body.

    Nature writes ``\\[...\\]``. On this TeX Live, standalone crops in a
    horizontal box, so ``\\[`` never enters display math. ``varwidth`` restores
    ``\\[`` but then clips anything wider than ``\\textwidth``. Display style
    is therefore ``$\\displaystyle ...$``, which is the same style without the clip.
    """
    if not isinstance(tex, str) or not tex.strip():
        raise TexFallbackError("LaTeX input is empty.")
    if full_document:
        return tex if tex.endswith("\n") else tex + "\n"
    body = _strip_math_delimiters(tex)
    if not body:
        raise TexFallbackError("LaTeX input is empty.")
    if display:
        wrapped = "$\\displaystyle\n" + body + "\n$"
    else:
        wrapped = "$" + INLINE_STRUT + " " + body + "$"
    extra = preamble.strip()
    extra_line = (extra + "\n") if extra else ""
    font_line = ""
    if font_size is not None:
        size = float(font_size)
        if size < 10 or size > 72:
            raise TexFallbackError("TeX font size must be between 10 and 72 pt")
        font_line = "\\fontsize{%.2f}{%.2f}\\selectfont\n" % (size, size * 1.2)
    return (
        "\\documentclass[border=2pt]{standalone}\n"
        "\\usepackage{lmodern}\n"
        "\\usepackage{amsmath,amssymb,mathtools,bm}\n"
        "\\usepackage{xcolor}\n"
        + extra_line
        + "\\begin{document}\n"
        + font_line
        + "\\color[HTML]{222222}\n"
        + wrapped
        + "\n\\end{document}\n"
    )


def write_tex_sidecar(path, document):
    dest = Path(path)
    dest.parent.mkdir(parents=True, exist_ok=True)
    if not document.endswith("\n"):
        document += "\n"
    dest.write_text(document, encoding="utf-8")
    return dest


def _run(command, timeout, cwd=None):
    try:
        result = subprocess.run(
            [str(part) for part in command],
            cwd=cwd,
            text=True,
            capture_output=True,
            timeout=timeout,
        )
    except subprocess.TimeoutExpired as exc:
        raise TexFallbackError(
            "Formula command timed out: " + " ".join(str(part) for part in command)
        ) from exc
    except OSError as exc:
        raise TexFallbackError("Formula command failed to start: " + str(exc)) from exc
    if result.returncode != 0:
        log = result.stdout or result.stderr or ""
        tail = "\n".join(log.splitlines()[-30:])
        raise TexFallbackError(
            "Formula command failed (%s): %s\n%s"
            % (result.returncode, " ".join(str(part) for part in command), tail)
        )
    return result


def _compile_pdf(document, engine, timeout, shell_escape):
    name, resolved = require_engine(engine)
    import tempfile

    folder = tempfile.TemporaryDirectory()
    work = Path(folder.name)
    (work / "formula.tex").write_text(document, encoding="utf-8")
    try:
        _run(
            latex_command(resolved, "formula.tex", shell_escape=shell_escape),
            timeout,
            cwd=work,
        )
    except TexFallbackError:
        folder.cleanup()
        raise
    pdf = work / "formula.pdf"
    if not pdf.is_file():
        folder.cleanup()
        raise TexFallbackError("LaTeX completed but formula.pdf was not produced.")
    return folder, pdf, name


def _assert_dpi(dpi):
    if int(dpi) < MIN_DPI:
        raise TexFallbackError("PNG dpi must be >= %d, got %s" % (MIN_DPI, dpi))
    return int(dpi)


def _read_png(path, label):
    if not path.is_file():
        raise TexFallbackError(label + " did not write a PNG")
    return path.read_bytes()


def convert_pdf_to_png(pdf, dpi, timeout):
    """Return ``(png_bytes, converter_name)`` before the white-paper knockout."""
    dpi = _assert_dpi(dpi)
    pdf = Path(pdf)
    out = pdf.with_name("formula-out.png")
    errors = []

    pdftocairo = shutil.which("pdftocairo")
    if pdftocairo:
        try:
            if out.exists():
                out.unlink()
            prefix = str(out.with_suffix(""))
            _run(
                [pdftocairo, "-png", "-transp", "-singlefile", "-r", str(dpi), str(pdf), prefix],
                timeout,
            )
            return _read_png(out, "pdftocairo"), "pdftocairo"
        except TexFallbackError as exc:
            errors.append("pdftocairo: " + str(exc).splitlines()[0])

    gs = shutil.which("gs")
    if gs:
        try:
            if out.exists():
                out.unlink()
            _run(
                [
                    gs,
                    "-dSAFER",
                    "-dBATCH",
                    "-dNOPAUSE",
                    "-q",
                    "-dTextAlphaBits=4",
                    "-dGraphicsAlphaBits=4",
                    "-sDEVICE=pngalpha",
                    "-r%d" % dpi,
                    "-sOutputFile=%s" % out,
                    str(pdf),
                ],
                timeout,
            )
            return _read_png(out, "ghostscript"), "ghostscript"
        except TexFallbackError as exc:
            errors.append("ghostscript: " + str(exc).splitlines()[0])

    magick = shutil.which("magick") or shutil.which("convert")
    if magick:
        try:
            if out.exists():
                out.unlink()
            _run(imagemagick_png_command(magick, pdf, out, dpi), timeout)
            return _read_png(out, "ImageMagick"), Path(magick).name
        except TexFallbackError as exc:
            errors.append("ImageMagick: " + str(exc).splitlines()[0])

    if errors:
        raise TexFallbackError("PNG conversion failed. " + " | ".join(errors))
    raise TexFallbackError(
        "PNG output requires pdftocairo, ghostscript (gs), or ImageMagick (magick or convert)."
    )


def convert_pdf_to_svg(pdf, out_path, timeout):
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    dvisvgm = shutil.which("dvisvgm")
    if dvisvgm:
        _run(dvisvgm_svg_command(dvisvgm, pdf, out_path), timeout)
        if not out_path.is_file():
            raise TexFallbackError("dvisvgm did not write an SVG")
        return "dvisvgm"
    pdf2svg = shutil.which("pdf2svg")
    if pdf2svg:
        _run([pdf2svg, str(pdf), str(out_path)], timeout)
        if not out_path.is_file():
            raise TexFallbackError("pdf2svg did not write an SVG")
        return "pdf2svg"
    raise TexFallbackError("SVG output requires dvisvgm or pdf2svg.")


def _pixel_rows(image):
    if hasattr(image, "get_flattened_data"):
        return list(image.get_flattened_data())
    return list(image.getdata())


def _corners_clear(image):
    width, height = image.size
    if width < 2 or height < 2:
        return False
    pixels = image.load()
    return all(
        pixels[x, y][3] < 8
        for x, y in ((0, 0), (width - 1, 0), (0, height - 1), (width - 1, height - 1))
    )


def _knockout_white_paper(image):
    """Opaque white scans become transparent ink. Already-clear renders skip this."""
    pixels = []
    for red, green, blue, alpha in _pixel_rows(image):
        if alpha < 8 or (red >= 248 and green >= 248 and blue >= 248):
            pixels.append((0, 0, 0, 0))
            continue
        if alpha < 250 and max(red, green, blue) < 40:
            pixels.append((INK[0], INK[1], INK[2], alpha))
            continue
        if alpha >= 250 and max(red, green, blue) < 80:
            pixels.append((red, green, blue, alpha))
            continue
        lum = (0.2126 * red) + (0.7152 * green) + (0.0722 * blue)
        ink = int(round(255 - lum))
        if alpha < 250:
            ink = int(round(ink * (alpha / 255.0)))
        if ink < 8:
            pixels.append((0, 0, 0, 0))
        else:
            pixels.append((INK[0], INK[1], INK[2], min(255, ink)))
    image.putdata(pixels)
    return image


def _save_png(image, dpi):
    buffer = io.BytesIO()
    image.save(buffer, format="PNG", dpi=(int(dpi), int(dpi)))
    blob = buffer.getvalue()
    if len(blob) < 26 or blob[25] != 6:
        raise TexFallbackError("TeX PNG was not RGBA")
    return blob


def _transparent_png(raw, pad=8, dpi=DEFAULT_DPI, trim=True):
    """RGBA PNG. A clear render keeps its alpha; a white page is knocked out.

    Inline callers pass ``trim=False`` so a shared strut still reserves depth.
    Trimming would crop that invisible box and the baselines would drift.
    """
    try:
        from PIL import Image
    except ImportError as exc:
        raise TexFallbackError(
            "Pillow is required to make the TeX PNG transparent."
        ) from exc
    image = Image.open(io.BytesIO(raw)).convert("RGBA")
    if not _corners_clear(image):
        image = _knockout_white_paper(image)
    if trim:
        bbox = image.getbbox()
        if not bbox:
            raise TexFallbackError("TeX PNG had no ink")
        image = image.crop(bbox)
        if pad:
            canvas = Image.new(
                "RGBA",
                (image.width + pad * 2, image.height + pad * 2),
                (0, 0, 0, 0),
            )
            canvas.paste(image, (pad, pad), image)
            image = canvas
    elif image.getbbox() is None:
        raise TexFallbackError("TeX PNG had no ink")
    return _save_png(image, dpi)


def stack_pngs(blobs):
    """Vertical stack of RGBA PNGs. One blob is returned unchanged."""
    if not blobs:
        raise TexFallbackError("No PNG blobs to stack")
    if len(blobs) == 1:
        return blobs[0]
    try:
        from PIL import Image
    except ImportError as exc:
        raise TexFallbackError("Pillow is required to stack TeX PNGs.") from exc
    images = [Image.open(io.BytesIO(blob)).convert("RGBA") for blob in blobs]
    gap = 16
    width = max(image.width for image in images)
    height = sum(image.height for image in images) + gap * (len(images) - 1)
    canvas = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    y = 0
    for image in images:
        canvas.paste(image, (0, y), image)
        y += image.height + gap
    buffer = io.BytesIO()
    canvas.save(buffer, format="PNG")
    return buffer.getvalue()


def render_tex_png(
    tex,
    *,
    display=True,
    dpi=DEFAULT_DPI,
    engine=None,
    timeout=DEFAULT_TIMEOUT,
    preamble="",
    keep_tex=None,
    full_document=False,
    shell_escape=False,
    font_size=None,
    trim=True,
):
    """Render one formula to a transparent RGBA PNG at ``dpi`` (>= 300)."""
    dpi = _assert_dpi(dpi)
    document = build_latex_document(
        tex,
        preamble=preamble,
        full_document=full_document,
        display=display,
        font_size=font_size,
    )
    if keep_tex is not None:
        write_tex_sidecar(keep_tex, document)
    folder, pdf, _name = _compile_pdf(document, engine, timeout, shell_escape)
    try:
        raw, _converter = convert_pdf_to_png(pdf, dpi=dpi, timeout=timeout)
        return _transparent_png(raw, dpi=dpi, trim=trim)
    finally:
        folder.cleanup()


def render_tex_svg(
    tex,
    out_path,
    *,
    display=True,
    engine=None,
    timeout=DEFAULT_TIMEOUT,
    preamble="",
    keep_tex=None,
    full_document=False,
    shell_escape=False,
):
    """Write an SVG sidecar. Callers must not use it as the only pptx fallback."""
    document = build_latex_document(
        tex, preamble=preamble, full_document=full_document, display=display
    )
    if keep_tex is not None:
        write_tex_sidecar(keep_tex, document)
    folder, pdf, name = _compile_pdf(document, engine, timeout, shell_escape)
    try:
        converter = convert_pdf_to_svg(pdf, out_path, timeout)
    finally:
        folder.cleanup()
    return {
        "out": str(out_path),
        "format": "svg",
        "engine": name,
        "converter": converter,
        "role": "sidecar",
    }
