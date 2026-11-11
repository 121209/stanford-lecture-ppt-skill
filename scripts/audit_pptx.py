#!/usr/bin/env python3
"""Read-only OOXML inventory. This is not a visual or scientific QA verdict.

Also estimates per-slide on-screen word counts. A word-cap warning is off unless
--word-cap is passed. It warns when the lower half of the content band is largely empty. A content slide
whose title names an algorithm, method, architecture, or pipeline, and that has no
diagram, picture, table, chart, or 「缺图」 marker, is a warning. So is raw LaTeX,
and OMML that is not PowerPoint's a14 choice plus a picture fallback. Visible body
text below 18pt outside the footer/source band, and text that is faint on white
(including #BFBFBF), are warnings too. None of these warnings change the process
exit code.
"""
import argparse
import hashlib
import json
import posixpath
import re
from pathlib import Path
import sys
import xml.etree.ElementTree as ET
import zipfile

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "a14": "http://schemas.microsoft.com/office/drawing/2010/main",
}
EMU_PER_PX = 9525
# Explicit run size is in hundredths of a point. 15pt source lines are excluded.
SOURCE_LINE_MAX_HUNDREDTHS_PT = 1500
# Body hard floor. Caption-sized 16pt text is still below this and may warn.
BODY_MIN_HUNDREDTHS_PT = 1800
# WCAG contrast of explicit text color against the white slide.
CONTRAST_MIN_ON_WHITE = 4.5
FORBIDDEN_TEXT_COLOR = "BFBFBF"
# Bottom band of the slide: ≤15pt paragraphs here are treated as source lines.
FOOTER_SOURCE_MIN_Y_RATIO = 0.85
# Content band: body below the title (~y145/720) down to the footer source band.
CONTENT_BAND_TOP_RATIO = 0.20
# A shape counts in a half only when it is at least this large and overlaps that half
# by this much. ~28px is one text line; a hairline or a few pixels past the midline
# does not fill the lower half.
CONTENT_BAND_MIN_EXTENT_PX = 28
CONTENT_BAND_MIN_EXTENT_EMU = int(round(CONTENT_BAND_MIN_EXTENT_PX * EMU_PER_PX))
CJK_CHAR = re.compile(r"[\u3400-\u9fff]")
LATIN_WORD = re.compile(r"[A-Za-z0-9]+(?:[-'’./][A-Za-z0-9]+)*")
PAGE_NUMBER_TEXT = re.compile(r"\d{1,3}")
SOURCE_MARK = re.compile(
    r"(来源\s*[:：]|出处\s*[:：]|图源\s*[:：]|资料来源|引自|Source\s*:|DOI\b|doi\s*:|arXiv|https?://)",
    re.IGNORECASE,
)
# Title text that should carry a native diagram, a picture, a table, a chart, or 「缺图」.
# Match concepts, not model names: a title may say Mamba or Transformer without this firing.
# Bare 「步骤」 stays out so a discussion title is not flagged. 「方法」 is included because
# matching is limited to the title.
ALGORITHM_INTENT = re.compile(
    r"算法|伪代码|方法|训练流程|推理流程|推理步骤|解题流程|流程图|流程|"
    r"架构|框架|模型结构|系统模块|"
    r"\b(?:encoder|decoder|pipeline|framework|architecture|algorithm|method)s?\b",
    re.IGNORECASE,
)
TITLE_PLACEHOLDERS = {"title", "ctrTitle"}
GAP_MARKER = re.compile(r"缺图")
# On-screen LaTeX source. Unicode λ is allowed; a backslash command is not.
# Keep aligned with scripts/inject_omml.py RAW_LATEX.
RAW_LATEX = re.compile(
    r"\\(?:frac|dfrac|tfrac|sum|prod|int|iint|iiint|lambda|alpha|beta|gamma|theta|sigma|"
    r"mu|nu|omega|phi|psi|delta|epsilon|left|right|cdot|times|sqrt|mathbb|mathrm|mathbf|"
    r"text|begin|end|overline|hat|bar|vec|partial|infty|leq|geq|neq|approx|equiv|"
    r"subset|rightarrow|leftarrow)\b"
    r"|\\[{}]"
    r"|\^\{"
    r"|_\{"
)


def tag(prefix, local):
    return "{" + NS[prefix] + "}" + local


def resolve_part(owner, target):
    resolved = (target.lstrip("/") if target.startswith("/") else
                posixpath.normpath(posixpath.join(posixpath.dirname(owner), target)))
    if resolved.startswith("../") or resolved == "..":
        raise ValueError("Relationship escapes package root: " + target)
    return resolved


def relationships(zf, owner):
    relpart = posixpath.join(posixpath.dirname(owner), "_rels", posixpath.basename(owner) + ".rels")
    if relpart not in zf.namelist():
        return {}
    return {r.get("Id"): dict(r.attrib) for r in ET.fromstring(zf.read(relpart))}


def objects(container, nested=False):
    """Choose first AlternateContent Choice for inventory, not a renderer claim."""
    for item in container:
        if item.tag == tag("mc", "AlternateContent"):
            choice = item.find("mc:Choice", NS)
            if choice is None:
                choice = item.find("mc:Fallback", NS)
            if choice is not None:
                yield from objects(choice, nested)
        elif item.tag == tag("p", "grpSp"):
            yield from objects(item, True)
        elif item.tag in (tag("p", "sp"), tag("p", "pic"), tag("p", "graphicFrame"), tag("p", "cxnSp")):
            yield item, nested


def text_of(node):
    return "\n".join("".join(t.itertext()) for t in node.iter()
                     if t.tag in (tag("a", "t"), tag("m", "t")))


def word_count(text):
    """CJK: 2 characters = 1 word. Other letters and numbers: one token = 1 word."""
    cjk = CJK_CHAR.findall(text)
    rest = CJK_CHAR.sub(" ", text)
    return len(cjk) / 2 + len(LATIN_WORD.findall(rest))


def format_word_count(value):
    rounded = round(value, 1)
    if rounded == int(rounded):
        return int(rounded)
    return rounded


def _run_size(run, default_sz):
    r_pr = run.find("a:rPr", NS)
    if r_pr is not None and r_pr.get("sz"):
        try:
            return int(r_pr.get("sz"))
        except ValueError:
            return default_sz
    return default_sz


def paragraph_pieces(paragraph):
    default_sz = None
    p_pr = paragraph.find("a:pPr", NS)
    if p_pr is not None:
        def_rpr = p_pr.find("a:defRPr", NS)
        if def_rpr is not None and def_rpr.get("sz"):
            try:
                default_sz = int(def_rpr.get("sz"))
            except ValueError:
                default_sz = None
    pieces = []
    for child in list(paragraph):
        if child.tag in (tag("a", "r"), tag("a", "fld")):
            text = "".join(node.text or "" for node in child.iter(tag("a", "t")))
            if text:
                pieces.append((text, _run_size(child, default_sz)))
        elif child.tag in (tag("m", "oMath"), tag("m", "oMathPara")):
            text = "".join(node.text or "" for node in child.iter(tag("m", "t")))
            if text:
                pieces.append((text, None))
    return pieces


def paragraph_words(paragraph, in_footer):
    """Count one paragraph. Drop page numbers and ≤15pt source lines; keep shrunk body text."""
    pieces = paragraph_pieces(paragraph)
    if not pieces:
        return 0.0
    visible = "".join(text for text, _ in pieces).strip()
    if not visible:
        return 0.0
    sizes = [sz for _, sz in pieces]
    small = all(sz is not None and sz <= SOURCE_LINE_MAX_HUNDREDTHS_PT for sz in sizes)
    if PAGE_NUMBER_TEXT.fullmatch(visible) and (small or in_footer or all(sz is None for sz in sizes)):
        return 0.0
    if small and (in_footer or SOURCE_MARK.search(visible)):
        return 0.0
    return word_count(visible)


def _srgb(element):
    if element is None:
        return None
    color = element.find("a:solidFill/a:srgbClr", NS)
    if color is None:
        return None
    value = (color.get("val") or "").strip().upper()
    if len(value) == 6 and all(char in "0123456789ABCDEF" for char in value):
        return value
    return None


def _channel_luminance(channel):
    unit = channel / 255.0
    if unit <= 0.04045:
        return unit / 12.92
    return ((unit + 0.055) / 1.055) ** 2.4


def contrast_on_white(hex_color):
    """WCAG contrast ratio of an sRGB hex color against pure white."""
    red = int(hex_color[0:2], 16)
    green = int(hex_color[2:4], 16)
    blue = int(hex_color[4:6], 16)
    luminance = (
        0.2126 * _channel_luminance(red)
        + 0.7152 * _channel_luminance(green)
        + 0.0722 * _channel_luminance(blue)
    )
    return 1.05 / (luminance + 0.05)


def _format_pt(value):
    rounded = round(value, 2)
    if rounded == int(rounded):
        return int(rounded)
    return rounded


def _paragraph_run_style(paragraph):
    """Visible drawing runs as (text, size hundredths or None, sRGB or None)."""
    default_sz = None
    default_color = None
    p_pr = paragraph.find("a:pPr", NS)
    if p_pr is not None:
        def_rpr = p_pr.find("a:defRPr", NS)
        if def_rpr is not None:
            if def_rpr.get("sz"):
                try:
                    default_sz = int(def_rpr.get("sz"))
                except ValueError:
                    default_sz = None
            default_color = _srgb(def_rpr)
    runs = []
    for child in list(paragraph):
        if child.tag not in (tag("a", "r"), tag("a", "fld")):
            continue
        text = "".join(node.text or "" for node in child.iter(tag("a", "t")))
        if not text:
            continue
        color = _srgb(child.find("a:rPr", NS)) or default_color
        runs.append((text, _run_size(child, default_sz), color))
    return runs


def _in_footer_band(box, slide_h, nested):
    return (
        (not nested)
        and box is not None
        and slide_h > 0
        and box[1] >= slide_h * FOOTER_SOURCE_MIN_Y_RATIO
    )


def typography_findings(shape, box, slide_h, nested):
    """Point sizes below 18pt, and faint explicit text colors, for one shape.

    Returns ``(sizes_pt, colors)``. Footer/source-band shapes, slide numbers,
    digit-only page numbers, and ≤15pt source lines are not body text.
    """
    ph = shape.find("p:nvSpPr/p:nvPr/p:ph", NS)
    if ph is not None and ph.get("type") == "sldNum":
        return [], []
    in_footer = _in_footer_band(box, slide_h, nested)
    small = []
    pale = []
    seen = set()
    for tx_name in (tag("p", "txBody"), tag("a", "txBody")):
        for body in shape.iter(tx_name):
            for paragraph in body.findall("a:p", NS):
                if id(paragraph) in seen:
                    continue
                seen.add(id(paragraph))
                runs = _paragraph_run_style(paragraph)
                if not runs:
                    continue
                visible = "".join(text for text, _, _ in runs).strip()
                if not visible:
                    continue
                sizes = [sz for _, sz, _ in runs]
                source_line = all(
                    sz is not None and sz <= SOURCE_LINE_MAX_HUNDREDTHS_PT for sz in sizes
                ) and bool(SOURCE_MARK.search(visible))
                page_number = bool(PAGE_NUMBER_TEXT.fullmatch(visible)) and (
                    in_footer
                    or all(sz is None or sz <= SOURCE_LINE_MAX_HUNDREDTHS_PT for sz in sizes)
                )
                skip_body_floor = in_footer or source_line or page_number
                for text, sz, color in runs:
                    if not text.strip():
                        continue
                    if (
                        not skip_body_floor
                        and sz is not None
                        and sz < BODY_MIN_HUNDREDTHS_PT
                    ):
                        small.append(sz / 100.0)
                    if color and (
                        color == FORBIDDEN_TEXT_COLOR
                        or contrast_on_white(color) < CONTRAST_MIN_ON_WHITE
                    ):
                        pale.append(color)
    return small, pale


def onscreen_words(shape, box, slide_h, nested):
    ph = shape.find("p:nvSpPr/p:nvPr/p:ph", NS)
    if ph is not None and ph.get("type") == "sldNum":
        return 0.0
    in_footer = (not nested) and box is not None and slide_h > 0 and box[1] >= slide_h * FOOTER_SOURCE_MIN_Y_RATIO
    total = 0.0
    seen = set()
    for tx_name in (tag("p", "txBody"), tag("a", "txBody")):
        for body in shape.iter(tx_name):
            for paragraph in body.findall("a:p", NS):
                if id(paragraph) in seen:
                    continue
                seen.add(id(paragraph))
                total += paragraph_words(paragraph, in_footer)
    return total


def geometry(node):
    xfrm = node.find("p:spPr/a:xfrm", NS)
    if xfrm is None:
        xfrm = node.find("p:xfrm", NS)
    if xfrm is None:
        return None
    off, ext = xfrm.find("a:off", NS), xfrm.find("a:ext", NS)
    if off is None or ext is None:
        return None
    return [int(off.get("x", 0)), int(off.get("y", 0)),
            int(ext.get("cx", 0)), int(ext.get("cy", 0))]


def group_geometry(group):
    xfrm = group.find("p:grpSpPr/a:xfrm", NS)
    if xfrm is None:
        return None
    off, ext = xfrm.find("a:off", NS), xfrm.find("a:ext", NS)
    if off is None or ext is None:
        return None
    return [int(off.get("x", 0)), int(off.get("y", 0)),
            int(ext.get("cx", 0)), int(ext.get("cy", 0))]


def _is_slide_number(shape):
    ph = shape.find("p:nvSpPr/p:nvPr/p:ph", NS)
    return ph is not None and ph.get("type") == "sldNum"


def is_diagram_object(shape):
    """Connector, picture, or a real autoshape. Text boxes and empty rectangles are not.

    Titles in this skill are often no-fill rectangles. Those must not satisfy an
    algorithm page. A rounded rectangle, diamond, ellipse, or filled rectangle does.
    """
    if shape.tag in (tag("p", "cxnSp"), tag("p", "pic")):
        return True
    if shape.tag != tag("p", "sp"):
        return False
    if _is_slide_number(shape):
        return False
    locks = shape.find("p:nvSpPr/p:cNvSpPr", NS)
    if locks is not None and locks.get("txBox") == "1":
        return False
    ph = shape.find("p:nvSpPr/p:nvPr/p:ph", NS)
    if ph is not None:
        return False
    geom = shape.find("p:spPr/a:prstGeom", NS)
    if geom is None:
        return False
    preset = geom.get("prst") or ""
    if preset == "rect":
        return shape.find("p:spPr/a:solidFill", NS) is not None
    return bool(preset)


def _occupies_content_band(shape):
    """Pictures, tables, charts, connectors, and groups count by geometry.

    A text box counts only when it has non-whitespace text, so an empty
    placeholder does not hide a blank lower half.
    """
    if shape.tag == tag("p", "sp"):
        return bool(text_of(shape).strip())
    return shape.tag in (tag("p", "pic"), tag("p", "graphicFrame"), tag("p", "cxnSp"), tag("p", "grpSp"))


def content_band_boxes(container):
    """Slide-coordinate boxes that can occupy the content band.

    Group children are not measured: their offsets are often local to the group.
    A group counts only through its own transform.
    """
    boxes = []
    for item in container:
        if item.tag == tag("mc", "AlternateContent"):
            choice = item.find("mc:Choice", NS)
            if choice is None:
                choice = item.find("mc:Fallback", NS)
            if choice is not None:
                boxes.extend(content_band_boxes(choice))
            continue
        if not _occupies_content_band(item):
            continue
        if item.tag == tag("p", "grpSp"):
            box = group_geometry(item)
        else:
            if _is_slide_number(item):
                continue
            box = geometry(item)
        if box is not None and box[2] > 0 and box[3] > 0:
            boxes.append(box)
    return boxes


def content_band_bounds(slide_h):
    top = int(slide_h * CONTENT_BAND_TOP_RATIO)
    bottom = int(slide_h * FOOTER_SOURCE_MIN_Y_RATIO)
    mid = (top + bottom) // 2
    return top, mid, bottom


def _overlap(start, length, y0, y1):
    return max(0, min(start + length, y1) - max(start, y0))


def _occupies_half(box, y0, y1, min_extent):
    if box[2] < min_extent or box[3] < min_extent:
        return False
    return _overlap(box[1], box[3], y0, y1) >= min_extent


def lower_content_band_is_empty(boxes, slide_h, min_extent=CONTENT_BAND_MIN_EXTENT_EMU):
    """True when the upper content band holds a shape and the lower half does not.

    Cover pages are not special-cased here; the caller exempts page 1.
    """
    top, mid, bottom = content_band_bounds(slide_h)
    if mid - top < min_extent or bottom - mid < min_extent:
        return False
    upper = False
    lower = False
    for box in boxes:
        if box is None:
            continue
        if _occupies_half(box, top, mid, min_extent):
            upper = True
        if _occupies_half(box, mid, bottom, min_extent):
            lower = True
        if upper and lower:
            return False
    return upper and not lower


def slide_title_text(tree, slide_h):
    """Title placeholder text, otherwise the topmost non-page-number text shape.

    Group children are skipped: their offsets are often local to the group.
    """
    placeholders = []
    topmost = None
    for shape, nested in objects(tree):
        if nested or shape.tag != tag("p", "sp") or _is_slide_number(shape):
            continue
        text = text_of(shape).strip()
        if not text or PAGE_NUMBER_TEXT.fullmatch(text):
            continue
        ph = shape.find("p:nvSpPr/p:nvPr/p:ph", NS)
        if ph is not None and ph.get("type") in TITLE_PLACEHOLDERS:
            placeholders.append(text)
            continue
        box = geometry(shape)
        y = box[1] if box else 0
        if topmost is None or y < topmost[0]:
            topmost = (y, text)
    if placeholders:
        return "\n".join(placeholders)
    if topmost:
        return topmost[1]
    return ""


def _parent_map(root):
    parents = {}
    for parent in root.iter():
        for child in list(parent):
            parents[child] = parent
    return parents


def _requires_a14(choice):
    required = (choice.get("Requires") or "").replace(",", " ").split()
    return "a14" in required


def omml_has_picture_fallback(omath, parents):
    """True when oMath sits in Choice/a14:m/oMathPara and Fallback has a picture."""
    node = omath
    saw_para = False
    saw_a14 = False
    choice = None
    alternate = None
    while node is not None:
        if node.tag == tag("m", "oMathPara"):
            saw_para = True
        elif node.tag == tag("a14", "m"):
            saw_a14 = True
        elif node.tag == tag("mc", "Choice") and _requires_a14(node):
            choice = node
        elif node.tag == tag("mc", "AlternateContent"):
            alternate = node
            break
        node = parents.get(node)
    if not (saw_para and saw_a14 and choice is not None and alternate is not None):
        return False
    fallback = next((child for child in list(alternate) if child.tag == tag("mc", "Fallback")), None)
    if fallback is None:
        return False
    return fallback.find(".//" + tag("a", "blip")) is not None


def office_math_fallback_warning(tree):
    parents = _parent_map(tree)
    missing = 0
    for omath in tree.iter(tag("m", "oMath")):
        if not omml_has_picture_fallback(omath, parents):
            missing += 1
    if not missing:
        return None
    return {"code": "office_math_missing_fallback", "count": missing}


def notes_body(zf, slide_part):
    rels = relationships(zf, slide_part)
    rel = next((r for r in rels.values() if r.get("Type", "").endswith("/notesSlide")), None)
    if rel is None:
        return None, ""
    if rel.get("TargetMode") == "External":
        return None, ""
    part = resolve_part(slide_part, rel["Target"])
    notes = ET.fromstring(zf.read(part))
    texts = []
    for shape in notes.findall(".//p:sp", NS):
        ph = shape.find("p:nvSpPr/p:nvPr/p:ph", NS)
        if ph is not None and ph.get("type") in {"sldImg", "sldNum", "dt", "hdr", "ftr"}:
            continue
        text = text_of(shape).strip()
        if text:
            texts.append(text)
    return part, "\n".join(texts)


def _word_cap(value):
    """None leaves the check off. A passed value must be a positive number."""
    if value is None:
        return None
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ValueError("word cap must be a positive number")
    if value <= 0 or value != value or value == float("inf"):
        raise ValueError("word cap must be a positive number")
    return value


def audit(path, word_cap=None):
    word_cap = _word_cap(word_cap)
    if word_cap is None:
        cap_shown = None
        cap_sentence = (
            "No word cap is applied. Pass --word-cap to record onscreen_text_over_word_cap. "
        )
    else:
        cap_shown = format_word_count(word_cap)
        cap_sentence = (
            "An optional word cap of %s is in effect; page 1 is exempt. " % cap_shown
        )
    source = Path(path).resolve(strict=True)
    result = {
        "schema_version": 1,
        "file": str(source),
        "sha256": hashlib.sha256(source.read_bytes()).hexdigest(),
        "claim_boundary": (
            "Read-only package inventory. Counts and large-image hints do not prove editability, "
            "visual correctness, notes usefulness, or scientific accuracy. Render and inspect separately. "
            "AlternateContent uses its first Choice when present; app rendering may choose differently. "
            "Grouped and inherited coordinates are not geometrically validated. "
            "onscreen_word_count estimates screen text at 2 CJK characters per word plus Latin tokens. "
            "It skips slide-number placeholders, digit-only page numbers, and ≤15pt source lines "
            "(a source marker, or a ≤15pt paragraph in the bottom 15% of the slide). "
            "Other ≤15pt body text still counts, so shrinking the font does not hide overflow. "
            + cap_sentence
            + "Runs without an explicit size are counted. "
            "This is not a line-break, bullet-limit, or visual-density verdict. "
            "empty_lower_content_band is a warning only and does not change the process exit code. "
            "The content band runs from 20% to 85% of slide height (body below the title, above the "
            "source/page-number band) and is split at its midline. The warning fires when the upper "
            "half contains a shape at least about 28px wide that extends at least about 28px into "
            "that half, and the lower half has no such shape. Page 1 is exempt. A one-line text shape "
            "in the lower half clears it. Empty text boxes do not. Pictures, tables, charts, "
            "connectors, and group boxes count by geometry. Group children are not measured; the "
            "group box is. Missing group transforms are not inferred. The check does not decide "
            "whether the fill is a conclusion, a figure note, a schematic, or a 缺图 page. "
            "That band check does not judge font size. "
            "algorithm_page_missing_diagram warns when the slide title matches a concept "
            "(方法, 算法, 架构, 框架, 流程, pipeline, framework) and the page has no "
            "connector, picture, table, chart, valid OMML, or non-textbox autoshape "
            "(a no-fill rectangle does not count) and does not contain 「缺图」. "
            "Valid OMML means office_math_count > 0 and no office_math_missing_fallback. "
            "Model names such as Mamba or Transformer are not keywords. Only the title is "
            "matched: a title placeholder, otherwise the topmost non-page-number text shape. "
            "Page 1 is the cover and is exempt. Body text that merely mentions a model does not match. "
            "raw_latex_on_screen warns when on-screen text, including math text, contains "
            "a LaTeX command or ^{ / _{. Unicode letters such as λ do not. "
            "office_math_count is the number of m:oMath elements. "
            "office_math_missing_fallback warns when an m:oMath element is not inside "
            "m:oMathPara / a14:m / mc:Choice Requires=a14, or that AlternateContent has no "
            "mc:Fallback picture (an a:blip). A bare m:oMath inside a:p is this warning. "
            "These warnings do not change the process exit code. "
            "font_below_min warns when a visible body run has an explicit size below 18pt "
            "and the shape sits outside the footer/source band (the bottom 15% of the slide). "
            "Slide-number placeholders, digit-only page numbers, and ≤15pt lines with a source "
            "marker are not body. Runs with no explicit size are not judged. A 16pt caption is "
            "below this body floor and may warn. "
            "low_contrast_text warns when an explicit sRGB text color contrasts below 4.5:1 "
            "with white, and whenever that text color is #BFBFBF. Shape fills and chart series "
            "colors are not text. Inherited theme colors with no explicit sRGB are not judged. "
            "These two warnings do not change the process exit code."
        ),
        "warnings": [],
    }
    with zipfile.ZipFile(source) as zf:
        broken = zf.testzip()
        if broken:
            raise ValueError("Corrupt ZIP member: " + broken)
        names = zf.namelist()
        if len(names) != len(set(names)):
            result["warnings"].append({"code": "duplicate_zip_members"})
        # Check all relationships without following external targets or network requests.
        external = []
        for name in names:
            if not name.endswith(".rels"):
                continue
            for rel in ET.fromstring(zf.read(name)):
                if rel.get("TargetMode") == "External":
                    external.append({"part": name, "id": rel.get("Id"), "type": rel.get("Type"), "target": rel.get("Target")})
        result["external_relationships"] = external
        pres = ET.fromstring(zf.read("ppt/presentation.xml"))
        size = pres.find("p:sldSz", NS)
        if size is None:
            raise ValueError("Missing p:sldSz")
        sw, sh = int(size.get("cx")), int(size.get("cy"))
        result["slide_size_emu"] = [sw, sh]
        result["slide_size_css_px"] = [round(sw / EMU_PER_PX, 2), round(sh / EMU_PER_PX, 2)]
        rels = relationships(zf, "ppt/presentation.xml")
        pages = []
        for page, sid in enumerate(pres.findall("p:sldIdLst/p:sldId", NS), 1):
            rel = rels[sid.get(tag("r", "id"))]
            if rel.get("TargetMode") == "External":
                raise ValueError("Slide relationship must not be external")
            part = resolve_part("ppt/presentation.xml", rel["Target"])
            slide = ET.fromstring(zf.read(part))
            tree = slide.find("p:cSld/p:spTree", NS)
            if tree is None:
                raise ValueError("Missing shape tree: " + part)
            record = {"page": page, "part": part, "native_text_shape_count": 0,
                      "picture_count": 0, "native_table_count": 0, "native_chart_count": 0,
                      "office_math_count": 0, "diagram_object_count": 0, "connector_count": 0,
                      "native_texts": [], "large_picture_hints": [], "warnings": []}
            words = 0.0
            small_pts = []
            pale_colors = []
            for shape, nested in objects(tree):
                if is_diagram_object(shape):
                    record["diagram_object_count"] += 1
                if shape.tag == tag("p", "cxnSp"):
                    record["connector_count"] += 1
                text = text_of(shape)
                info = shape.find(".//p:cNvPr", NS)
                identity = {"id": info.get("id") if info is not None else None,
                            "name": info.get("name") if info is not None else None}
                box = geometry(shape)
                if shape.tag == tag("p", "sp") and text.strip():
                    record["native_text_shape_count"] += 1
                    entry = {**identity, "text": text, "nested_in_group": nested}
                    if box:
                        entry["box_css_px"] = [round(v / EMU_PER_PX, 2) for v in box]
                    record["native_texts"].append(entry)
                    if box and not nested:
                        x, y, w, h = box
                        if x < 0 or y < 0 or x+w > sw or y+h > sh:
                            record["warnings"].append({"code": "direct_text_frame_outside_canvas", **identity})
                record["native_table_count"] += len(shape.findall(".//a:tbl", NS))
                record["native_chart_count"] += len(shape.findall(".//c:chart", NS))
                record["office_math_count"] += len(shape.findall(".//m:oMath", NS))
                if shape.tag == tag("p", "pic"):
                    record["picture_count"] += 1
                    if box and not nested:
                        x, y, w, h = box
                        area = max(0, min(x+w, sw)-max(0, x)) * max(0, min(y+h, sh)-max(0, y)) / (sw*sh)
                        if area >= .75:
                            record["large_picture_hints"].append({**identity, "canvas_area_ratio": round(area, 4), "interpretation": "Inspect visually; may be legitimate source figure or flattened slide."})
                words += onscreen_words(shape, box, sh, nested)
                found_small, found_pale = typography_findings(shape, box, sh, nested)
                small_pts.extend(found_small)
                pale_colors.extend(found_pale)
            shown = format_word_count(words)
            record["onscreen_word_count"] = shown
            record["onscreen_word_cap"] = cap_shown
            record["onscreen_word_cap_exempt"] = word_cap is not None and page == 1
            if word_cap is not None and page != 1 and words > word_cap:
                record["warnings"].append({
                    "code": "onscreen_text_over_word_cap",
                    "word_count": shown,
                    "cap": cap_shown,
                })
            band_boxes = content_band_boxes(tree)
            if page != 1 and lower_content_band_is_empty(band_boxes, sh):
                top, mid, bottom = content_band_bounds(sh)
                record["warnings"].append({
                    "code": "empty_lower_content_band",
                    "content_band_css_px": [round(top / EMU_PER_PX, 2), round(bottom / EMU_PER_PX, 2)],
                    "lower_half_css_px": [round(mid / EMU_PER_PX, 2), round(bottom / EMU_PER_PX, 2)],
                })
            screen = text_of(tree)
            title = slide_title_text(tree, sh)
            record["title_text"] = title
            record["algorithm_intent"] = bool(ALGORITHM_INTENT.search(title))
            record["has_gap_marker"] = bool(GAP_MARKER.search(screen))
            math_warning = office_math_fallback_warning(tree)
            has_valid_office_math = record["office_math_count"] > 0 and math_warning is None
            has_visual = (
                record["diagram_object_count"] > 0
                or record["picture_count"] > 0
                or record["native_table_count"] > 0
                or record["native_chart_count"] > 0
                or has_valid_office_math
            )
            if (
                page != 1
                and record["algorithm_intent"]
                and not has_visual
                and not record["has_gap_marker"]
            ):
                record["warnings"].append({"code": "algorithm_page_missing_diagram"})
            if math_warning:
                record["warnings"].append(math_warning)
            raw_hits = RAW_LATEX.findall(screen)
            if raw_hits:
                record["warnings"].append({
                    "code": "raw_latex_on_screen",
                    "examples": sorted(set(raw_hits))[:6],
                })
            if small_pts:
                record["warnings"].append({
                    "code": "font_below_min",
                    "count": len(small_pts),
                    "min_pt": _format_pt(min(small_pts)),
                    "floor_pt": 18,
                })
            if pale_colors:
                record["warnings"].append({
                    "code": "low_contrast_text",
                    "count": len(pale_colors),
                    "examples": sorted(set(pale_colors))[:6],
                    "minimum_ratio": CONTRAST_MIN_ON_WHITE,
                })
            notes_part, notes = notes_body(zf, part)
            count = sum(not c.isspace() for c in notes)
            record.update({"notes_part": notes_part, "has_nonempty_body_notes": bool(count),
                           "notes_non_whitespace_char_count": count, "notes_text": notes})
            pages.append(record)
        result["slides"] = pages
        result["slide_count"] = len(pages)
        result["slides_without_body_notes"] = [p["page"] for p in pages if not p["has_nonempty_body_notes"]]
        result["slides_with_large_picture_hints"] = [p["page"] for p in pages if p["large_picture_hints"]]
        result["slides_over_word_cap"] = [
            p["page"] for p in pages
            if any(item.get("code") == "onscreen_text_over_word_cap" for item in p["warnings"])
        ]
        result["slides_with_empty_lower_content_band"] = [
            p["page"] for p in pages
            if any(item.get("code") == "empty_lower_content_band" for item in p["warnings"])
        ]
        result["slides_missing_diagram"] = [
            p["page"] for p in pages
            if any(item.get("code") == "algorithm_page_missing_diagram" for item in p["warnings"])
        ]
        result["slides_with_raw_latex"] = [
            p["page"] for p in pages
            if any(item.get("code") == "raw_latex_on_screen" for item in p["warnings"])
        ]
        result["slides_with_invalid_office_math"] = [
            p["page"] for p in pages
            if any(item.get("code") == "office_math_missing_fallback" for item in p["warnings"])
        ]
        result["office_math_count"] = sum(p["office_math_count"] for p in pages)
        result["slides_with_office_math"] = [
            p["page"] for p in pages if p["office_math_count"]
        ]
        result["slides_with_font_below_min"] = [
            p["page"] for p in pages
            if any(item.get("code") == "font_below_min" for item in p["warnings"])
        ]
        result["slides_with_low_contrast_text"] = [
            p["page"] for p in pages
            if any(item.get("code") == "low_contrast_text" for item in p["warnings"])
        ]
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pptx", help="Existing source PPTX; never modified")
    parser.add_argument("--json", dest="json_path", help="Optional new report file (refuses overwrite)")
    parser.add_argument(
        "--word-cap",
        type=float,
        default=None,
        help="Optional on-screen word cap. Off unless passed. 2 CJK characters = 1 word. Page 1 is exempt when the check runs.",
    )
    args = parser.parse_args()
    try:
        report = audit(args.pptx, word_cap=args.word_cap)
        encoded = json.dumps(report, ensure_ascii=False, indent=2)
        if args.json_path:
            target = Path(args.json_path).resolve()
            if target == Path(args.pptx).resolve():
                raise ValueError("Report cannot overwrite the source PPTX")
            target.parent.mkdir(parents=True, exist_ok=True)
            with target.open("x", encoding="utf-8") as stream:
                stream.write(encoded + "\n")
            print(json.dumps({"report": str(target), "slide_count": report["slide_count"], "sha256": report["sha256"]}, ensure_ascii=False))
        else:
            print(encoded)
    except (OSError, ValueError, KeyError, ET.ParseError, zipfile.BadZipFile) as exc:
        print(json.dumps({"error": type(exc).__name__, "message": str(exc)}, ensure_ascii=False), file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
