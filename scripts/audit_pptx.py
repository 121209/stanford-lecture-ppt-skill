#!/usr/bin/env python3
"""Read-only OOXML inventory. This is not a visual or scientific QA verdict.

Also estimates per-slide on-screen word counts against a 40-word cap (cover exempt),
and warns when the lower half of the content band is largely empty. Algorithm-intent
pages without a diagram shape, connector, picture, or 「缺图」 marker, and slides
that show raw LaTeX, are warnings too. None of these warnings change the process
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
}
EMU_PER_PX = 9525
WORD_CAP = 40
# Explicit run size is in hundredths of a point. 15pt source lines are excluded.
SOURCE_LINE_MAX_HUNDREDTHS_PT = 1500
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
# Page text that should carry a native diagram, a picture, or an explicit gap.
# Bare 「步骤」 and 「方法」 are omitted so discussion pages are not flagged.
ALGORITHM_INTENT = re.compile(
    r"算法|伪代码|训练流程|推理流程|推理步骤|解题流程|流程图|"
    r"架构|模型结构|系统模块|"
    r"encoder|decoder|transformer|mamba|pipeline",
    re.IGNORECASE,
)
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


def audit(path):
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
            "Page 1 is exempt as the cover. Runs without an explicit size are counted. "
            "This is not a line-break, bullet-limit, or visual-density verdict. "
            "empty_lower_content_band is a warning only and does not change the process exit code. "
            "The content band runs from 20% to 85% of slide height (body below the title, above the "
            "source/page-number band) and is split at its midline. The warning fires when the upper "
            "half contains a shape at least about 28px wide that extends at least about 28px into "
            "that half, and the lower half has no such shape. Page 1 is exempt. A one-line text shape "
            "in the lower half clears it. Empty text boxes do not. Pictures, tables, charts, "
            "connectors, and group boxes count by geometry. Group children are not measured; the "
            "group box is. Missing group transforms are not inferred. The check does not decide "
            "whether the fill is a conclusion, a figure note, a schematic, or a 缺图 page, and it "
            "does not judge font size. "
            "algorithm_page_missing_diagram warns when on-screen text matches an algorithm, "
            "architecture, or multi-step flow cue and the page has no connector, picture, "
            "or non-textbox autoshape (a no-fill rectangle does not count) and does not "
            "contain 「缺图」. raw_latex_on_screen warns when on-screen text, including math "
            "text, contains a LaTeX command or ^{ / _{. Unicode letters such as λ do not. "
            "office_math_count is the number of m:oMath elements. These warnings do not "
            "change the process exit code."
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
            shown = format_word_count(words)
            record["onscreen_word_count"] = shown
            record["onscreen_word_cap"] = WORD_CAP
            record["onscreen_word_cap_exempt"] = page == 1
            if page != 1 and words > WORD_CAP:
                record["warnings"].append({
                    "code": "onscreen_text_over_word_cap",
                    "word_count": shown,
                    "cap": WORD_CAP,
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
            record["algorithm_intent"] = bool(ALGORITHM_INTENT.search(screen))
            record["has_gap_marker"] = bool(GAP_MARKER.search(screen))
            if record["algorithm_intent"] and record["diagram_object_count"] == 0 and not record["has_gap_marker"]:
                record["warnings"].append({"code": "algorithm_page_missing_diagram"})
            raw_hits = RAW_LATEX.findall(screen)
            if raw_hits:
                record["warnings"].append({
                    "code": "raw_latex_on_screen",
                    "examples": sorted(set(raw_hits))[:6],
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
        result["office_math_count"] = sum(p["office_math_count"] for p in pages)
        result["slides_with_office_math"] = [
            p["page"] for p in pages if p["office_math_count"]
        ]
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pptx", help="Existing source PPTX; never modified")
    parser.add_argument("--json", dest="json_path", help="Optional new report file (refuses overwrite)")
    args = parser.parse_args()
    try:
        report = audit(args.pptx)
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
