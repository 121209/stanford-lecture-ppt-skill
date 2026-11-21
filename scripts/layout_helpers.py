#!/usr/bin/env python3
"""Editable layout blocks for a 16:9 lecture slide.

Requires python-pptx. These helpers draw text boxes, a native table, a
picture plus caption, and a thin definition box. They do not draw formulas,
flow diagrams, or video. Charts stay in ``chart_helpers.py``. Click entrance
animations stay in ``appear.py`` and are not applied here.

Colors are the lecture palette: white, #222222, #C00000, #606060. A same-question
comparison may paint its column bars with that accent, or with ink or an optional
navy when the caller asks. Navy is not the default and does not replace the accent.
Body text is at least 20pt. Source lines are 13pt and must carry a source marker.
Page text is digits only. Nothing in this file is a measured result.
"""
import argparse
import io
import math
import re
import struct
import sys
import zlib
from pathlib import Path

SLIDE_W = 12192000
SLIDE_H = 6858000
EMU_IN = 914400
FONT = "Microsoft YaHei"
INK = (0x22, 0x22, 0x22)
ACCENT = (0xC0, 0x00, 0x00)
MUTED = (0x60, 0x60, 0x60)
WHITE = (0xFF, 0xFF, 0xFF)
# Same split as scripts/audit_pptx.py so a helper cap and the audit agree.
CJK_CHAR = re.compile(r"[\u3400-\u9fff]")
LATIN_WORD = re.compile(r"[A-Za-z0-9]+(?:[-'’./][A-Za-z0-9]+)*")
SOURCE_MARK = re.compile(
    r"(来源\s*[:：]|出处\s*[:：]|图源\s*[:：]|资料来源|引自|Source\s*:|DOI\b|doi\s*:|arXiv|https?://)",
    re.IGNORECASE,
)
CONTENT_TOP_RATIO = 0.20
FOOTER_TOP_RATIO = 0.85
MIN_EXTENT = int(round(28 * 9525))
MARGIN_X = int(round(0.70 * EMU_IN))
TEXT_W = int(round(11.90 * EMU_IN))
READ_Y = int(round(5.15 * EMU_IN))
READ_H = int(round(0.50 * EMU_IN))
ROW_H = int(round(0.50 * EMU_IN))
MAX_TABLE_COLS = 6
MAX_TABLE_ROWS = 5
CELL_WORD_CAP = 6
# One 20pt line in a half-width column. This is a line-fit limit, not a slide cap.
POINT_WORD_CAP = 10
# Same-question columns. Caps keep one 20pt line inside the column after inset.
# At 20pt a CJK em is 20/72 in. Counts use the audit rule (2 CJK = 1 word).
# 3 columns: inner text width 3.59 in → 12 CJK → 6 words.
# 2 columns: inner text width 5.57 in → 18 CJK → 9 words, one glyph of slack.
# 3-column headers also hold a digit, so the title itself stops at 5 words.
COMPARE_MIN_COLUMNS = 2
COMPARE_MAX_COLUMNS = 3
COMPARE_MIN_LABELS = 2
COMPARE_MAX_LABELS = 4
COMPARE_MAX_JUDGMENT_LINES = 2
COMPARE_HEADER_PT = 20
COMPARE_FIELD_PT = 20
COMPARE_JUDGMENT_PT = 20
COMPARE_FRAME_PT = 1.0
COMPARE_GUTTER_IN = {2: 0.20, 3: 0.14}
COMPARE_CARD_TOP_IN = 1.12
COMPARE_HEADER_H_IN = 0.46
COMPARE_AFTER_HEADER_IN = 0.10
COMPARE_LABEL_H_IN = 0.36
COMPARE_BODY_H_IN = 0.40
COMPARE_AFTER_FIELD_IN = 0.08
COMPARE_CARD_PAD_IN = 0.12
COMPARE_INSET_IN = 0.14
COMPARE_JUDGMENT_Y_IN = 5.22
COMPARE_JUDGMENT_H_IN = 0.42
COMPARE_JUDGMENT_GAP_IN = 0.06
COMPARE_HEADER_WORD_CAP = {2: 8, 3: 5}
COMPARE_BODY_WORD_CAP = {2: 9, 3: 6}
COMPARE_LABEL_WORD_CAP = 4
COMPARE_KEYWORD_WORD_CAP = 8
COMPARE_JUDGMENT_WORD_CAP = 18
COMPARE_BAR_TEXT = WHITE
# Optional bar colors for this layout only. "accent" is the lecture red.
# "navy" (#1F4E79) is opt-in: white text on it clears 4.5:1, and the hex is
# not measured from a photograph. It does not change ACCENT or other slides.
COMPARE_SCHEMES = {
    "accent": {"bar": ACCENT},
    "ink": {"bar": INK},
    "navy": {"bar": (0x1F, 0x4E, 0x79)},
}
SAME_QUESTION_CUES = (
    "同一个问题", "同一问题", "同题", "同维", "同构比较", "同构栏",
    "两种做法", "三种做法", "两种输入", "三种输入", "两种方案", "三种方案",
    "并列方案", "same question", "same-dimension",
)
PAPER_NOTES_CUES = ("论文工作", "个人收获", "本次笔记", "阅读收获")
DIAGRAM_CUES = ("架构图", "流程图", "流程步骤", "双栏架构")
TABLE_CUES = ("数值表", "结果表", "原生表格")
_SPLIT_HINT = "Split onto the next slide and keep 20pt; do not shrink the type."


def emu(inches):
    return int(round(inches * EMU_IN))


def _pptx():
    try:
        import pptx
        from pptx.dml.color import RGBColor
        from pptx.enum.shapes import MSO_SHAPE
        from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
        from pptx.oxml.ns import qn
        from pptx.util import Emu, Pt
    except ImportError as exc:
        raise RuntimeError("python-pptx is required for layout_helpers.py") from exc
    return {
        "Presentation": pptx.Presentation,
        "RGBColor": RGBColor,
        "MSO_SHAPE": MSO_SHAPE,
        "MSO_ANCHOR": MSO_ANCHOR,
        "PP_ALIGN": PP_ALIGN,
        "qn": qn,
        "Emu": Emu,
        "Pt": Pt,
    }


def _rgb(api, triple):
    return api["RGBColor"](*triple)


def _colors(api):
    return {name: _rgb(api, triple) for name, triple in (
        ("ink", INK), ("accent", ACCENT), ("muted", MUTED), ("white", WHITE),
    )}


def _slide_size(slide):
    presentation = slide.part.package.presentation_part.presentation
    return int(presentation.slide_width), int(presentation.slide_height)


def _require_widescreen(slide):
    width, height = _slide_size(slide)
    if width != SLIDE_W or height != SLIDE_H:
        raise ValueError("layout helpers expect a 16:9 slide, 12192000 by 6858000 EMU")


def _words(text):
    cjk = CJK_CHAR.findall(text)
    rest = CJK_CHAR.sub(" ", text)
    return len(cjk) / 2.0 + len(LATIN_WORD.findall(rest))


def _line(text, limit, label):
    if not isinstance(text, str) or not text.strip():
        raise ValueError(label + " must be non-empty text")
    if "\n" in text:
        raise ValueError(label + " must be a single line")
    cleaned = text.strip()
    count = _words(cleaned)
    if count > limit:
        raise ValueError("%s is %.1f words; keep it to %d so the type stays 20pt" % (
            label, count, limit,
        ))
    return cleaned


def _budget(parts, limit, label):
    """Combined length for one sparse block, so the type stays on one line."""
    total = sum(_words(part) for part in parts)
    if total > limit:
        raise ValueError("%s is %.1f words; keep it to %d so the type stays one line" % (
            label, total, limit,
        ))
    return total


def _strip_style(shape, api):
    style = shape._element.find(api["qn"]("p:style"))
    if style is not None:
        shape._element.remove(style)


def _anchor(frame, api, anchor):
    body = frame._txBody.find(api["qn"]("a:bodyPr"))
    body.set("anchor", anchor)
    for child in list(body):
        if child.tag.endswith("Fit"):
            body.remove(child)
    body.append(body.makeelement(api["qn"]("a:noAutofit"), {}))


def _style_run(run, size, color, api, bold=False):
    run.font.size = api["Pt"](size)
    run.font.bold = bold
    run.font.italic = False
    run.font.color.rgb = color
    run.font.name = FONT
    r_pr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        element = r_pr.find(api["qn"](tag))
        if element is None:
            element = r_pr.makeelement(api["qn"](tag), {})
            r_pr.append(element)
        element.set("typeface", FONT)


def _blank_layout(prs):
    for layout in prs.slide_layouts:
        if (layout.name or "").lower() == "blank":
            return layout
    return prs.slide_layouts[len(prs.slide_layouts) - 1]


def new_deck():
    api = _pptx()
    prs = api["Presentation"]()
    prs.slide_width = api["Emu"](SLIDE_W)
    prs.slide_height = api["Emu"](SLIDE_H)
    return prs


def add_blank_slide(prs):
    api = _pptx()
    slide = prs.slides.add_slide(_blank_layout(prs))
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = _colors(api)["white"]
    return slide


def _textbox(slide, x, y, w, h, api, name=None):
    box = slide.shapes.add_textbox(api["Emu"](x), api["Emu"](y), api["Emu"](w), api["Emu"](h))
    _strip_style(box, api)
    if name:
        box.name = name
    box.text_frame.word_wrap = True
    return box


def add_text(slide, x, y, w, h, text, size, color, align="left", bold=False, anchor="t", name=None):
    """One editable text box. ``color`` is an RGB tuple. Size is in points."""
    api = _pptx()
    _require_widescreen(slide)
    box = _textbox(slide, x, y, w, h, api, name=name)
    _anchor(box.text_frame, api, "ctr" if anchor == "ctr" else "t")
    paragraph = box.text_frame.paragraphs[0]
    paragraph.alignment = {
        "left": api["PP_ALIGN"].LEFT,
        "center": api["PP_ALIGN"].CENTER,
        "right": api["PP_ALIGN"].RIGHT,
    }[align]
    paragraph.text = text
    _style_run(paragraph.runs[0], size, _rgb(api, color), api, bold=bold)
    return box


def _set_runs(paragraph, runs, size, api, align="left"):
    """Replace the paragraph's runs. ``runs`` is (text, RGB tuple, bold)."""
    paragraph.alignment = {
        "left": api["PP_ALIGN"].LEFT,
        "center": api["PP_ALIGN"].CENTER,
    }[align]
    parent = paragraph._p
    for child in list(parent):
        if child.tag == api["qn"]("a:r"):
            parent.remove(child)
    for text, color, bold in runs:
        run = paragraph.add_run()
        run.text = text
        _style_run(run, size, _rgb(api, color), api, bold=bold)
    return paragraph


def _margins(frame, api, left, right, top, bottom):
    frame.margin_left = api["Emu"](left)
    frame.margin_right = api["Emu"](right)
    frame.margin_top = api["Emu"](top)
    frame.margin_bottom = api["Emu"](bottom)


def _write_lines(box, lines, size, color, api, align="left", bold=False):
    frame = box.text_frame
    frame.word_wrap = True
    _anchor(frame, api, "t")
    alignment = api["PP_ALIGN"].CENTER if align == "center" else api["PP_ALIGN"].LEFT
    for index, line in enumerate(lines):
        paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
        paragraph.alignment = alignment
        paragraph.text = line
        _style_run(paragraph.runs[0], size, color, api, bold=bold)
    return box


def add_title(slide, text):
    """32pt accent title in the title band, above the content band."""
    cleaned = _line(text, 16, "title")
    return add_text(
        slide, MARGIN_X, emu(0.36), TEXT_W, emu(0.70),
        cleaned, 32, ACCENT, name="slide-title",
    )


def add_read_line(slide, text, accent=False):
    """One 20pt line in the lower content band. This is a conclusion, not a new paragraph."""
    cleaned = _line(text, 16, "read line")
    return add_text(
        slide, MARGIN_X, READ_Y, TEXT_W, READ_H,
        cleaned, 20, ACCENT if accent else INK, name="read-line",
    )


def format_page(page):
    """Digits only, so the audit can tell a page number from body text."""
    if isinstance(page, bool) or isinstance(page, float):
        raise ValueError("page must be a positive integer or a 1–3 digit string")
    if isinstance(page, int):
        if not 1 <= page <= 999:
            raise ValueError("page must be from 1 to 999")
        return "%02d" % page if page < 100 else str(page)
    if isinstance(page, str) and page.isdigit() and 1 <= len(page) <= 3 and int(page) >= 1:
        return page
    raise ValueError("page must be digits only, not a fraction or a label")


def _named(slide, name):
    for shape in slide.shapes:
        if shape.name == name:
            return shape
    return None


def add_footer(slide, source=None, page=None):
    """13pt source line and a digits-only page number, both in the footer band.

    The source must include a marker such as ``来源：`` or ``Source:``. Calling
    this twice on one slide raises, so a deck does not stack two page numbers.
    """
    _require_widescreen(slide)
    if source is None and page is None:
        raise ValueError("footer needs a source, a page number, or both")
    if _named(slide, "source-line") is not None or _named(slide, "page-number") is not None:
        raise ValueError("footer already exists on this slide")
    if source is not None:
        cleaned = _line(source, 24, "source")
        if not SOURCE_MARK.search(cleaned):
            raise ValueError("source must include 来源： or Source: so it stays a source line")
        add_text(
            slide, MARGIN_X, emu(6.58), emu(10.60), emu(0.30),
            cleaned, 13, MUTED, name="source-line",
        )
    if page is not None:
        add_text(
            slide, emu(11.85), emu(6.58), emu(0.90), emu(0.30),
            format_page(page), 13, MUTED, align="center", anchor="ctr", name="page-number",
        )
    return slide


def set_notes(slide, text):
    if not isinstance(text, str) or not text.strip():
        raise ValueError("notes must be non-empty")
    slide.notes_slide.notes_text_frame.text = text.strip()
    return slide


def add_cover(slide, title, presenter, date, kicker=None, authors=None, venue=None):
    """Paper cover: type, title, authors, venue and year, then presenter and date.

    Authors and venue are optional for a talk that is not a paper. They are
    body text in the content area, not a 13pt footer. The caller should put
    this on slide 1.
    """
    _require_widescreen(slide)
    title = _line(title, 24, "title")
    presenter = _line(presenter, 12, "presenter")
    date = _line(date, 12, "date")
    if kicker is not None:
        kicker = _line(kicker, 8, "kicker")
    if authors is not None:
        authors = _line(authors, 16, "authors")
    if venue is not None:
        venue = _line(venue, 16, "venue")
    y = emu(1.45)
    if kicker:
        add_text(slide, emu(1.10), y, emu(11.10), emu(0.46), kicker, 20, MUTED, align="center", name="kicker")
        y += emu(0.62)
    add_text(slide, emu(1.10), y, emu(11.10), emu(1.30), title, 32, ACCENT, align="center", name="cover-title")
    y += emu(1.55)
    if authors:
        add_text(slide, emu(1.10), y, emu(11.10), emu(0.46), authors, 20, INK, align="center", name="authors")
        y += emu(0.52)
    if venue:
        add_text(slide, emu(1.10), y, emu(11.10), emu(0.46), venue, 20, INK, align="center", name="venue")
        y += emu(0.70)
    else:
        y += emu(0.24)
    add_text(slide, emu(1.10), y, emu(11.10), emu(0.46), presenter, 20, INK, align="center", name="presenter")
    y += emu(0.50)
    add_text(slide, emu(1.10), y, emu(11.10), emu(0.46), date, 20, INK, align="center", name="date")
    return slide


def _occupies_lower(y, h):
    top = int(SLIDE_H * CONTENT_TOP_RATIO)
    bottom = int(SLIDE_H * FOOTER_TOP_RATIO)
    mid = (top + bottom) // 2
    if h < MIN_EXTENT:
        return False
    overlap = max(0, min(y + h, bottom) - max(y, mid))
    return overlap >= MIN_EXTENT


def _box(region):
    if isinstance(region, (str, bytes)) or not isinstance(region, (list, tuple)) or len(region) != 4:
        raise ValueError("region must be (x, y, w, h) in EMU")
    x, y, w, h = (int(value) for value in region)
    if w <= 0 or h <= 0:
        raise ValueError("region width and height must be positive")
    if min(x, y) < 0 or x + w > SLIDE_W or y + h > SLIDE_H:
        raise ValueError("region must lie on the 16:9 slide")
    return x, y, w, h


def add_section_divider(slide, title, kicker=None, source=None, page=None):
    """A sparse section title. The title box is one line and straddles the midline.

    There is no agenda list, rule, or color band. A long contents page does not
    belong here.
    """
    _require_widescreen(slide)
    title = _line(title, 12, "section title")
    if kicker is not None:
        kicker = _line(kicker, 8, "kicker")
    _budget([part for part in (kicker, title) if part], 16, "section divider")
    # 32pt line centered on the content midline. The box is only as tall as the
    # type, so the lower band is the title itself rather than an empty stretch.
    # The kicker sits under the title: the audit treats the topmost text as the
    # slide title, and that should be the section name.
    add_text(
        slide, MARGIN_X, emu(3.52), TEXT_W, emu(0.82),
        title, 32, ACCENT, align="center", anchor="ctr", name="section-title",
    )
    if kicker:
        add_text(
            slide, MARGIN_X, emu(4.50), TEXT_W, emu(0.46),
            kicker, 20, MUTED, align="center", name="kicker",
        )
    if source is not None or page is not None:
        add_footer(slide, source=source, page=page)
    return slide


def add_two_column(slide, title, left_title, left_points, right_title, right_points, conclusion, source=None, page=None):
    """Text comparison. This is not the architecture column layout in diagram_helpers.

    Each column has a short red label and 1–3 black lines. The conclusion is
    required and sits in the lower band in red.
    """
    _require_widescreen(slide)
    title = _line(title, 14, "title")
    left_title = _line(left_title, 8, "left title")
    right_title = _line(right_title, 8, "right title")
    conclusion = _line(conclusion, 16, "conclusion")
    columns = []
    for label, points in (("left", left_points), ("right", right_points)):
        if isinstance(points, (str, bytes)) or not isinstance(points, (list, tuple)) or not 1 <= len(points) <= 3:
            raise ValueError(label + " column needs 1 to 3 points")
        columns.append([_line(point, POINT_WORD_CAP, label + " point") for point in points])
    add_title(slide, title)
    api = _pptx()
    colors = _colors(api)
    left_x = MARGIN_X
    right_x = emu(6.75)
    width = emu(5.70)
    add_text(slide, left_x, emu(1.72), width, emu(0.46), left_title, 20, ACCENT, bold=True, name="left-title")
    add_text(slide, right_x, emu(1.72), width, emu(0.46), right_title, 20, ACCENT, bold=True, name="right-title")
    for origin, points, name in ((left_x, columns[0], "left-points"), (right_x, columns[1], "right-points")):
        box = _textbox(slide, origin, emu(2.28), width, emu(0.48) * len(points), api, name=name)
        _write_lines(box, ["· " + point for point in points], 20, colors["ink"], api)
    add_read_line(slide, conclusion, accent=True)
    if source is not None or page is not None:
        add_footer(slide, source=source, page=page)
    return slide


def _cell_text(value, label):
    if isinstance(value, bool) or value is None:
        raise ValueError(label + " must be text or a finite number")
    if isinstance(value, str):
        return _line(value, CELL_WORD_CAP, label)
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if not math.isfinite(value):
            raise ValueError(label + " must be a finite number")
        return _line(format(value, "g"), CELL_WORD_CAP, label)
    raise ValueError(label + " must be text or a finite number")


def _numeric(text):
    try:
        float(text)
    except ValueError:
        return False
    return True


def _paint_cell(cell, text, header, api):
    from lxml import etree

    colors = _colors(api)
    cell.text = text
    cell.vertical_anchor = api["MSO_ANCHOR"].MIDDLE
    for margin in ("margin_left", "margin_right"):
        setattr(cell, margin, api["Emu"](emu(0.06)))
    for margin in ("margin_top", "margin_bottom"):
        setattr(cell, margin, api["Emu"](emu(0.04)))
    paragraph = cell.text_frame.paragraphs[0]
    paragraph.alignment = api["PP_ALIGN"].RIGHT if (not header and _numeric(text)) else api["PP_ALIGN"].LEFT
    color = colors["accent"] if header else colors["ink"]
    _style_run(paragraph.runs[0], 20, color, api, bold=header)
    cell.fill.solid()
    cell.fill.fore_color.rgb = colors["white"]
    tc_pr = cell._tc.get_or_add_tcPr()
    qn = api["qn"]
    for edge in ("lnL", "lnR", "lnT", "lnB"):
        found = tc_pr.find(qn("a:" + edge))
        if found is not None:
            tc_pr.remove(found)
    lines = []
    for edge in ("lnL", "lnR", "lnT", "lnB"):
        accent_edge = header and edge == "lnB"
        line = etree.Element(qn("a:" + edge))
        line.set("w", "19050" if accent_edge else "9525")
        line.set("cap", "flat")
        line.set("cmpd", "sng")
        line.set("algn", "ctr")
        solid = etree.SubElement(line, qn("a:solidFill"))
        pigment = etree.SubElement(solid, qn("a:srgbClr"))
        pigment.set("val", "C00000" if accent_edge else "606060")
        dash = etree.SubElement(line, qn("a:prstDash"))
        dash.set("val", "solid")
        lines.append(line)
    for line in reversed(lines):
        tc_pr.insert(0, line)


def _quiet_table(table, api):
    props = table._tbl.tblPr
    if props is None:
        return
    for flag in ("bandRow", "bandCol", "firstRow", "firstCol", "lastRow", "lastCol"):
        props.set(flag, "0")
    style = props.find(api["qn"]("a:tableStyleId"))
    if style is not None:
        props.remove(style)


def add_table(slide, headers, rows, region, conclusion=None, caption=None, source=None, page=None):
    """Native table. Header and body cells are selectable text, at 20pt.

    Row height is fixed. A short table does not stretch to fill ``region``.
    When the table and its caption stay in the upper content band, ``conclusion``
    is required and is placed in the lower band. Pass the numbers you were given.
    Cells stay short so 20pt fits one line. There is no whole-slide word cap;
    split the slide when the table no longer fits at that size.
    """
    _require_widescreen(slide)
    api = _pptx()
    if isinstance(headers, (str, bytes)) or not isinstance(headers, (list, tuple)):
        raise ValueError("headers must be a list of strings")
    if not 1 <= len(headers) <= MAX_TABLE_COLS:
        raise ValueError("table needs 1 to %d columns" % MAX_TABLE_COLS)
    header_text = [_line(header, CELL_WORD_CAP, "header") for header in headers]
    if isinstance(rows, (str, bytes)) or not isinstance(rows, (list, tuple)) or not rows:
        raise ValueError("rows must be a non-empty list")
    if len(rows) > MAX_TABLE_ROWS:
        raise ValueError("table accepts at most %d body rows" % MAX_TABLE_ROWS)
    body = []
    for index, row in enumerate(rows, 1):
        if isinstance(row, (str, bytes)) or not isinstance(row, (list, tuple)) or len(row) != len(headers):
            raise ValueError("row %d must have %d cells" % (index, len(headers)))
        body.append([_cell_text(value, "row %d" % index) for value in row])
    if caption is not None:
        caption = _line(caption, 12, "caption")
    x, y, w, h = _box(region)
    table_h = ROW_H * (1 + len(body))
    if table_h > h:
        raise ValueError("region is shorter than the 20pt rows; do not shrink the type")
    caption_h = emu(0.46) if caption else 0
    caption_y = y + table_h + emu(0.08) if caption else y
    stack_bottom = (caption_y + caption_h) if caption else (y + table_h)
    in_lower = _occupies_lower(y, table_h) or (caption is not None and _occupies_lower(caption_y, caption_h))
    if not in_lower and conclusion is None:
        raise ValueError("table stays in the upper band; pass a one-line conclusion")
    if conclusion is not None:
        conclusion = _line(conclusion, 16, "conclusion")
        if _overlaps_y(y, table_h, READ_Y, READ_H) or (
            caption is not None and _overlaps_y(caption_y, caption_h, READ_Y, READ_H)
        ):
            raise ValueError("conclusion overlaps the table; keep the table in the upper band")
    frame = slide.shapes.add_table(
        1 + len(body), len(header_text),
        api["Emu"](x), api["Emu"](y), api["Emu"](w), api["Emu"](table_h),
    )
    frame.name = "data-table"
    _strip_style(frame, api)
    table = frame.table
    _quiet_table(table, api)
    col_w = w // len(header_text)
    for index, column in enumerate(table.columns):
        column.width = api["Emu"](w - col_w * (len(header_text) - 1) if index == len(header_text) - 1 else col_w)
    for row in table.rows:
        row.height = api["Emu"](ROW_H)
    for col, text in enumerate(header_text):
        _paint_cell(table.cell(0, col), text, True, api)
    for r_index, row in enumerate(body, 1):
        for c_index, text in enumerate(row):
            _paint_cell(table.cell(r_index, c_index), text, False, api)
    if caption:
        if stack_bottom > emu(6.40):
            raise ValueError("caption would collide with the footer")
        add_text(slide, x, caption_y, w, caption_h, caption, 20, INK, name="table-caption")
    if conclusion:
        add_read_line(slide, conclusion)
    if source is not None or page is not None:
        add_footer(slide, source=source, page=page)
    return frame


def _overlaps_y(y1, h1, y2, h2):
    return not (y1 + h1 + emu(0.06) <= y2 or y2 + h2 + emu(0.06) <= y1)


def _image_source(image):
    if isinstance(image, (bytes, bytearray)):
        return io.BytesIO(bytes(image))
    if isinstance(image, (str, Path)):
        path = Path(image)
        if not path.is_file():
            raise ValueError("image file was not found")
        return str(path)
    raise ValueError("image must be a path or PNG bytes")


def add_picture_caption(slide, image, region, caption):
    """Picture plus an editable 20pt caption.

    When the picture stays in the upper band, the caption moves to the lower
    band and is the read line. The picture is not a substitute for a native
    chart or diagram.
    """
    _require_widescreen(slide)
    caption = _line(caption, 12, "caption")
    x, y, w, h = _box(region)
    api = _pptx()
    picture = slide.shapes.add_picture(
        _image_source(image), api["Emu"](x), api["Emu"](y), api["Emu"](w), api["Emu"](h),
    )
    picture.name = "figure"
    cap_h = emu(0.46)
    cap_y = y + h + emu(0.10) if _occupies_lower(y, h) else READ_Y
    if _overlaps_y(y, h, cap_y, cap_h):
        raise ValueError("caption overlaps the picture; use a shorter picture region")
    if cap_y + cap_h > emu(6.40):
        raise ValueError("caption would collide with the footer")
    cap_w = min(max(w, emu(6.0)), SLIDE_W - x)
    caption_box = add_text(slide, x, cap_y, cap_w, cap_h, caption, 20, INK, name="caption")
    return picture, caption_box


def _no_effects(shape, api):
    _strip_style(shape, api)
    sp_pr = shape._element.spPr
    effect = sp_pr.find(api["qn"]("a:effectLst"))
    if effect is not None:
        sp_pr.remove(effect)


def add_callout(slide, region, term, body, read=None):
    """White rectangle, 1.25pt red rule, no shadow and no gradient.

    ``term`` is red. ``body`` is black. Both are editable. A callout that sits
    in the upper band needs ``read`` so the lower band is a real sentence.
    """
    _require_widescreen(slide)
    term = _line(term, 8, "term")
    body = _line(body, 16, "definition")
    x, y, w, h = _box(region)
    if not _occupies_lower(y, h) and read is None:
        raise ValueError("callout stays in the upper band; pass a one-line read")
    if read is not None:
        read = _line(read, 16, "read line")
        if _overlaps_y(y, h, READ_Y, READ_H):
            raise ValueError("read line overlaps the callout")
    api = _pptx()
    colors = _colors(api)
    shape = slide.shapes.add_shape(
        api["MSO_SHAPE"].RECTANGLE, api["Emu"](x), api["Emu"](y), api["Emu"](w), api["Emu"](h),
    )
    shape.name = "callout"
    _no_effects(shape, api)
    shape.fill.solid()
    shape.fill.fore_color.rgb = colors["white"]
    shape.line.color.rgb = colors["accent"]
    shape.line.width = api["Pt"](1.25)
    frame = shape.text_frame
    frame.word_wrap = True
    frame.margin_left = api["Emu"](emu(0.18))
    frame.margin_right = api["Emu"](emu(0.16))
    frame.margin_top = api["Emu"](emu(0.10))
    frame.margin_bottom = api["Emu"](emu(0.08))
    _anchor(frame, api, "t")
    term_p = frame.paragraphs[0]
    term_p.alignment = api["PP_ALIGN"].LEFT
    term_p.text = term
    _style_run(term_p.runs[0], 20, colors["accent"], api, bold=True)
    body_p = frame.add_paragraph()
    body_p.alignment = api["PP_ALIGN"].LEFT
    body_p.text = body
    _style_run(body_p.runs[0], 20, colors["ink"], api, bold=False)
    if read:
        add_read_line(slide, read)
    return shape


def _has_cue(text, cues):
    folded = text.lower()
    for cue in cues:
        if all(ord(char) < 128 for char in cue):
            if cue.lower() in folded:
                return cue
        elif cue in text:
            return cue
    return None


def choose_compare_layout(brief):
    """Route a page brief to one comparison layout, or None.

    ``same_question`` needs an explicit shared-question cue (同一个问题, 同题,
    同维, 两种/三种做法 or 输入 or 方案, 并列方案, ``same question``,
    ``same-dimension``). A bare 对比 is not enough.

    ``two_column`` is paper work versus notes. ``diagram`` is an architecture
    or a step sequence. ``table`` is a numeric grid. Two of these cues in one
    brief raise, so the page is not given both layouts.
    """
    if not isinstance(brief, str) or not brief.strip():
        raise ValueError("brief must be non-empty text")
    text = brief.strip()
    hits = []
    if _has_cue(text, SAME_QUESTION_CUES):
        hits.append("same_question")
    if _has_cue(text, PAPER_NOTES_CUES):
        hits.append("two_column")
    if _has_cue(text, DIAGRAM_CUES):
        hits.append("diagram")
    if _has_cue(text, TABLE_CUES):
        hits.append("table")
    if len(hits) > 1:
        raise ValueError("brief matches %s; one page uses one layout" % " and ".join(hits))
    if not hits:
        return None
    return hits[0]


def compare_column_boxes(n_columns, n_labels, n_judgments=1):
    """EMU boxes for a 2- or 3-column same-question slide.

    Column widths match, gutters match, and the judgment stays below the cards
    and above the footer. The function raises instead of shrinking type.
    """
    if isinstance(n_columns, bool) or not isinstance(n_columns, int):
        raise ValueError("column count must be an integer")
    if n_columns not in (COMPARE_MIN_COLUMNS, COMPARE_MAX_COLUMNS):
        raise ValueError(
            "same-question compare takes 2 or 3 columns; got %d. "
            "Split extra options onto the next slide at 20pt, and repeat the shared question."
            % n_columns
        )
    if isinstance(n_labels, bool) or not isinstance(n_labels, int):
        raise ValueError("label count must be an integer")
    if not COMPARE_MIN_LABELS <= n_labels <= COMPARE_MAX_LABELS:
        raise ValueError(
            "same-question compare needs %d to %d shared labels; got %d. "
            "Split extra dimensions onto the next slide at 20pt."
            % (COMPARE_MIN_LABELS, COMPARE_MAX_LABELS, n_labels)
        )
    if isinstance(n_judgments, bool) or not isinstance(n_judgments, int) or not (
        1 <= n_judgments <= COMPARE_MAX_JUDGMENT_LINES
    ):
        raise ValueError("judgment must be 1 or 2 lines so 20pt stays above the footer")
    gutter = emu(COMPARE_GUTTER_IN[n_columns])
    usable = TEXT_W - gutter * (n_columns - 1)
    width = usable // n_columns
    if width * n_columns != usable:
        raise ValueError("column width does not divide evenly; check the gutter")
    inset = emu(COMPARE_INSET_IN)
    card_y = emu(COMPARE_CARD_TOP_IN)
    header_h = emu(COMPARE_HEADER_H_IN)
    cursor = card_y + header_h + emu(COMPARE_AFTER_HEADER_IN)
    label_h = emu(COMPARE_LABEL_H_IN)
    body_h = emu(COMPARE_BODY_H_IN)
    field_gap = emu(COMPARE_AFTER_FIELD_IN)
    fields = []
    for index in range(n_labels):
        fields.append({
            "label_y": cursor,
            "label_h": label_h,
            "body_y": cursor + label_h,
            "body_h": body_h,
        })
        cursor += label_h + body_h
        if index != n_labels - 1:
            cursor += field_gap
    card_bottom = cursor + emu(COMPARE_CARD_PAD_IN)
    judgment_y = emu(COMPARE_JUDGMENT_Y_IN)
    judgment_h = emu(COMPARE_JUDGMENT_H_IN)
    judgment_gap = emu(COMPARE_JUDGMENT_GAP_IN)
    if card_bottom + emu(0.10) > judgment_y:
        raise ValueError(
            "columns would collide with the judgment; split a label row onto the next slide at 20pt"
        )
    judgment_bottom = judgment_y + n_judgments * judgment_h + (n_judgments - 1) * judgment_gap
    if judgment_bottom > emu(6.40):
        raise ValueError("judgment would collide with the footer; keep it to two 20pt lines")
    columns = []
    for index in range(n_columns):
        x = MARGIN_X + index * (width + gutter)
        columns.append({
            "x": x,
            "w": width,
            "frame": (x, card_y, width, card_bottom - card_y),
            "header": (x, card_y, width, header_h),
            "text_x": x + inset,
            "text_w": width - 2 * inset,
        })
    return {
        "columns": columns,
        "fields": fields,
        "card_bottom": card_bottom,
        "judgment_y": judgment_y,
        "judgment_h": judgment_h,
        "judgment_gap": judgment_gap,
        "gutter": gutter,
    }


def _fit_line(text, limit, label):
    try:
        return _line(text, limit, label)
    except ValueError as exc:
        raise ValueError(str(exc) + " " + _SPLIT_HINT) from None


def _one_judgment(item):
    if isinstance(item, (list, tuple)):
        if len(item) != 2:
            raise ValueError("judgment pair must be (keyword, conclusion)")
        keyword, conclusion = item
    elif isinstance(item, str):
        if "：" in item:
            keyword, conclusion = item.split("：", 1)
        elif ":" in item:
            keyword, conclusion = item.split(":", 1)
        else:
            raise ValueError("judgment needs a keyword and a conclusion, separated by ：")
    else:
        raise ValueError("judgment line must be text or a (keyword, conclusion) pair")
    keyword = _fit_line(keyword, COMPARE_KEYWORD_WORD_CAP, "judgment keyword")
    conclusion = _fit_line(conclusion, COMPARE_JUDGMENT_WORD_CAP, "judgment conclusion")
    try:
        _budget([keyword, conclusion], COMPARE_JUDGMENT_WORD_CAP, "judgment line")
    except ValueError as exc:
        raise ValueError(str(exc) + " " + _SPLIT_HINT) from None
    return keyword, conclusion


def _judgment_items(judgment):
    if isinstance(judgment, (str, tuple)):
        items = [judgment]
    elif isinstance(judgment, list):
        items = list(judgment)
    else:
        raise ValueError("judgment must be 1 or 2 lines")
    if not 1 <= len(items) <= COMPARE_MAX_JUDGMENT_LINES:
        raise ValueError(
            "judgment must be 1 or 2 lines. Split the synthesis onto the next slide at 20pt."
        )
    return [_one_judgment(item) for item in items]


def _column_fields(column, index):
    if isinstance(column, dict) and "title" in column and isinstance(column.get("fields"), dict):
        title, fields = column["title"], column["fields"]
    elif (
        isinstance(column, (list, tuple))
        and len(column) == 2
        and isinstance(column[1], dict)
    ):
        title, fields = column
    else:
        raise ValueError("column %d must be {title, fields} or (title, fields)" % index)
    return title, fields


def _normalize_compare(columns, labels, judgment):
    if isinstance(columns, (str, bytes)) or not isinstance(columns, (list, tuple)):
        raise ValueError(
            "same-question compare takes 2 or 3 columns. "
            "Split extra options onto the next slide at 20pt, and repeat the shared question."
        )
    if not COMPARE_MIN_COLUMNS <= len(columns) <= COMPARE_MAX_COLUMNS:
        raise ValueError(
            "same-question compare takes 2 or 3 columns; got %d. "
            "Split extra options onto the next slide at 20pt, and repeat the shared question."
            % len(columns)
        )
    parsed = []
    for index, column in enumerate(columns, 1):
        parsed.append(_column_fields(column, index))
    if labels is None:
        labels = list(parsed[0][1].keys())
    elif isinstance(labels, (str, bytes)) or not isinstance(labels, (list, tuple)):
        raise ValueError("labels must be the shared list, in one order")
    else:
        labels = list(labels)
    if not COMPARE_MIN_LABELS <= len(labels) <= COMPARE_MAX_LABELS:
        raise ValueError(
            "same-question compare needs %d to %d shared labels; got %d. "
            "Split extra dimensions onto the next slide at 20pt."
            % (COMPARE_MIN_LABELS, COMPARE_MAX_LABELS, len(labels))
        )
    cleaned_labels = []
    for label in labels:
        text = _fit_line(label, COMPARE_LABEL_WORD_CAP, "label")
        if text in cleaned_labels:
            raise ValueError("labels must be unique and shared")
        cleaned_labels.append(text)
    n_columns = len(parsed)
    header_cap = COMPARE_HEADER_WORD_CAP[n_columns]
    body_cap = COMPARE_BODY_WORD_CAP[n_columns]
    normalized = []
    for index, (title, fields) in enumerate(parsed, 1):
        title = _fit_line(title, header_cap, "column %d title" % index)
        if list(fields.keys()) != cleaned_labels:
            raise ValueError(
                "column %d labels must be %s, in that order"
                % (index, " / ".join(cleaned_labels))
            )
        body = []
        for label in cleaned_labels:
            body.append(_fit_line(fields[label], body_cap, "column %d field %s" % (index, label)))
        normalized.append((title, body))
    judgments = _judgment_items(judgment)
    return cleaned_labels, normalized, judgments


def _ruled_rect(slide, name, box, fill, line, api):
    x, y, w, h = box
    shape = slide.shapes.add_shape(
        api["MSO_SHAPE"].RECTANGLE, api["Emu"](x), api["Emu"](y), api["Emu"](w), api["Emu"](h),
    )
    shape.name = name
    _no_effects(shape, api)
    shape.fill.solid()
    shape.fill.fore_color.rgb = _rgb(api, fill)
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = _rgb(api, line)
        shape.line.width = api["Pt"](COMPARE_FRAME_PT)
    return shape


def add_same_question(slide, title, columns, judgment, labels=None, scheme="accent", source=None, page=None):
    """One question, two or three isomorphic columns, then a judgment.

    Every column uses the same labels in the same order. The numbered bar is
    the lecture accent unless ``scheme`` is ``ink`` or ``navy``. ``navy``
    colors only these bars. Titles, field text, and the judgment keyword stay
    on the lecture palette. Field text is one 20pt line; a longer line or a
    fourth column raises instead of shrinking the type.

    ``judgment`` is ``"关键词：结论"``, ``(keyword, conclusion)``, or a list
    of one or two of those. The keyword is accent and the conclusion is ink.
    """
    _require_widescreen(slide)
    if scheme not in COMPARE_SCHEMES:
        raise ValueError("scheme must be accent, ink, or navy")
    cleaned_labels, normalized, judgments = _normalize_compare(columns, labels, judgment)
    metrics = compare_column_boxes(len(normalized), len(cleaned_labels), len(judgments))
    add_title(slide, title)
    api = _pptx()
    bar = COMPARE_SCHEMES[scheme]["bar"]
    frames = []
    headers = []
    label_boxes = []
    body_boxes = []
    for index, (column_title, bodies) in enumerate(normalized, 1):
        column = metrics["columns"][index - 1]
        frame = _ruled_rect(
            slide, "compare-frame-%d" % index, column["frame"], WHITE, MUTED, api,
        )
        frames.append(frame)
        header = _ruled_rect(
            slide, "compare-header-%d" % index, column["header"], bar, None, api,
        )
        text_frame = header.text_frame
        text_frame.word_wrap = False
        _margins(text_frame, api, emu(0.12), emu(0.08), 0, 0)
        _anchor(text_frame, api, "ctr")
        _set_runs(
            text_frame.paragraphs[0],
            [("%d  %s" % (index, column_title), COMPARE_BAR_TEXT, True)],
            COMPARE_HEADER_PT, api,
        )
        headers.append(header)
        for field_index, (label, body) in enumerate(zip(cleaned_labels, bodies), 1):
            field = metrics["fields"][field_index - 1]
            label_box = _textbox(
                slide, column["text_x"], field["label_y"], column["text_w"], field["label_h"],
                api, name="compare-label-%d-%d" % (index, field_index),
            )
            _margins(label_box.text_frame, api, 0, 0, 0, 0)
            _write_lines(label_box, [label], COMPARE_FIELD_PT, _rgb(api, INK), api, bold=True)
            body_box = _textbox(
                slide, column["text_x"], field["body_y"], column["text_w"], field["body_h"],
                api, name="compare-body-%d-%d" % (index, field_index),
            )
            _margins(body_box.text_frame, api, 0, 0, 0, 0)
            _write_lines(body_box, [body], COMPARE_FIELD_PT, _rgb(api, INK), api, bold=False)
            label_boxes.append(label_box)
            body_boxes.append(body_box)
    judgment_boxes = []
    for index, (keyword, conclusion) in enumerate(judgments, 1):
        y = metrics["judgment_y"] + (index - 1) * (metrics["judgment_h"] + metrics["judgment_gap"])
        box = _textbox(
            slide, MARGIN_X, y, TEXT_W, metrics["judgment_h"], api,
            name="compare-judgment-%d" % index,
        )
        _margins(box.text_frame, api, 0, 0, 0, 0)
        _anchor(box.text_frame, api, "ctr")
        _set_runs(
            box.text_frame.paragraphs[0],
            [(keyword + "：", ACCENT, True), (conclusion, INK, False)],
            COMPARE_JUDGMENT_PT, api,
        )
        judgment_boxes.append(box)
    if source is not None or page is not None:
        add_footer(slide, source=source, page=page)
    return {
        "frames": frames,
        "headers": headers,
        "labels": label_boxes,
        "bodies": body_boxes,
        "judgments": judgment_boxes,
    }


def _png_chunk(tag, data):
    crc = zlib.crc32(tag + data) & 0xFFFFFFFF
    return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", crc)


def _synthetic_mark_png(width=740, height=325):
    """White PNG with one red bar. A placeholder, not a measurement."""
    row_white = b"\x00" + (b"\xff\xff\xff" * width)
    row_red = b"\x00" + (b"\xc0\x00\x00" * width)
    raw = []
    for y in range(height):
        raw.append(row_red if height // 2 - 6 <= y <= height // 2 + 6 else row_white)
    ihdr = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    return b"\x89PNG\r\n\x1a\n" + _png_chunk(b"IHDR", ihdr) + _png_chunk(b"IDAT", zlib.compress(b"".join(raw), 9)) + _png_chunk(b"IEND", b"")


def _finish(slide, source, page, notes):
    if source is not None or page is not None:
        add_footer(slide, source=source, page=page)
    if notes:
        set_notes(slide, notes)


def write_sample_deck(path):
    """Ten anonymous slides. Every number is a placeholder, not a result."""
    charts = _load("chart_helpers")
    appear = _load("appear")
    prs = new_deck()
    cover = add_blank_slide(prs)
    add_cover(
        cover,
        title="示意题目：匿名步骤比较",
        kicker="论文精读",
        authors="甲作者，乙作者",
        venue="匿名会议，2024",
        presenter="汇报人：示意",
        date="日期：示意",
    )
    _finish(cover, None, 1, "封面只放题目、作者和会议。这些名字是占位，不是一篇真实论文。")

    section = add_blank_slide(prs)
    add_section_divider(
        section, "这一节只讲设定", kicker="第二节",
        source="来源：匿名提纲，不是论文目录。", page=2,
    )
    set_notes(section, "节分隔只有一行标题。不要在这里列很长的议程。")

    table = add_blank_slide(prs)
    add_title(table, "两组示意整数")
    add_table(
        table,
        ["组别", "甲", "乙"],
        [["一组", 2, 4], ["二组", 8, 8]],
        (emu(0.70), emu(1.80), emu(10.80), emu(2.20)),
        caption="列是占位，不是指标。",
        conclusion="大的格子只是占位。",
        source="来源：手写整数，不是实验。",
        page=3,
    )
    set_notes(table, "表是原生单元格，可以选中改字。下半部那一句是读表，不是新的结果。")

    columns = add_blank_slide(prs)
    add_two_column(
        columns,
        "论文工作和笔记分开",
        "论文工作",
        ["只定义步骤", "只给出整数"],
        "本次笔记",
        ["步骤不是结论", "整数不能外推"],
        "右栏是笔记，不是已证。",
        source="来源：匿名笔记，不是论文原文。",
        page=4,
    )
    set_notes(columns, "左栏是论文做了什么，右栏是这次的理解。不要把右栏说成论文已经证明。")

    line = add_blank_slide(prs)
    add_title(line, "折线只说明画法")
    charts.add_line_chart(
        line, ["甲", "乙", "丙", "丁"], [2, 5, 9, 9],
        (emu(0.80), emu(1.70), emu(11.40), emu(3.15)),
    )
    add_read_line(line, "末两点相同，没有结论。")
    _finish(line, "来源：手写整数，不是测量。", 5, "折线是原生图表。四个数是占位，不能当成测量。")

    scatter = add_blank_slide(prs)
    add_title(scatter, "散点只是四个位置")
    charts.add_scatter_chart(
        scatter, [(1, 2), (2, 8), (3, 3), (4, 4)],
        (emu(0.80), emu(1.70), emu(11.40), emu(3.15)),
    )
    add_read_line(scatter, "四点不成规律，不能外推。")
    _finish(scatter, "来源：手写坐标，不是测量。", 6, "散点没有连线。坐标是手写的，不要读出趋势。")

    figure = add_blank_slide(prs)
    add_title(figure, "红线只是占位图")
    add_picture_caption(
        figure, _synthetic_mark_png(),
        (emu(0.80), emu(1.70), emu(7.40), emu(3.25)),
        "图注可选中，不是数据。",
    )
    _finish(figure, "来源：生成的色条，不是论文图。", 7, "这张图是一条红线，用来占位。图注是文本框，可以改。")

    definition = add_blank_slide(prs)
    add_title(definition, "定义放在细框里")
    add_callout(
        definition,
        (emu(0.70), emu(1.85), emu(11.60), emu(1.70)),
        "示意项",
        "没有测量含义。",
        read="框线不是卡片，也没有阴影。",
    )
    _finish(definition, "来源：现场定义，不是引文。", 8, "定义框是白底红线。术语和解释都可以选中修改。")

    compare = add_blank_slide(prs)
    add_same_question(
        compare,
        "同一个问题，三种示意做法",
        [
            {"title": "示意甲", "fields": {"输入": "占位短句甲", "输出": "占位短句乙", "局限": "不能当结果"}},
            {"title": "示意乙", "fields": {"输入": "另一句占位", "输出": "仍是占位", "局限": "没有测量"}},
            {"title": "示意丙", "fields": {"输入": "第三句占位", "输出": "第三句输出", "局限": "不可外推"}},
        ],
        [("示例判断", "三栏文字都是占位，不能当成检索或实验。")],
        source="来源：手写占位，不是论文。",
        page=9,
    )
    set_notes(
        compare,
        "示例页。三栏共用输入、输出、局限。栏内句子是占位，不是投影截图里的检索式或文献条数。默认栏条是讲义红，不是蓝。",
    )

    build = add_blank_slide(prs)
    add_title(build, "单击后三行才到齐")
    lines = [
        add_text(build, MARGIN_X, emu(1.90), TEXT_W, emu(0.50), "第一下出现这行", 20, INK, name="step-1"),
        add_text(build, MARGIN_X, emu(2.70), TEXT_W, emu(0.50), "第二下出现这行", 20, INK, name="step-2"),
        add_text(build, MARGIN_X, emu(4.70), TEXT_W, emu(0.50), "下半部这行最后出现", 20, INK, name="step-3"),
    ]
    appear.add_click_appear(build, lines, effect="appear", together=False)
    _finish(build, "来源：演示单击，不是影片。", 10, "这三行是可选的单击出现，默认不要开。它不是视频，文字仍可编辑。")

    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(target))
    return target


def _load(name):
    import importlib.util
    path = Path(__file__).resolve().parent / (name + ".py")
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--samples", help="Write the anonymous layout sample deck here")
    args = parser.parse_args(argv)
    try:
        if not args.samples:
            parser.error("Pass --samples")
        write_sample_deck(args.samples)
    except (OSError, ValueError, RuntimeError) as exc:
        print(type(exc).__name__ + ": " + str(exc), file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
