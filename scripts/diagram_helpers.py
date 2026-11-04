#!/usr/bin/env python3
"""Place editable flow and architecture shapes on a 16:9 slide.

Requires python-pptx. Colors stay on the skill palette (white, #222222, #C00000,
#606060). No image API, Mermaid binary, Graphviz, officecli, or TikZ.
"""
import argparse
import json
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

SLIDE_W = 12192000
SLIDE_H = 6858000
EMU_IN = 914400
MAX_NODES = 12
FONT = "Microsoft YaHei"

FLOW_NODES = [
    {"id": "read", "label": "读入", "shape": "round", "accent": True},
    {"id": "encode", "label": "编码", "shape": "round"},
    {"id": "update", "label": "更新", "shape": "round"},
    {"id": "score", "label": "评价", "shape": "round"},
]
FLOW_EDGES = [
    {"src": "read", "dst": "encode"},
    {"src": "encode", "dst": "update"},
    {"src": "update", "dst": "score"},
]
ARCH_NODES = [
    {"id": "encode", "label": "编码", "shape": "rect", "column": 0, "accent": True},
    {"id": "state", "label": "状态", "shape": "rect", "column": 0},
    {"id": "decode", "label": "解码", "shape": "rect", "column": 1, "accent": True},
    {"id": "out", "label": "输出", "shape": "round", "column": 1},
]
ARCH_EDGES = [
    {"src": "encode", "dst": "state"},
    {"src": "decode", "dst": "out"},
    {"src": "encode", "dst": "decode"},
]


def emu(inches):
    return int(round(inches * EMU_IN))


def _pptx():
    try:
        import pptx
        from pptx.dml.color import RGBColor
        from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
        from pptx.enum.text import PP_ALIGN
        from pptx.oxml.ns import qn
        from pptx.util import Emu, Pt
    except ImportError as exc:
        raise RuntimeError(
            "python-pptx is required for diagram_helpers.py"
        ) from exc
    return {
        "Presentation": pptx.Presentation,
        "RGBColor": RGBColor,
        "MSO_CONNECTOR": MSO_CONNECTOR,
        "MSO_SHAPE": MSO_SHAPE,
        "PP_ALIGN": PP_ALIGN,
        "qn": qn,
        "Emu": Emu,
        "Pt": Pt,
    }


def _colors(api):
    rgb = api["RGBColor"]
    return {
        "ink": rgb(0x22, 0x22, 0x22),
        "accent": rgb(0xC0, 0x00, 0x00),
        "muted": rgb(0x60, 0x60, 0x60),
        "white": rgb(0xFF, 0xFF, 0xFF),
    }


def _strip_style(shape, qn):
    style = shape._element.find(qn("p:style"))
    if style is not None:
        shape._element.remove(style)


def _style_run(run, size, color, api, bold=False):
    run.font.size = api["Pt"](size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = FONT
    r_pr = run._r.get_or_add_rPr()
    qn = api["qn"]
    for tag in ("a:latin", "a:ea", "a:cs"):
        element = r_pr.find(qn(tag))
        if element is None:
            element = r_pr.makeelement(qn(tag), {})
            r_pr.append(element)
        element.set("typeface", FONT)


def _anchor(text_frame, api, anchor="ctr"):
    qn = api["qn"]
    body = text_frame._txBody.find(qn("a:bodyPr"))
    body.set("anchor", anchor)
    for child in list(body):
        if child.tag.endswith("Fit"):
            body.remove(child)
    body.append(body.makeelement(qn("a:noAutofit"), {}))


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
    colors = _colors(api)
    slide = prs.slides.add_slide(_blank_layout(prs))
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = colors["white"]
    return slide


def add_text(slide, x, y, w, h, text, size, color, api, align="left", bold=False):
    colors = _colors(api)
    box = slide.shapes.add_textbox(api["Emu"](x), api["Emu"](y), api["Emu"](w), api["Emu"](h))
    _strip_style(box, api["qn"])
    frame = box.text_frame
    frame.word_wrap = True
    _anchor(frame, api, "ctr" if align == "center" else "t")
    paragraph = frame.paragraphs[0]
    paragraph.alignment = api["PP_ALIGN"].CENTER if align == "center" else api["PP_ALIGN"].LEFT
    paragraph.text = text
    _style_run(paragraph.runs[0], size, color, api, bold)
    return box


def _shape_kind(name, api):
    kinds = {
        "round": api["MSO_SHAPE"].ROUNDED_RECTANGLE,
        "rect": api["MSO_SHAPE"].RECTANGLE,
        "diamond": api["MSO_SHAPE"].DIAMOND,
        "ellipse": api["MSO_SHAPE"].OVAL,
    }
    if name not in kinds:
        raise ValueError("Unknown shape: " + str(name))
    return kinds[name]


def _text_units(label):
    units = 0.0
    for char in label:
        if "\u3400" <= char <= "\u9fff":
            units += 1.0
        elif char.isspace():
            units += 0.35
        else:
            units += 0.55
    return max(units, 0.5)


def _diamond_font_size(label, width):
    """Shrink so one line fits the diamond's center band. 20pt is the body size."""
    usable = max(width * 0.55, emu(0.4))
    size = 20
    while size > 12 and _text_units(label) * (size / 72.0) * EMU_IN > usable:
        size -= 1
    return size


def _fit_diamond_box(label, x, y, w, h):
    """Widen a narrow diamond just enough for a 14pt one-line label."""
    need = int(_text_units(label) * (14 / 72.0) * EMU_IN / 0.55) + emu(0.12)
    target = min(max(w, need), emu(3.2))
    if target > w:
        x -= (target - w) // 2
        w = target
    if x < 0:
        x = 0
    if x + w > SLIDE_W:
        x = max(0, SLIDE_W - w)
    return x, y, w, h


def _diamond_body(text_frame, api):
    qn = api["qn"]
    body = text_frame._txBody.find(qn("a:bodyPr"))
    body.set("wrap", "none")
    for inset in ("lIns", "tIns", "rIns", "bIns"):
        body.set(inset, "0")


def add_node(slide, kind, x, y, w, h, label, accent, api):
    colors = _colors(api)
    if kind == "diamond":
        x, y, w, h = _fit_diamond_box(label, x, y, w, h)
    shape = slide.shapes.add_shape(
        _shape_kind(kind, api), api["Emu"](x), api["Emu"](y), api["Emu"](w), api["Emu"](h)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = colors["white"]
    shape.line.color.rgb = colors["accent"] if accent else colors["ink"]
    shape.line.width = api["Pt"](1.75)
    _strip_style(shape, api["qn"])
    frame = shape.text_frame
    frame.word_wrap = kind != "diamond"
    _anchor(frame, api, "ctr")
    if kind == "diamond":
        _diamond_body(frame, api)
    paragraph = frame.paragraphs[0]
    paragraph.alignment = api["PP_ALIGN"].CENTER
    paragraph.text = label
    size = _diamond_font_size(label, w) if kind == "diamond" else 20
    _style_run(paragraph.runs[0], size, colors["ink"], api)
    return shape


def add_arrow(slide, x1, y1, x2, y2, api):
    colors = _colors(api)
    qn = api["qn"]
    vertical = abs(x2 - x1) < abs(y2 - y1)
    shape = slide.shapes.add_connector(
        api["MSO_CONNECTOR"].STRAIGHT,
        api["Emu"](x1), api["Emu"](y1), api["Emu"](x2), api["Emu"](y2),
    )
    shape.line.color.rgb = colors["ink"]
    shape.line.width = api["Pt"](2.25 if vertical else 1.5)
    line = shape.line._ln
    head = "lg" if vertical else "med"
    tail = line.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": head, "len": head})
    line.append(tail)
    _strip_style(shape, api["qn"])
    return shape


def _check_graph(nodes, edges):
    if not isinstance(nodes, list) or not nodes:
        raise ValueError("nodes must be a non-empty list")
    if len(nodes) > MAX_NODES:
        raise ValueError(f"At most {MAX_NODES} nodes on one slide")
    ids = []
    for node in nodes:
        if not isinstance(node, dict) or not node.get("id") or not node.get("label"):
            raise ValueError("Each node needs id and label")
        if node["id"] in ids:
            raise ValueError("Duplicate node id: " + node["id"])
        ids.append(node["id"])
    known = set(ids)
    for edge in edges or []:
        if edge.get("src") not in known or edge.get("dst") not in known:
            raise ValueError("Edge endpoint is not a node: " + str(edge))
    return ids


def _positions(nodes, region, layout):
    left, top, width, height = region
    if layout == "columns":
        groups = []
        index = {}
        for node in nodes:
            column = int(node.get("column", 0))
            if column not in index:
                index[column] = len(groups)
                groups.append([])
            groups[index[column]].append(node)
        gutter = emu(0.55) if len(groups) > 1 else 0
        col_w = (width - gutter * (len(groups) - 1)) // len(groups)
        positions = {}
        for i, group in enumerate(groups):
            x = left + i * (col_w + gutter)
            # Vertical connectors are the gap. 0.18in left only the arrowhead.
            gap_y = emu(0.50) if len(group) > 1 else 0
            node_h = (height - gap_y * (len(group) - 1)) // max(len(group), 1)
            if len(group) > 1 and node_h < emu(0.72):
                gap_y = max(emu(0.36), (height - emu(0.72) * len(group)) // (len(group) - 1))
                node_h = (height - gap_y * (len(group) - 1)) // len(group)
            y = top
            for node in group:
                positions[node["id"]] = (x, y, col_w, node_h)
                y += node_h + gap_y
        return positions
    if layout != "row":
        raise ValueError("layout must be row or columns")
    count = len(nodes)
    gap = emu(0.36) if count > 1 else 0
    node_w = (width - gap * (count - 1)) // count
    node_h = min(height, emu(0.92))
    y = top + max(0, (height - node_h) // 2)
    positions = {}
    x = left
    for node in nodes:
        positions[node["id"]] = (x, y, node_w, node_h)
        x += node_w + gap
    return positions


def _arrow_ends(src, dst):
    ax, ay, aw, ah = src
    bx, by, bw, bh = dst
    acx, acy = ax + aw / 2, ay + ah / 2
    bcx, bcy = bx + bw / 2, by + bh / 2
    if abs(acx - bcx) >= abs(acy - bcy):
        if acx <= bcx:
            return ax + aw, acy, bx, bcy
        return ax, acy, bx + bw, bcy
    if acy <= bcy:
        return acx, ay + ah, bcx, by
    return acx, ay, bcx, by + bh


def add_diagram(slide, nodes, edges, region, layout="row"):
    """Draw nodes and straight arrows. region is (x, y, w, h) in EMU."""
    api = _pptx()
    _check_graph(nodes, edges or [])
    positions = _positions(nodes, region, layout)
    for node in nodes:
        x, y, w, h = positions[node["id"]]
        add_node(
            slide,
            node.get("shape") or "round",
            x, y, w, h,
            node["label"],
            bool(node.get("accent")),
            api,
        )
    for edge in edges or []:
        x1, y1, x2, y2 = _arrow_ends(positions[edge["src"]], positions[edge["dst"]])
        add_arrow(slide, x1, y1, x2, y2, api)
    return positions


def _notes(slide, text):
    if text:
        slide.notes_slide.notes_text_frame.text = text


def populate_content_slide(slide, spec):
    """Title, diagram, lower-band conclusion, source line, page number."""
    api = _pptx()
    colors = _colors(api)
    title = spec.get("title")
    if title:
        add_text(slide, emu(0.7), emu(0.38), emu(11.6), emu(0.62), title, 32, colors["accent"], api)
    layout = spec.get("layout") or "row"
    region = spec.get("region")
    if region is None:
        region = (emu(0.7), emu(1.70), emu(11.9), emu(3.20)) if layout == "columns" else (
            emu(0.7), emu(1.95), emu(11.9), emu(1.15)
        )
    else:
        region = tuple(region)
    if spec.get("nodes"):
        add_diagram(slide, spec["nodes"], spec.get("edges") or [], region, layout)
    conclusion = spec.get("conclusion")
    if conclusion:
        add_text(
            slide, emu(0.7), emu(5.15), emu(11.9), emu(0.55),
            conclusion, 20, colors["ink"], api,
        )
    source = spec.get("source")
    if source:
        add_text(
            slide, emu(0.7), emu(6.55), emu(9.5), emu(0.32),
            source, 12, colors["muted"], api,
        )
    page = spec.get("page")
    if page:
        add_text(
            slide, emu(11.9), emu(6.55), emu(0.7), emu(0.32),
            str(page), 13, colors["muted"], api, align="center",
        )
    _notes(slide, spec.get("notes"))
    return slide


def render_math_png(path, expr=r"$\lambda(t)=\alpha e^{-\beta t}$"):
    """Write a high-DPI math image. The LaTeX string stays in this file, not on the slide."""
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
    except ImportError as exc:
        raise RuntimeError("matplotlib is required to render the sample formula image") from exc
    matplotlib.rcParams["mathtext.fontset"] = "cm"
    figure = plt.figure(figsize=(7.2, 1.35), dpi=400)
    figure.patch.set_facecolor("white")
    axis = figure.add_axes([0, 0, 1, 1])
    axis.set_axis_off()
    axis.text(0.02, 0.5, expr, fontsize=36, color="#222222", va="center", ha="left")
    figure.savefig(path, dpi=400, bbox_inches="tight", facecolor="white", pad_inches=0.08)
    plt.close(figure)
    return path


def _add_formula_slide(prs, png_path):
    api = _pptx()
    colors = _colors(api)
    slide = add_blank_slide(prs)
    add_text(
        slide, emu(0.7), emu(0.38), emu(11.6), emu(0.62),
        "公式：衰减项在回答什么", 32, colors["accent"], api,
    )
    slide.shapes.add_picture(str(png_path), api["Emu"](emu(0.75)), api["Emu"](emu(2.05)), api["Emu"](emu(6.5)))
    box = add_text(
        slide, emu(7.6), emu(2.05), emu(4.8), emu(2.2),
        "λ：速率", 22, colors["ink"], api,
    )
    frame = box.text_frame
    for line in ("α：起点", "β：衰减", "t：时间"):
        paragraph = frame.add_paragraph()
        paragraph.text = line
        _style_run(paragraph.runs[0], 22, colors["ink"], api)
    add_text(
        slide, emu(0.7), emu(5.15), emu(11.9), emu(0.55),
        "只解释符号，不代表实测。", 20, colors["ink"], api,
    )
    add_text(
        slide, emu(0.7), emu(6.55), emu(9.5), emu(0.32),
        "来源：匿名示意式，不是论文结果。", 12, colors["muted"], api,
    )
    add_text(slide, emu(11.9), emu(6.55), emu(0.7), emu(0.32), "03", 13, colors["muted"], api, align="center")
    _notes(
        slide,
        "先用图里的式子说明这一项在比较什么，再点右边四个符号。"
        "式子是匿名示意，不是论文里的似然。屏幕上不写 LaTeX 源码。",
    )
    return slide


def write_sample_deck(path):
    """Three anonymous slides: 4-step flow, two-column architecture, formula plus gloss."""
    prs = new_deck()
    slides = [
        {
            "title": "训练流程：四步示意",
            "layout": "row",
            "nodes": FLOW_NODES,
            "edges": FLOW_EDGES,
            "conclusion": "顺序是汇报归纳，不是原文算法。",
            "source": "来源：匿名示意，不是论文实验。",
            "page": "01",
            "notes": "用四个方框讲清顺序。这是汇报归纳，不要说成论文规定的算法。下半部那一句是结论，解释留在备注。",
        },
        {
            "title": "架构：编码端与解码端",
            "layout": "columns",
            "nodes": ARCH_NODES,
            "edges": ARCH_EDGES,
            "conclusion": "模块为示意，不代表论文原图。",
            "source": "来源：匿名结构，不是论文 Figure。",
            "page": "02",
            "notes": "左栏是编码，右栏是解码。箭头只表示讲解时的数据方向。没有原图，所以不用生成图代替。",
        },
    ]
    for spec in slides:
        populate_content_slide(add_blank_slide(prs), spec)
    with tempfile.TemporaryDirectory() as folder:
        png = Path(folder) / "formula.png"
        render_math_png(png)
        _add_formula_slide(prs, png)
        target = Path(path)
        target.parent.mkdir(parents=True, exist_ok=True)
        prs.save(str(target))
    return target


def build_slide_from_spec(spec, output):
    prs = new_deck()
    populate_content_slide(add_blank_slide(prs), spec)
    target = Path(output)
    target.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(target))
    return target


def _flow_slide_xml():
    """One figure-first slide as XML, for appending without rewriting other slides."""
    prs = new_deck()
    populate_content_slide(add_blank_slide(prs), {
        "title": "算法：用形状画出步骤",
        "layout": "row",
        "nodes": FLOW_NODES,
        "edges": FLOW_EDGES,
        "conclusion": "纯文字步骤不合格，缺图要写明。",
        "source": "来源：匿名示意，不是论文算法。",
        "page": "07",
    })
    with tempfile.TemporaryDirectory() as folder:
        path = Path(folder) / "one.pptx"
        prs.save(str(path))
        with zipfile.ZipFile(path) as package:
            xml = package.read("ppt/slides/slide1.xml")
    if b"r:embed" in xml:
        raise RuntimeError("Demo extra slide must not depend on a picture part")
    return xml


def _replace_notes_body(xml_bytes, text):
    ns = {
        "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
        "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    }
    root = ET.fromstring(xml_bytes)
    for shape in root.findall(".//p:sp", ns):
        ph = shape.find("p:nvSpPr/p:nvPr/p:ph", ns)
        if ph is not None and ph.get("type") == "body":
            node = shape.find(".//a:t", ns)
            if node is None:
                raise ValueError("Notes body has no text node")
            node.text = text
            break
    else:
        raise ValueError("Notes master body placeholder not found")
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def append_figure_page(demo_path):
    """Append one native-diagram slide. Existing slide parts are copied unchanged."""
    demo = Path(demo_path)
    with zipfile.ZipFile(demo) as package:
        names = set(package.namelist())
        if "ppt/slides/slide7.xml" in names:
            raise ValueError("slide7 already exists; refusing to append twice")
        original = {name: package.read(name) for name in package.namelist()}
    slide_xml = _flow_slide_xml()
    notes_xml = _replace_notes_body(
        original["ppt/notesSlides/notesSlide6.xml"],
        "本页演示图优先：算法步骤用可编辑的圆角矩形和箭头，不用纯文字列表。"
        "四步是匿名示意，不是论文算法。下半部的一句结论用来填满内容带，推导留在备注。",
    )
    slide_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="/ppt/slideLayouts/slideLayout6.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="/ppt/notesSlides/notesSlide7.xml"/></Relationships>
"""
    notes_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="/ppt/slides/slide7.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="/ppt/notesMasters/notesMaster1.xml"/></Relationships>
"""
    presentation = original["ppt/presentation.xml"].decode("utf-8")
    marker = "</p:sldIdLst>"
    if marker not in presentation:
        raise ValueError("presentation.xml has no slide list")
    presentation = presentation.replace(
        marker,
        '<p:sldId id="262" r:id="Rround3figpage07" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>'
        + marker,
        1,
    )
    rels = original["ppt/_rels/presentation.xml.rels"].decode("utf-8")
    rels = rels.replace(
        "</Relationships>",
        '<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="/ppt/slides/slide7.xml" Id="Rround3figpage07"/></Relationships>',
        1,
    )
    content_types = original["[Content_Types].xml"].decode("utf-8")
    content_types = content_types.replace(
        "</Types>",
        '<Override PartName="/ppt/slides/slide7.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        '<Override PartName="/ppt/notesSlides/notesSlide7.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>'
        "</Types>",
        1,
    )
    additions = {
        "ppt/slides/slide7.xml": slide_xml,
        "ppt/slides/_rels/slide7.xml.rels": slide_rels.encode("utf-8"),
        "ppt/notesSlides/notesSlide7.xml": notes_xml,
        "ppt/notesSlides/_rels/notesSlide7.xml.rels": notes_rels.encode("utf-8"),
        "ppt/presentation.xml": presentation.encode("utf-8"),
        "ppt/_rels/presentation.xml.rels": rels.encode("utf-8"),
        "[Content_Types].xml": content_types.encode("utf-8"),
    }
    replaced = dict(original)
    replaced.update(additions)
    tmp = demo.with_suffix(".pptx.tmp")
    with zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as package:
        for name, data in replaced.items():
            package.writestr(name, data)
    tmp.replace(demo)
    return demo


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--spec", help="JSON file describing one diagram slide")
    parser.add_argument("--output", help="Destination pptx for --spec")
    parser.add_argument("--samples", help="Write the 3-slide anonymous sample deck here")
    parser.add_argument("--append-demo", help="Append one figure-first slide onto this pptx")
    args = parser.parse_args(argv)
    try:
        if args.samples:
            write_sample_deck(args.samples)
        elif args.append_demo:
            append_figure_page(args.append_demo)
        elif args.spec and args.output:
            spec = json.loads(Path(args.spec).read_text(encoding="utf-8"))
            build_slide_from_spec(spec, args.output)
        else:
            parser.error("Pass --samples, --append-demo, or both --spec and --output")
    except (OSError, ValueError, RuntimeError, json.JSONDecodeError) as exc:
        print(type(exc).__name__ + ": " + str(exc), file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
