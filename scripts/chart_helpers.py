#!/usr/bin/env python3
"""Native editable bar charts for result slides.

Requires python-pptx. One series, no legend, no chart title, data labels on,
value axis hidden, and the baseline fixed at zero. The highlighted bar is
#C00000; the others are #BFBFBF. Numbers are whatever the caller passes.
This module does not invent categories or values.

python-pptx can write negative ``c:axId`` / ``c:crossAx`` values. Those fail
Office schema checks (``xsd:unsignedInt``). ``add_bar_chart`` remaps them to
stable positive ids before the package is saved.
"""
import argparse
import json
import sys
from pathlib import Path

INK = (0x22, 0x22, 0x22)
ACCENT = (0xC0, 0x00, 0x00)
MUTED = (0x60, 0x60, 0x60)
OTHER = (0xBF, 0xBF, 0xBF)
FONT = "Microsoft YaHei"
# First remapped axis id. Unique inside one chart; axes only need to agree
# with each other, not with other charts.
_AXIS_ID_BASE = 100000


def _rgb(api, triple):
    return api["RGBColor"](*triple)


def _pptx():
    try:
        import pptx
        from pptx.chart.data import CategoryChartData
        from pptx.dml.color import RGBColor
        from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
        from pptx.enum.text import PP_ALIGN
        from pptx.oxml.ns import qn
        from pptx.util import Emu, Pt
    except ImportError as exc:
        raise RuntimeError("python-pptx is required for chart_helpers.py") from exc
    return {
        "Presentation": pptx.Presentation,
        "CategoryChartData": CategoryChartData,
        "RGBColor": RGBColor,
        "XL_CHART_TYPE": XL_CHART_TYPE,
        "XL_LABEL_POSITION": XL_LABEL_POSITION,
        "PP_ALIGN": PP_ALIGN,
        "qn": qn,
        "Emu": Emu,
        "Pt": Pt,
    }


def repair_chart_axis_ids(chart):
    """Rewrite ``c:axId`` and ``c:crossAx`` to positive ids.

    python-pptx samples signed 32-bit ids, and the negatives fail
    ``xsd:unsignedInt``. The same original value always maps to the same new
    value, so a category axis and a value axis still name each other.
    """
    element = chart.element
    mapping = {}
    for node in element.iter():
        if not isinstance(node.tag, str):
            continue
        local = node.tag.rsplit("}", 1)[-1]
        if local not in {"axId", "crossAx"} or "val" not in node.attrib:
            continue
        raw = node.get("val")
        if raw not in mapping:
            mapping[raw] = str(_AXIS_ID_BASE + len(mapping))
        node.set("val", mapping[raw])
    return mapping


def _numbers(values):
    if isinstance(values, (str, bytes)) or not isinstance(values, (list, tuple)) or not values:
        raise ValueError("values must be a non-empty list of numbers")
    numbers = []
    for value in values:
        if isinstance(value, bool) or not isinstance(value, (int, float)):
            raise ValueError("bar chart values must be numbers")
        if value < 0:
            raise ValueError("bar chart values must be non-negative so the baseline stays at zero")
        numbers.append(float(value))
    return numbers


def _categories(categories, count):
    if isinstance(categories, (str, bytes)) or not isinstance(categories, (list, tuple)):
        raise ValueError("categories must be a list of strings")
    if len(categories) != count:
        raise ValueError("categories and values must be the same length")
    labels = []
    for label in categories:
        if not isinstance(label, str) or not label.strip():
            raise ValueError("each category must be a non-empty string")
        labels.append(label)
    return labels


def _region(region):
    if not isinstance(region, (list, tuple)) or len(region) != 4:
        raise ValueError("region must be (x, y, w, h) in EMU")
    x, y, w, h = (int(value) for value in region)
    if w <= 0 or h <= 0:
        raise ValueError("region width and height must be positive")
    return x, y, w, h


def _slide_size(slide):
    presentation = slide.part.package.presentation_part.presentation
    return int(presentation.slide_width), int(presentation.slide_height)


def _source_line(slide, api, x, y, w, text):
    _width, height = _slide_size(slide)
    box_h = int(round(0.32 * 914400))
    if y + box_h > height:
        y = max(0, height - box_h)
    box = slide.shapes.add_textbox(api["Emu"](x), api["Emu"](y), api["Emu"](w), api["Emu"](box_h))
    frame = box.text_frame
    frame.word_wrap = True
    paragraph = frame.paragraphs[0]
    paragraph.alignment = api["PP_ALIGN"].LEFT
    run = paragraph.add_run()
    run.text = text
    run.font.size = api["Pt"](13)
    run.font.color.rgb = _rgb(api, MUTED)
    run.font.name = FONT
    return box


def add_bar_chart(slide, categories, values, region, highlight=None, source=None):
    """Add one clustered bar chart. ``region`` is ``(x, y, w, h)`` in EMU.

    ``highlight`` is the index of the accent bar. ``source``, when set, is a
    13pt line under the chart. The caller supplies every number.
    """
    if highlight is not None and (isinstance(highlight, bool) or not isinstance(highlight, int)):
        raise ValueError("highlight must be an index")
    numbers = _numbers(values)
    labels = _categories(categories, len(numbers))
    if highlight is not None and not 0 <= highlight < len(numbers):
        raise ValueError("highlight index is outside the series")
    x, y, w, h = _region(region)
    api = _pptx()
    data = api["CategoryChartData"]()
    data.categories = labels
    # The series name is not shown: the chart has no legend.
    data.add_series("value", tuple(numbers))
    frame = slide.shapes.add_chart(
        api["XL_CHART_TYPE"].BAR_CLUSTERED,
        api["Emu"](x), api["Emu"](y), api["Emu"](w), api["Emu"](h),
        data,
    )
    chart = frame.chart
    chart.has_title = False
    chart.has_legend = False
    chart.font.size = api["Pt"](16)
    chart.font.name = FONT
    chart.font.color.rgb = _rgb(api, INK)
    plot = chart.plots[0]
    plot.gap_width = 60
    plot.has_data_labels = True
    labels_fmt = plot.data_labels
    labels_fmt.number_format = "General"
    labels_fmt.number_format_is_linked = False
    labels_fmt.position = api["XL_LABEL_POSITION"].OUTSIDE_END
    labels_fmt.font.size = api["Pt"](16)
    labels_fmt.font.name = FONT
    labels_fmt.font.color.rgb = _rgb(api, INK)
    series = plot.series[0]
    for index, point in enumerate(series.points):
        point.format.fill.solid()
        color = ACCENT if index == highlight else OTHER
        point.format.fill.fore_color.rgb = _rgb(api, color)
    value_axis = chart.value_axis
    value_axis.minimum_scale = 0
    peak = max(numbers) if numbers else 0
    value_axis.maximum_scale = 1 if peak <= 0 else peak * 1.15
    value_axis.has_major_gridlines = False
    value_axis.visible = False
    category_axis = chart.category_axis
    category_axis.has_major_gridlines = False
    category_axis.format.line.color.rgb = _rgb(api, MUTED)
    category_axis.tick_labels.font.size = api["Pt"](16)
    category_axis.tick_labels.font.name = FONT
    category_axis.tick_labels.font.color.rgb = _rgb(api, INK)
    repair_chart_axis_ids(chart)
    if isinstance(source, str) and source.strip():
        _source_line(slide, api, x, y + h + int(round(0.08 * 914400)), w, source.strip())
    return frame


def _blank_deck(api):
    prs = api["Presentation"]()
    prs.slide_width = api["Emu"](12192000)
    prs.slide_height = api["Emu"](6858000)
    layout = prs.slide_layouts[len(prs.slide_layouts) - 1]
    for candidate in prs.slide_layouts:
        if (candidate.name or "").lower() == "blank":
            layout = candidate
            break
    return prs, prs.slides.add_slide(layout)


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--spec", help="JSON with categories, values, region, optional highlight and source")
    parser.add_argument("--output", help="Destination pptx")
    args = parser.parse_args(argv)
    try:
        if not args.spec or not args.output:
            parser.error("Pass --spec and --output")
        spec = json.loads(Path(args.spec).read_text(encoding="utf-8"))
        api = _pptx()
        prs, slide = _blank_deck(api)
        add_bar_chart(
            slide,
            spec["categories"],
            spec["values"],
            spec["region"],
            highlight=spec.get("highlight"),
            source=spec.get("source"),
        )
        target = Path(args.output)
        target.parent.mkdir(parents=True, exist_ok=True)
        prs.save(str(target))
    except (OSError, ValueError, RuntimeError, KeyError, json.JSONDecodeError) as exc:
        print(type(exc).__name__ + ": " + str(exc), file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
