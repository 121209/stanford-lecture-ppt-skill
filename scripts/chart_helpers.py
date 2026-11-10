#!/usr/bin/env python3
"""Native editable charts for lecture slides.

Requires python-pptx. ``add_bar_chart`` is one series, no legend, no chart
title, data labels on, value axis hidden, and the baseline fixed at zero.
The highlighted bar is #C00000; the others are #BFBFBF.

``add_line_chart`` and ``add_scatter_chart`` are the same idea for a short
trend or a handful of points: one series, no legend, no chart title, axes
visible, gridlines off, series in #C00000. The caller passes every number.
These helpers do not invent categories, values, or points.

python-pptx can write negative ``c:axId`` / ``c:crossAx`` values. Those fail
Office schema checks (``xsd:unsignedInt``). Each helper remaps them to stable
positive ids before the package is saved.
"""
import argparse
import json
import math
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


def _finite(value, label):
    if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value):
        raise ValueError(label + " must be a finite number")
    return float(value)


def _series_numbers(values, label):
    if isinstance(values, (str, bytes)) or not isinstance(values, (list, tuple)) or len(values) < 2:
        raise ValueError(label + " must be a list of at least two numbers")
    if len(values) > 8:
        raise ValueError(label + " accepts at most 8 points so the axis stays readable")
    return [_finite(value, label) for value in values]


def _points(points):
    if isinstance(points, (str, bytes)) or not isinstance(points, (list, tuple)) or len(points) < 2:
        raise ValueError("scatter chart needs at least two (x, y) points")
    if len(points) > 12:
        raise ValueError("scatter chart accepts at most 12 points")
    parsed = []
    for item in points:
        if isinstance(item, (str, bytes)) or not isinstance(item, (list, tuple)) or len(item) != 2:
            raise ValueError("each scatter point must be (x, y)")
        parsed.append((_finite(item[0], "x"), _finite(item[1], "y")))
    return parsed


def _hide_chrome(chart):
    chart.has_title = False
    chart.has_legend = False


def _style_axis(axis, api, color):
    axis.has_major_gridlines = False
    axis.has_minor_gridlines = False
    axis.format.line.color.rgb = color
    labels = axis.tick_labels
    labels.font.size = api["Pt"](16)
    labels.font.name = FONT
    labels.font.color.rgb = color
    labels.font.bold = False


def _scale_from_zero(axis, numbers):
    """Pin a non-negative series to zero. Leave a negative series on auto scale."""
    if min(numbers) < 0:
        return
    peak = max(numbers)
    axis.minimum_scale = 0
    axis.maximum_scale = 1 if peak <= 0 else peak * 1.15


def _accent_marker(series, api):
    from pptx.enum.chart import XL_MARKER_STYLE

    series.marker.style = XL_MARKER_STYLE.CIRCLE
    series.marker.size = 8
    series.marker.format.fill.solid()
    series.marker.format.fill.fore_color.rgb = _rgb(api, ACCENT)
    series.marker.format.line.fill.solid()
    series.marker.format.line.color.rgb = _rgb(api, ACCENT)


def _base_chart(slide, chart_type, data, region, api):
    x, y, w, h = _region(region)
    frame = slide.shapes.add_chart(
        chart_type,
        api["Emu"](x), api["Emu"](y), api["Emu"](w), api["Emu"](h),
        data,
    )
    chart = frame.chart
    _hide_chrome(chart)
    chart.font.size = api["Pt"](16)
    chart.font.name = FONT
    chart.font.color.rgb = _rgb(api, INK)
    plot = chart.plots[0]
    if hasattr(plot._element, "_remove_dLbls"):
        plot.has_data_labels = False
    if hasattr(plot._element, "varyColors") or hasattr(plot._element, "_add_varyColors"):
        plot.vary_by_categories = False
    return frame, chart, plot, (x, y, w, h)


def add_line_chart(slide, categories, values, region, source=None):
    """Add one marked line. ``region`` is ``(x, y, w, h)`` in EMU.

    The line and its markers are #C00000. Both axes stay visible and start at
    zero when every value is non-negative. ``source``, when set, is a 13pt
    line under the chart. The caller supplies every number.
    """
    numbers = _series_numbers(values, "line chart values")
    labels = _categories(categories, len(numbers))
    api = _pptx()
    data = api["CategoryChartData"]()
    data.categories = labels
    data.add_series("value", tuple(numbers))
    frame, chart, _plot, (x, y, w, h) = _base_chart(
        slide, api["XL_CHART_TYPE"].LINE_MARKERS, data, region, api,
    )
    series = chart.series[0]
    series.format.line.fill.solid()
    series.format.line.color.rgb = _rgb(api, ACCENT)
    series.format.line.width = api["Pt"](2.25)
    _accent_marker(series, api)
    ink = _rgb(api, INK)
    muted = _rgb(api, MUTED)
    _style_axis(chart.category_axis, api, ink)
    chart.category_axis.format.line.color.rgb = muted
    _style_axis(chart.value_axis, api, ink)
    chart.value_axis.format.line.color.rgb = muted
    _scale_from_zero(chart.value_axis, numbers)
    repair_chart_axis_ids(chart)
    if isinstance(source, str) and source.strip():
        _source_line(slide, api, x, y + h + int(round(0.08 * 914400)), w, source.strip())
    return frame


def add_scatter_chart(slide, points, region, source=None):
    """Add one scatter series with no connecting line.

    ``points`` is a list of ``(x, y)`` numbers. Markers are #C00000. Both axes
    stay visible and start at zero when every coordinate is non-negative.
    """
    from pptx.chart.data import XyChartData

    parsed = _points(points)
    api = _pptx()
    data = XyChartData()
    series_data = data.add_series("value")
    for x_value, y_value in parsed:
        series_data.add_data_point(x_value, y_value)
    frame, chart, _plot, (x, y, w, h) = _base_chart(
        slide, api["XL_CHART_TYPE"].XY_SCATTER, data, region, api,
    )
    series = chart.series[0]
    series.format.line.fill.background()
    _accent_marker(series, api)
    ink = _rgb(api, INK)
    muted = _rgb(api, MUTED)
    _style_axis(chart.category_axis, api, ink)
    chart.category_axis.format.line.color.rgb = muted
    _style_axis(chart.value_axis, api, ink)
    chart.value_axis.format.line.color.rgb = muted
    _scale_from_zero(chart.category_axis, [item[0] for item in parsed])
    _scale_from_zero(chart.value_axis, [item[1] for item in parsed])
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
        kind = spec.get("kind") or "bar"
        source = spec.get("source")
        if kind == "bar":
            add_bar_chart(
                slide,
                spec["categories"],
                spec["values"],
                spec["region"],
                highlight=spec.get("highlight"),
                source=source,
            )
        elif kind == "line":
            add_line_chart(slide, spec["categories"], spec["values"], spec["region"], source=source)
        elif kind == "scatter":
            add_scatter_chart(slide, spec["points"], spec["region"], source=source)
        else:
            raise ValueError("kind must be bar, line, or scatter")
        target = Path(args.output)
        target.parent.mkdir(parents=True, exist_ok=True)
        prs.save(str(target))
    except (OSError, ValueError, RuntimeError, KeyError, json.JSONDecodeError) as exc:
        print(type(exc).__name__ + ": " + str(exc), file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
