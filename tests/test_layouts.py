"""Layout, table, chart, and click-appear helpers. Fixtures are synthetic."""
import importlib.util
import sys
import tempfile
import unittest
import zipfile
from contextlib import redirect_stdout
from io import StringIO
from pathlib import Path
from xml.etree import ElementTree as ET

import test_helpers

ROOT = Path(__file__).resolve().parents[1]
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"


def load(name):
    spec = importlib.util.spec_from_file_location(name, ROOT / "scripts" / ("%s.py" % name))
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


layout = load("layout_helpers")
charts = load("chart_helpers")
appear = load("appear")
diagrams = load("diagram_helpers")
audit = test_helpers.audit


def _slide_xml(path, index=1):
    with zipfile.ZipFile(path) as package:
        return package.read("ppt/slides/slide%d.xml" % index)


def _chart_xml(path):
    with zipfile.ZipFile(path) as package:
        return package.read("ppt/charts/chart1.xml")


class LayoutHelperTests(unittest.TestCase):
    def setUp(self):
        try:
            import pptx  # noqa: F401
        except ImportError:
            self.skipTest("python-pptx is not installed")

    def _deck(self):
        prs = layout.new_deck()
        return prs, layout.add_blank_slide(prs)

    def _save(self, prs, folder, name="deck.pptx"):
        path = Path(folder) / name
        prs.save(str(path))
        return path

    def test_table_is_native_and_upper_band_needs_a_conclusion(self):
        prs, slide = self._deck()
        region = (layout.emu(0.7), layout.emu(1.8), layout.emu(10.8), layout.emu(4.5))
        with self.assertRaises(ValueError):
            layout.add_table(slide, ["组别", "甲"], [["一组", 2]], region)
        with self.assertRaises(ValueError):
            layout.add_table(slide, ["组别", "甲"], [["一组"]], region, conclusion="缺一列。")
        layout.add_title(slide, "两组示意整数")
        layout.add_table(
            slide,
            ["组别", "甲"],
            [["一组", 2], ["二组", 4]],
            (layout.emu(0.7), layout.emu(1.8), layout.emu(10.8), layout.emu(2.2)),
            caption="列是占位。",
            conclusion="大的格子只是占位。",
            source="来源：手写整数，不是实验。",
            page=3,
        )
        with tempfile.TemporaryDirectory() as folder:
            path = self._save(prs, folder)
            root = ET.fromstring(_slide_xml(path))
            cells = list(root.iter("{%s}tc" % A_NS))
            self.assertEqual(len(cells), 6)
            texts = ["".join(node.itertext()) for node in cells]
            self.assertEqual(texts, ["组别", "甲", "一组", "2", "二组", "4"])
            for cell in cells:
                sizes = [int(node.get("sz")) for node in cell.iter("{%s}rPr" % A_NS) if node.get("sz")]
                self.assertTrue(sizes)
                self.assertGreaterEqual(min(sizes), 2000)
            xml = ET.tostring(root, encoding="unicode")
            self.assertIn("大的格子只是占位。", xml)
            self.assertIn("来源：手写整数，不是实验。", xml)
            self.assertIn(">03<", xml)
            self.assertNotIn("outerShdw", xml)
            report = audit.audit(path)
            self.assertEqual(report["slides"][0]["native_table_count"], 1)
            self.assertEqual(report["slides"][0]["warnings"], [])

    def test_cover_comparison_callout_picture_and_section(self):
        prs = layout.new_deck()
        cover = layout.add_blank_slide(prs)
        layout.add_cover(
            cover,
            title="示意题目：匿名步骤比较",
            kicker="论文精读",
            authors="甲作者，乙作者",
            venue="匿名会议，2024",
            presenter="汇报人：示意",
            date="日期：示意",
        )
        section = layout.add_blank_slide(prs)
        layout.add_section_divider(section, "这一节只讲设定", kicker="第二节", page=2)
        columns = layout.add_blank_slide(prs)
        layout.add_two_column(
            columns,
            "论文工作和笔记分开",
            "论文工作",
            ["只定义步骤"],
            "本次笔记",
            ["步骤不是结论"],
            "右栏是笔记，不是已证。",
            source="来源：匿名笔记，不是论文原文。",
            page=4,
        )
        with self.assertRaises(ValueError):
            layout.add_two_column(
                columns, "再来一栏", "左", ["一", "二", "三", "四"], "右", ["一"], "太多。",
            )
        figure = layout.add_blank_slide(prs)
        layout.add_title(figure, "红线只是占位图")
        layout.add_picture_caption(
            figure,
            layout._synthetic_mark_png(),
            (layout.emu(0.8), layout.emu(1.7), layout.emu(7.4), layout.emu(3.25)),
            "图注可选中，不是数据。",
        )
        definition = layout.add_blank_slide(prs)
        layout.add_title(definition, "定义放在细框里")
        layout.add_callout(
            definition,
            (layout.emu(0.7), layout.emu(1.85), layout.emu(11.6), layout.emu(1.7)),
            "示意项",
            "没有测量含义。",
            read="框线不是卡片，也没有阴影。",
        )
        with self.assertRaises(ValueError):
            layout.format_page("3 / 9")
        with self.assertRaises(ValueError):
            layout.add_footer(cover, source="没有标记的说明")
        with tempfile.TemporaryDirectory() as folder:
            path = self._save(prs, folder)
            cover_xml = _slide_xml(path, 1).decode("utf-8")
            self.assertIn("甲作者，乙作者", cover_xml)
            self.assertIn("匿名会议，2024", cover_xml)
            self.assertIn('sz="2000"', cover_xml)
            self.assertIn('sz="3200"', cover_xml)
            report = audit.audit(path)
            self.assertEqual(report["slides"][1]["title_text"], "这一节只讲设定")
            self.assertEqual(report["slides"][3]["picture_count"], 1)
            self.assertIn("图注可选中，不是数据。", report["slides"][3]["native_texts"][1]["text"])
            callout = _slide_xml(path, 5).decode("utf-8")
            self.assertNotIn("outerShdw", callout)
            self.assertNotIn("gradFill", callout)
            self.assertIn("C00000", callout)
            self.assertIn("FFFFFF", callout)
            for slide in report["slides"]:
                self.assertEqual(slide["warnings"], [])

    def test_line_and_scatter_are_editable_charts(self):
        try:
            from lxml import etree
        except ImportError:
            self.skipTest("lxml is not installed")
        schema_dir = test_helpers._ooxml_schema_dir()
        if schema_dir is None:
            self.skipTest("OOXML schemas are not present")
        chart_schema = etree.XMLSchema(etree.parse(str(schema_dir / "dml-chart.xsd")))
        slide_schema = etree.XMLSchema(etree.parse(str(schema_dir / "pml.xsd")))
        prs, slide = self._deck()
        charts.add_line_chart(
            slide, ["甲", "乙", "丙", "丁"], [2, 5, 9, 9],
            (layout.emu(0.8), layout.emu(1.7), layout.emu(11.4), layout.emu(3.15)),
            source="来源：手写整数，不是测量。",
        )
        with tempfile.TemporaryDirectory() as folder:
            path = self._save(prs, folder, "line.pptx")
            chart_xml = _chart_xml(path)
            slide_xml = _slide_xml(path)
            text = chart_xml.decode("utf-8")
            self.assertIn("<c:lineChart", text)
            self.assertNotIn("<c:barChart", text)
            self.assertNotIn("<c:legend", text)
            self.assertNotIn("<c:title", text)
            self.assertIn(">2.0<", text)
            self.assertIn(">9.0<", text)
            self.assertIn("C00000", text)
            self.assertNotRegex(text, r'<c:(?:axId|crossAx) val="-\d+"')
            self.assertTrue(chart_schema.validate(etree.fromstring(chart_xml)), chart_schema.error_log)
            self.assertTrue(slide_schema.validate(etree.fromstring(slide_xml)), slide_schema.error_log)
        prs, slide = self._deck()
        diagrams.add_scatter_chart(
            slide, [(1, 2), (2, 8), (3, 3), (4, 4)],
            (layout.emu(0.8), layout.emu(1.7), layout.emu(11.4), layout.emu(3.15)),
        )
        with tempfile.TemporaryDirectory() as folder:
            path = self._save(prs, folder, "scatter.pptx")
            text = _chart_xml(path).decode("utf-8")
            self.assertIn("<c:scatterChart", text)
            self.assertNotIn("<c:lineChart", text)
            self.assertIn(">8.0<", text)
            self.assertIn("<a:noFill", text)
            self.assertIn("C00000", text)
            self.assertTrue(chart_schema.validate(etree.fromstring(text.encode("utf-8"))), chart_schema.error_log)
        prs, slide = self._deck()
        charts.add_line_chart(
            slide, ["甲", "乙"], [1, -2],
            (layout.emu(0.8), layout.emu(1.7), layout.emu(8), layout.emu(3)),
        )
        with tempfile.TemporaryDirectory() as folder:
            path = self._save(prs, folder, "signed.pptx")
            text = _chart_xml(path).decode("utf-8")
            self.assertIn(">-2.0<", text)
            self.assertNotIn('<c:min val="0"', text)
        with self.assertRaises(ValueError):
            charts.add_line_chart(slide, ["甲"], [1], (0, 0, 100, 100))
        with self.assertRaises(ValueError):
            charts.add_scatter_chart(slide, [(1, 2)], (0, 0, 100, 100))

    def test_click_appear_is_opt_in_and_schema_legal(self):
        try:
            from lxml import etree
        except ImportError:
            self.skipTest("lxml is not installed")
        schema_dir = test_helpers._ooxml_schema_dir()
        if schema_dir is None:
            self.skipTest("OOXML schemas are not present")
        slide_schema = etree.XMLSchema(etree.parse(str(schema_dir / "pml.xsd")))
        prs, slide = self._deck()
        first = layout.add_text(slide, layout.emu(0.7), layout.emu(1.9), layout.emu(8), layout.emu(0.5), "第一下出现这行", 20, layout.INK)
        second = layout.add_text(slide, layout.emu(0.7), layout.emu(2.7), layout.emu(8), layout.emu(0.5), "第二下出现这行", 20, layout.INK)
        self.assertIsNone(slide._element.find("{%s}timing" % P_NS))
        appear.add_click_appear(slide, [first, second])
        with self.assertRaises(ValueError):
            appear.add_click_appear(slide, [first])
        with tempfile.TemporaryDirectory() as folder:
            path = self._save(prs, folder, "appear.pptx")
            xml = _slide_xml(path)
            self.assertTrue(slide_schema.validate(etree.fromstring(xml)), slide_schema.error_log)
            text = xml.decode("utf-8")
            self.assertEqual(text.count('nodeType="clickEffect"'), 2)
            self.assertEqual(text.count('delay="indefinite"'), 2)
            self.assertNotIn("withEffect", text)
            self.assertIn('presetID="1"', text)
            self.assertIn("style.visibility", text)
            self.assertNotIn("<p:transition", text)
            self.assertNotIn("video", text)
        prs, slide = self._deck()
        first = layout.add_text(slide, layout.emu(0.7), layout.emu(1.9), layout.emu(8), layout.emu(0.5), "一起出现", 20, layout.INK)
        second = layout.add_text(slide, layout.emu(0.7), layout.emu(2.7), layout.emu(8), layout.emu(0.5), "同一次单击", 20, layout.INK)
        other = layout.add_blank_slide(prs)
        outsider = layout.add_text(other, layout.emu(0.7), layout.emu(1.9), layout.emu(8), layout.emu(0.5), "另一页", 20, layout.INK)
        with self.assertRaises(ValueError):
            appear.add_click_appear(slide, [outsider])
        with self.assertRaises(ValueError):
            appear.add_click_appear(slide, [first, second], effect="fade", duration_ms=400)
        appear.add_click_appear(slide, [first, second], effect="fade", together=True, duration_ms=600)
        with tempfile.TemporaryDirectory() as folder:
            path = self._save(prs, folder, "fade.pptx")
            xml = _slide_xml(path)
            self.assertTrue(slide_schema.validate(etree.fromstring(xml)), slide_schema.error_log)
            text = xml.decode("utf-8")
            self.assertEqual(text.count('nodeType="clickEffect"'), 1)
            self.assertEqual(text.count('nodeType="withEffect"'), 1)
            self.assertEqual(text.count('delay="indefinite"'), 1)
            self.assertIn('presetID="10"', text)
            self.assertIn('filter="fade"', text)
            self.assertIn('dur="600"', text)

    def test_sample_deck_stays_inside_the_lecture_rules(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "basics.pptx"
            self.assertEqual(layout.main(["--samples", str(path)]), 0)
            assert_sample(self, path)


class CheckedInLayoutSampleTests(unittest.TestCase):
    def test_checked_in_sample_audits_clean(self):
        assert_sample(self, ROOT / "assets" / "layout-samples" / "basics.pptx")


def assert_sample(case, path):
    report = audit.audit(path)
    case.assertEqual(report["slide_count"], 9)
    case.assertEqual(report["slides_over_word_cap"], [])
    case.assertEqual(report["slides_with_empty_lower_content_band"], [])
    case.assertEqual(report["slides_missing_diagram"], [])
    case.assertEqual(report["slides_with_raw_latex"], [])
    case.assertEqual(report["slides_with_font_below_min"], [])
    case.assertEqual(report["slides_with_low_contrast_text"], [])
    case.assertEqual(report["slides_without_body_notes"], [])
    case.assertEqual(report["slides"][2]["native_table_count"], 1)
    case.assertEqual(report["slides"][4]["native_chart_count"], 1)
    case.assertEqual(report["slides"][5]["native_chart_count"], 1)
    case.assertEqual(report["slides"][6]["picture_count"], 1)
    case.assertEqual(report["slides"][1]["title_text"], "这一节只讲设定")
    with zipfile.ZipFile(path) as package:
        names = package.namelist()
        case.assertFalse(any(name.endswith(".mp4") for name in names))
        timed = []
        for index in range(1, 10):
            xml = package.read("ppt/slides/slide%d.xml" % index)
            if b"timing" in xml:
                timed.append(index)
        case.assertEqual(timed, [9])
        last = package.read("ppt/slides/slide9.xml").decode("utf-8")
        case.assertEqual(last.count('nodeType="clickEffect"'), 3)
        case.assertNotIn("<p:transition", last)
    saved = sys.argv
    try:
        sys.argv = ["audit_pptx.py", str(path)]
        with redirect_stdout(StringIO()):
            code = audit.main()
    finally:
        sys.argv = saved
    case.assertEqual(code, 0)


if __name__ == "__main__":
    unittest.main()
