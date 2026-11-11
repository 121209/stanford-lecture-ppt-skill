"""Helpers that add slide content must leave a preexisting sentinel in place.

Render checks run only when LibreOffice is installed. Without it, identity is
the slide XML: name, shape id, text, and position.
"""
import importlib.util
import shutil
import subprocess
import tempfile
import unittest
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SENTINEL = "PREEXISTING_SENTINEL"


def load(name):
    spec = importlib.util.spec_from_file_location(name, ROOT / "scripts" / f"{name}.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def load_if_present(name):
    path = ROOT / "scripts" / f"{name}.py"
    if not path.is_file():
        return None
    return load(name)


diagrams = load("diagram_helpers")
charts = load("chart_helpers")
embed = load("embed_motion")
inject = load("inject_omml")
inline = load("inline_math")
safe = load("safe_edit")


def _parents(root):
    found = {}
    for parent in root.iter():
        for child in list(parent):
            found[child] = parent
    return found


def fingerprint(path, slide_part="ppt/slides/slide1.xml"):
    with zipfile.ZipFile(path) as package:
        root = ET.fromstring(package.read(slide_part))
    parents = _parents(root)
    nodes = [
        element for element in root.iter()
        if isinstance(element.tag, str)
        and element.tag.endswith("}cNvPr")
        and element.get("name") == SENTINEL
    ]
    if len(nodes) != 1:
        raise AssertionError("expected one %s shape, found %d" % (SENTINEL, len(nodes)))
    node = nodes[0]
    shape = node
    while shape is not None and not str(shape.tag).endswith("}sp"):
        shape = parents.get(shape)
    if shape is None:
        raise AssertionError("sentinel is not a p:sp")
    texts = []
    off = ext = None
    for element in shape.iter():
        if not isinstance(element.tag, str):
            continue
        local = element.tag.rsplit("}", 1)[-1]
        if local == "t" and element.text:
            texts.append(element.text)
        elif local == "off" and off is None:
            off = (element.get("x"), element.get("y"))
        elif local == "ext" and ext is None:
            ext = (element.get("cx"), element.get("cy"))
    return {
        "id": node.get("id"),
        "name": node.get("name"),
        "hidden": node.get("hidden"),
        "text": "".join(texts),
        "off": off,
        "ext": ext,
    }


def plant(path, placeholder=None):
    from pptx import Presentation
    from pptx.util import Emu

    presentation = Presentation()
    presentation.slide_width = Emu(diagrams.SLIDE_W)
    presentation.slide_height = Emu(diagrams.SLIDE_H)
    layout = presentation.slide_layouts[len(presentation.slide_layouts) - 1]
    for candidate in presentation.slide_layouts:
        if (candidate.name or "").lower() == "blank":
            layout = candidate
            break
    slide = presentation.slides.add_slide(layout)
    box = slide.shapes.add_textbox(Emu(9000000), Emu(180000), Emu(2500000), Emu(360000))
    box.name = SENTINEL
    box.text_frame.paragraphs[0].text = SENTINEL
    if placeholder:
        formula = slide.shapes.add_textbox(Emu(640080), Emu(1600200), Emu(5486400), Emu(731520))
        formula.text_frame.paragraphs[0].text = "{{MATH:%s}}" % placeholder
    path.parent.mkdir(parents=True, exist_ok=True)
    presentation.save(str(path))
    found = fingerprint(path)
    if found["text"] != SENTINEL or found["name"] != SENTINEL:
        raise AssertionError("failed to plant sentinel: %r" % found)
    if found["hidden"] not in (None, "0"):
        raise AssertionError("planted sentinel is hidden: %r" % found)
    return found


class PreserveExistingTests(unittest.TestCase):
    def setUp(self):
        try:
            import pptx  # noqa: F401
        except ImportError:
            self.skipTest("python-pptx is not installed")

    def _exercise(self, produce, placeholder=None):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / "source.pptx"
            before = plant(source, placeholder=placeholder)
            source_bytes = source.read_bytes()
            opened = safe.prepare_working_copy(source, dest=root / "work.pptx")
            self.assertFalse(opened["in_place"])
            self.assertEqual(source.read_bytes(), source_bytes)
            edited = produce(Path(opened["working_copy"]), root)
            after = fingerprint(edited)
            self.assertEqual(after, before)
            self.assertEqual(after["text"], SENTINEL)
            self.assertEqual(after["name"], SENTINEL)
            self.assertIsNotNone(after["id"])
            self.assertIsNotNone(after["off"])
            self.assertIsNotNone(after["ext"])
            self.assertIn(after["hidden"], (None, "0"))
            self.assertNotEqual(edited.read_bytes(), source_bytes)
            self.assertEqual(source.read_bytes(), source_bytes)
        self.assertFalse(root.exists())

    def test_diagram_helpers_keep_the_sentinel(self):
        def produce(path, _folder):
            from pptx import Presentation

            presentation = Presentation(str(path))
            slide = presentation.slides[0]
            api = diagrams._pptx()
            ink = api["RGBColor"](0x22, 0x22, 0x22)
            diagrams.add_text(
                slide, diagrams.emu(0.7), diagrams.emu(1.3),
                diagrams.emu(4.0), diagrams.emu(0.45),
                "新增说明", 20, ink, api,
            )
            diagrams.add_node(
                slide, "round", diagrams.emu(0.7), diagrams.emu(2.0),
                diagrams.emu(1.8), diagrams.emu(0.72), "节点", False, api,
            )
            diagrams.add_arrow(
                slide, diagrams.emu(2.6), diagrams.emu(2.36),
                diagrams.emu(4.2), diagrams.emu(2.36), api,
            )
            diagrams.add_diagram(
                slide,
                [{"id": "a", "label": "甲"}, {"id": "b", "label": "乙"}],
                [{"src": "a", "dst": "b"}],
                (diagrams.emu(0.7), diagrams.emu(3.1), diagrams.emu(6.0), diagrams.emu(1.0)),
            )
            diagrams.add_bar_chart(
                slide, ["甲", "乙"], [1, 2],
                (diagrams.emu(7.0), diagrams.emu(3.1), diagrams.emu(5.0), diagrams.emu(2.4)),
                highlight=0,
            )
            diagrams.populate_content_slide(slide, {
                "title": "补充：仍保留原对象",
                "layout": "row",
                "nodes": [{"id": "one", "label": "一步"}, {"id": "two", "label": "二步"}],
                "edges": [{"src": "one", "dst": "two"}],
                "conclusion": "这是新增结论，不是替换原形状。",
                "source": "来源：测试，不是论文。",
                "page": "9",
                "notes": "备注写在备注页，不删幻灯片上的哨兵。",
            })
            presentation.save(str(path))
            return path

        self._exercise(produce)

    def test_chart_helpers_keep_the_sentinel(self):
        def produce(path, _folder):
            from pptx import Presentation

            presentation = Presentation(str(path))
            charts.add_bar_chart(
                presentation.slides[0],
                ["甲", "乙"],
                [1.0, 2.0],
                (700000, 1600000, 6000000, 2800000),
                highlight=1,
                source="来源：测试数值，不是论文结果。",
            )
            presentation.save(str(path))
            return path

        self._exercise(produce)

    def test_embed_motion_keeps_the_sentinel(self):
        def produce(path, folder):
            poster = folder / "poster.png"
            poster.write_bytes(embed._white_png())
            dest = folder / "motion.pptx"
            embed.embed_motion(
                dest, poster=poster, deck=path, slide=1, caption="新增图注，不是原形状。",
            )
            return dest

        self._exercise(produce)

    def test_inject_omml_keeps_the_sentinel(self):
        try:
            import matplotlib  # noqa: F401
        except ImportError:
            self.skipTest("matplotlib is not installed")

        def produce(path, folder):
            dest = folder / "omml.pptx"
            inject.inject_pptx(path, {
                "rate": {"runs": [{"t": "λ"}, {"t": "="}, {"sup": {"base": "e", "sup": "−t"}}]},
            }, dest)
            return dest

        self._exercise(produce, placeholder="rate")

    def test_inline_math_keeps_the_sentinel(self):
        try:
            import matplotlib  # noqa: F401
        except ImportError:
            self.skipTest("matplotlib is not installed")

        def produce(path, folder):
            dest = folder / "inline.pptx"
            inline.inject_inline_pptx(path, {"dk": "d_k"}, dest)
            return dest

        self._exercise(produce, placeholder="dk")

    def test_layout_helpers_keep_the_sentinel_when_present(self):
        layout = load_if_present("layout_helpers")
        if layout is None:
            self.skipTest("layout_helpers.py is not on this branch")

        def produce(path, _folder):
            from pptx import Presentation

            presentation = Presentation(str(path))
            layout.add_text(
                presentation.slides[0],
                layout.emu(0.7), layout.emu(1.4), layout.emu(4.0), layout.emu(0.5),
                "版式新增", 20, (0x22, 0x22, 0x22),
            )
            presentation.save(str(path))
            return path

        self._exercise(produce)

    def test_appear_keeps_the_sentinel_when_present(self):
        appear = load_if_present("appear")
        if appear is None:
            self.skipTest("appear.py is not on this branch")

        def produce(path, _folder):
            from pptx import Presentation

            presentation = Presentation(str(path))
            slide = presentation.slides[0]
            api = diagrams._pptx()
            added = diagrams.add_text(
                slide, diagrams.emu(0.7), diagrams.emu(1.4),
                diagrams.emu(4.0), diagrams.emu(0.5),
                "单击进入", 20, api["RGBColor"](0x22, 0x22, 0x22), api,
            )
            appear.add_click_appear(slide, [added], effect="appear")
            presentation.save(str(path))
            return path

        self._exercise(produce)

    def test_libreoffice_render_is_optional(self):
        soffice = shutil.which("soffice") or shutil.which("libreoffice")
        if not soffice:
            self.skipTest("LibreOffice is not installed")
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / "source.pptx"
            before = plant(source)
            out = root / "render"
            out.mkdir()
            completed = subprocess.run(
                [
                    soffice, "--headless", "--norestore", "--nolockcheck", "--nologo",
                    "-env:UserInstallation=file://%s" % (out / "profile"),
                    "--convert-to", "pdf", "--outdir", str(out), str(source),
                ],
                capture_output=True, text=True, timeout=120,
            )
            self.assertEqual(completed.returncode, 0, completed.stderr[-400:])
            self.assertTrue(next(out.glob("*.pdf")).is_file())
            self.assertEqual(fingerprint(source), before)


if __name__ == "__main__":
    unittest.main()
