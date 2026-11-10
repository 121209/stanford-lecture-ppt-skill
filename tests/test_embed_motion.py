"""ManimGX stays optional: command planning and PPTX embed do not need the tool."""
import importlib.util
import json
import sys
import tempfile
import unittest
import zipfile
from contextlib import redirect_stdout
from io import StringIO
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def load(name):
    spec = importlib.util.spec_from_file_location(name, ROOT / "scripts" / f"{name}.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


embed = load("embed_motion")
audit = load("audit_pptx")

NS = {
    "p": audit.NS["p"],
    "a": audit.NS["a"],
    "r": audit.NS["r"],
    "rel": audit.NS["rel"],
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
}
VIDEO_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"
MEDIA_REL = "http://schemas.microsoft.com/office/2007/relationships/media"
# Not a playable film. python-pptx stores the bytes; it does not decode them.
MP4_STUB = b"\x00\x00\x00\x18ftypisom\x00\x00\x00\x00isomiso2mp41\x00\x00\x00\x08mdat"


def _png():
    return embed._white_png(8, 8)


class MotionCommandTests(unittest.TestCase):
    def test_probe_missing_is_not_an_error(self):
        info = embed.probe_manimgx(executable="/nonexistent/manimgx")
        self.assertFalse(info["available"])
        self.assertEqual(info["python_required"], "3.13 or 3.14")
        self.assertIn("uv tool install --python 3.14 manimgx", info["install"])
        self.assertIsNone(info["version"])

    def test_check_exits_zero_without_manimgx(self):
        def missing(*_args, **_kwargs):
            raise FileNotFoundError("manimgx")

        info = embed.probe_manimgx(executable=sys.executable, runner=missing)
        self.assertFalse(info["available"])
        buf = StringIO()
        with redirect_stdout(buf):
            code = embed.main(["--check"])
        self.assertEqual(code, 0)
        self.assertIn("available", json.loads(buf.getvalue()))

    def test_planned_commands_use_documented_flags_only(self):
        commands = embed.planned_commands(
            "scene.py",
            scene_name="Rearrange",
            video_path="clip.mp4",
            storyboard_path="clip.storyboard.png",
        )
        self.assertEqual(
            commands["inspect"],
            ["manimgx", "inspect", "-o", "clip.storyboard.png", "scene.py", "Rearrange"],
        )
        self.assertEqual(
            commands["render"],
            [
                "manimgx", "render", "-o", "clip.mp4", "-r", "1920x1080", "--fps", "30",
                "scene.py", "Rearrange",
            ],
        )
        self.assertEqual(list(commands), ["inspect", "render"])
        joined = " ".join(commands["inspect"] + commands["render"])
        for flag in ("--png", "--svg", "--format", "-ql", "--transparent"):
            self.assertNotIn(flag, joined)

    def test_commands_cli_does_not_require_the_scene_file(self):
        buf = StringIO()
        with redirect_stdout(buf):
            code = embed.main([
                "--commands", "scene.py", "--scene", "Rearrange",
                "--video", "clip.mp4", "--storyboard", "board.png",
            ])
        self.assertEqual(code, 0)
        payload = json.loads(buf.getvalue())
        self.assertEqual(payload["inspect"][0], "manimgx")
        self.assertNotIn("--png", payload["render"])

    def test_rejects_odd_resolution_and_bad_scene_name(self):
        with self.assertRaises(ValueError):
            embed.planned_commands("scene.py", resolution="1919x1080")
        with self.assertRaises(ValueError):
            embed.planned_commands("scene.py", scene_name="-o")
        with self.assertRaises(ValueError):
            embed.planned_commands("scene.py", fps=0)


class MotionEmbedTests(unittest.TestCase):
    def test_video_keeps_editable_caption_and_notes(self):
        try:
            from lxml import etree
            from pptx import Presentation
        except ImportError:
            self.skipTest("python-pptx or lxml is not installed")
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            base = root / "base.pptx"
            deck = Presentation()
            deck.slide_width = embed.SLIDE_W
            deck.slide_height = embed.SLIDE_H
            deck.slides.add_slide(deck.slide_layouts[6])
            deck.save(str(base))
            video = root / "clip.mp4"
            video.write_bytes(MP4_STUB)
            poster = root / "board.png"
            poster.write_bytes(_png())
            dest = root / "deck.pptx"
            result = embed.embed_motion(
                dest,
                video=video,
                poster=poster,
                deck=base,
                title="移项怎么变",
                caption="单击播放。可编辑式子仍走公式，不看这段影片。",
                notes="动画是示意。式子的可编辑版本在公式页。",
            )
            self.assertEqual(result["slide"], 2)
            self.assertEqual(result["poster"], "supplied")
            self.assertTrue(result["video"])
            report = audit.audit(dest)
            self.assertEqual(report["slide_count"], 2)
            slide = report["slides"][1]
            self.assertEqual(slide["picture_count"], 1)
            self.assertTrue(slide["has_nonempty_body_notes"])
            self.assertEqual(slide["large_picture_hints"], [])
            self.assertNotIn(2, report["slides_with_empty_lower_content_band"])
            self.assertNotIn(2, report["slides_over_word_cap"])
            self.assertNotIn(2, report["slides_with_font_below_min"])
            texts = [item["text"] for item in slide["native_texts"]]
            self.assertTrue(any("移项" in text for text in texts))
            self.assertTrue(any("单击播放" in text for text in texts))
            self._assert_video_package(dest, etree, slide_part="ppt/slides/slide2.xml")

    def test_missing_tool_is_not_required_and_placeholder_is_not_a_frame(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            video = root / "clip.mp4"
            video.write_bytes(MP4_STUB)
            dest = root / "deck.pptx"
            result = embed.embed_motion(dest, video=video, caption="单击播放这段示意。")
            self.assertEqual(result["poster"], "placeholder")
            with zipfile.ZipFile(dest) as package:
                images = [
                    name for name in package.namelist()
                    if name.startswith("ppt/media/image")
                ]
                self.assertTrue(images)
                self.assertTrue(package.read(images[0]).startswith(b"\x89PNG"))

    def test_storyboard_still_without_video_is_a_picture(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            poster = root / "board.png"
            poster.write_bytes(_png())
            dest = root / "still.pptx"
            result = embed.embed_motion(
                dest,
                poster=poster,
                caption="这是分镜检查图，不是可编辑公式。",
                notes="没有 MP4。公式仍用 OMML。",
            )
            self.assertEqual(result["poster"], "still")
            self.assertFalse(result["video"])
            with zipfile.ZipFile(dest) as package:
                rels = package.read("ppt/slides/_rels/slide1.xml.rels").decode("utf-8")
            self.assertNotIn(VIDEO_REL, rels)
            self.assertNotIn(MEDIA_REL, rels)

    def test_append_keeps_the_existing_slide(self):
        try:
            from pptx import Presentation
            from pptx.util import Inches
        except ImportError:
            self.skipTest("python-pptx is not installed")
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / "base.pptx"
            deck = Presentation()
            deck.slide_width = embed.SLIDE_W
            deck.slide_height = embed.SLIDE_H
            slide = deck.slides.add_slide(deck.slide_layouts[6])
            slide.shapes.add_textbox(Inches(0.7), Inches(0.4), Inches(8), Inches(0.6)).text_frame.text = "原页标题"
            deck.save(str(source))
            video = root / "clip.mp4"
            video.write_bytes(MP4_STUB)
            dest = root / "out.pptx"
            result = embed.embed_motion(dest, video=video, deck=source, notes="第二页才是动画。")
            self.assertEqual(result["slide"], 2)
            report = audit.audit(dest)
            self.assertEqual(report["slide_count"], 2)
            self.assertIn("原页标题", report["slides"][0]["native_texts"][0]["text"])
            self.assertTrue(report["slides"][1]["has_nonempty_body_notes"])
            self.assertNotEqual(source.read_bytes(), dest.read_bytes())

    def test_refuses_to_overwrite_the_input_deck(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / "base.pptx"
            source.write_bytes(b"not-used")
            with self.assertRaises(ValueError):
                embed.embed_motion(source, video=root / "clip.mp4", deck=source)

    def _assert_video_package(self, path, etree, slide_part="ppt/slides/slide1.xml"):
        import xml.etree.ElementTree as ET
        rel_part = slide_part.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels"
        with zipfile.ZipFile(path) as package:
            slide = ET.fromstring(package.read(slide_part))
            rels = ET.fromstring(package.read(rel_part))
            types = package.read("[Content_Types].xml").decode("utf-8")
            names = package.namelist()
            slide_xml = package.read(slide_part)
        self.assertIn("video/mp4", types)
        pic = slide.find(".//p:pic", NS)
        self.assertIsNotNone(pic)
        video = pic.find(".//a:videoFile", NS)
        self.assertIsNotNone(video)
        link = video.get("{%s}link" % NS["r"])
        targets = {}
        for rel in rels:
            targets[rel.get("Id")] = (rel.get("Type"), rel.get("Target"))
        self.assertEqual(targets[link][0], VIDEO_REL)
        media_rel = None
        for rel_type, target in targets.values():
            if rel_type == MEDIA_REL:
                media_rel = target
        self.assertEqual(media_rel, targets[link][1])
        media_name = "ppt/media/" + targets[link][1].split("/")[-1]
        self.assertIn(media_name, names)
        blip = pic.find(".//a:blip", NS)
        self.assertIsNotNone(blip)
        timing = slide.find("p:timing", NS)
        self.assertIsNotNone(timing)
        shape_id = pic.find(".//p:cNvPr", NS).get("id")
        target_ids = [node.get("spid") for node in timing.findall(".//p:spTgt", NS)]
        self.assertIn(shape_id, target_ids)
        if etree is None:
            return
        schema_dir = ROOT / "tests" / "xsd"
        if not (schema_dir / "pml.xsd").is_file():
            return
        schema = etree.XMLSchema(etree.parse(str(schema_dir / "pml.xsd")))
        doc = etree.fromstring(slide_xml)
        if not schema.validate(doc):
            errors = [err.message for err in schema.error_log][:6]
            self.fail("slide failed XSD: " + "; ".join(errors))


if __name__ == "__main__":
    unittest.main()
