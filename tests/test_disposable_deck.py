"""Working-copy saves, doctor profile, and sample writers stay off the demo deck."""
import hashlib
import importlib.util
import json
import os
import tempfile
import unittest
from contextlib import redirect_stderr, redirect_stdout
from io import StringIO
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SENTINEL_SECRET = "sk-sentinel-do-not-leak"


def load(name):
    spec = importlib.util.spec_from_file_location(name, ROOT / "scripts" / f"{name}.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


safe = load("safe_edit")
doctor = load("doctor")


def _walk_strings(value):
    if isinstance(value, str):
        yield value
    elif isinstance(value, dict):
        for key, item in value.items():
            yield key
            yield from _walk_strings(item)
    elif isinstance(value, (list, tuple)):
        for item in value:
            yield from _walk_strings(item)


class SafeEditTests(unittest.TestCase):
    def test_default_saves_a_copy_and_records_the_source_hash(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / "deck.pptx"
            payload = b"PK disposable deck bytes"
            source.write_bytes(payload)
            record = root / "edit.json"
            info = safe.prepare_working_copy(source, record=record)
            self.assertFalse(info["in_place"])
            self.assertTrue(info["copied"])
            self.assertEqual(info["sha256"], hashlib.sha256(payload).hexdigest())
            self.assertEqual(source.read_bytes(), payload)
            copy = Path(info["working_copy"])
            self.assertNotEqual(copy.resolve(), source.resolve())
            self.assertEqual(copy.read_bytes(), payload)
            saved = json.loads(record.read_text(encoding="utf-8"))
            self.assertEqual(saved["sha256"], info["sha256"])
            self.assertNotIn("api_key", saved)
            self.assertNotIn(SENTINEL_SECRET, json.dumps(saved))

    def test_cli_without_in_place_does_not_overwrite(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / "deck.pptx"
            source.write_bytes(b"original")
            buf = StringIO()
            with redirect_stdout(buf):
                code = safe.main([str(source)])
            self.assertEqual(code, 0)
            info = json.loads(buf.getvalue())
            self.assertFalse(info["in_place"])
            self.assertEqual(source.read_bytes(), b"original")
            self.assertTrue(Path(info["working_copy"]).is_file())
            self.assertEqual(Path(info["working_copy"]).read_bytes(), b"original")

    def test_in_place_is_only_the_boolean_true(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / "deck.pptx"
            source.write_bytes(b"original")
            for value in ("yes", "true", 1, "in_place"):
                with self.assertRaises(ValueError):
                    safe.prepare_working_copy(source, in_place=value)
                self.assertEqual(source.read_bytes(), b"original")
                self.assertFalse((root / "deck.work.pptx").exists())
            info = safe.prepare_working_copy(source, in_place=True)
            self.assertTrue(info["in_place"])
            self.assertFalse(info["copied"])
            self.assertEqual(Path(info["working_copy"]).resolve(), source.resolve())
            self.assertEqual(source.read_bytes(), b"original")
            self.assertFalse((root / "deck.work.pptx").exists())

    def test_dest_equal_to_source_is_refused(self):
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder) / "deck.pptx"
            source.write_bytes(b"original")
            with self.assertRaises(ValueError):
                safe.prepare_working_copy(source, dest=source)
            self.assertEqual(source.read_bytes(), b"original")

    def test_existing_dest_and_taken_sibling_are_not_clobbered(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / "deck.pptx"
            source.write_bytes(b"original")
            taken = root / "deck.work.pptx"
            taken.write_bytes(b"keep-me")
            explicit = root / "out.pptx"
            explicit.write_bytes(b"also-keep")
            with self.assertRaises(ValueError):
                safe.prepare_working_copy(source, dest=explicit)
            self.assertEqual(explicit.read_bytes(), b"also-keep")
            info = safe.prepare_working_copy(source)
            self.assertEqual(taken.read_bytes(), b"keep-me")
            self.assertTrue(Path(info["working_copy"]).name.startswith("deck.work-"))
            self.assertEqual(Path(info["working_copy"]).read_bytes(), b"original")


class DoctorTests(unittest.TestCase):
    def test_profile_has_no_secrets_or_author_paths(self):
        previous_key = os.environ.get("OPENAI_API_KEY")
        previous_home = os.environ.get("HOME")
        os.environ["OPENAI_API_KEY"] = SENTINEL_SECRET
        os.environ["HOME"] = "/home/author-should-not-be-required"
        try:
            profile = doctor.collect_profile()
        finally:
            if previous_key is None:
                os.environ.pop("OPENAI_API_KEY", None)
            else:
                os.environ["OPENAI_API_KEY"] = previous_key
            if previous_home is None:
                os.environ.pop("HOME", None)
            else:
                os.environ["HOME"] = previous_home
        blob = json.dumps(profile)
        self.assertNotIn(SENTINEL_SECRET, blob)
        self.assertNotIn("OPENAI", blob)
        self.assertNotIn("/home/", blob)
        self.assertNotIn("\\", blob)
        self.assertTrue(profile["secret_free"])
        self.assertEqual(profile["schema"], "stanford-lecture-ppt-runtime-profile")
        for text in _walk_strings(profile):
            self.assertNotIn(SENTINEL_SECRET, text)
            self.assertFalse(text.startswith("/"))
            self.assertNotIn("api_key", text)
        self.assertNotIn("path", profile)
        self.assertNotIn("executable", profile["optional"]["manimgx"])
        self.assertIn("formula_picture", profile["degrade"])
        self.assertIn(profile["degrade"]["formula_picture"], ("tex-png", "mathtext"))
        self.assertIn(profile["degrade"]["render_checks"], ("libreoffice", "skip"))
        self.assertIn(profile["degrade"]["motion"], ("manimgx", "editable-only"))

    def test_optional_gaps_still_exit_zero(self):
        originals = {
            "python_status": doctor.python_status,
            "pptx_status": doctor.pptx_status,
            "tex_status": doctor.tex_status,
            "pdftocairo_status": doctor.pdftocairo_status,
            "libreoffice_status": doctor.libreoffice_status,
            "manimgx_status": doctor.manimgx_status,
        }
        doctor.python_status = lambda: {"version": "3.12.0", "minimum": "3.10", "ok": True}
        doctor.pptx_status = lambda: {"available": True, "version": "1.0.2"}
        doctor.tex_status = lambda: {
            "available": False, "engine": None, "png_converter": None,
            "tex_png": False, "disabled_by_env": False,
        }
        doctor.pdftocairo_status = lambda: {"available": False}
        doctor.libreoffice_status = lambda: {"available": False, "command": None}
        doctor.manimgx_status = lambda: {
            "available": False, "version": None, "python_required": "3.13 or 3.14",
        }
        try:
            with tempfile.TemporaryDirectory() as folder:
                target = Path(folder) / "runtime-profile.json"
                buf = StringIO()
                with redirect_stdout(buf):
                    code = doctor.main(["--json", "--write", str(target)])
                profile = json.loads(buf.getvalue())
                written = json.loads(target.read_text(encoding="utf-8"))
            self.assertEqual(code, 0)
            self.assertTrue(profile["ok"])
            self.assertEqual(written, profile)
            self.assertEqual(
                profile["missing_optional"],
                ["tex", "pdftocairo", "libreoffice", "manimgx"],
            )
            self.assertEqual(profile["degrade"]["formula_picture"], "mathtext")
            self.assertEqual(profile["degrade"]["render_checks"], "skip")
            self.assertEqual(profile["degrade"]["motion"], "editable-only")
            text = doctor.render_text(profile)
            self.assertIn("status: ok", text)
            self.assertIn("missing optional: tex, pdftocairo, libreoffice, manimgx", text)
        finally:
            for name, original in originals.items():
                setattr(doctor, name, original)

    def test_missing_python_or_pptx_exits_nonzero(self):
        originals = (doctor.python_status, doctor.pptx_status)
        doctor.python_status = lambda: {"version": "3.9.0", "minimum": "3.10", "ok": False}
        doctor.pptx_status = lambda: {"available": False, "version": None}
        try:
            buf = StringIO()
            with redirect_stdout(buf):
                code = doctor.main(["--json"])
            profile = json.loads(buf.getvalue())
            self.assertEqual(code, 1)
            self.assertFalse(profile["ok"])
            self.assertIn("status: required-missing", doctor.render_text(profile))
        finally:
            doctor.python_status, doctor.pptx_status = originals

    def test_profile_refuses_to_replace_a_pptx(self):
        with tempfile.TemporaryDirectory() as folder:
            target = Path(folder) / "deck.pptx"
            target.write_bytes(b"keep")
            with redirect_stderr(StringIO()):
                code = doctor.main(["--write", str(target)])
            self.assertEqual(code, 2)
            self.assertEqual(target.read_bytes(), b"keep")

    def test_real_doctor_exits_zero_when_pptx_imports(self):
        try:
            import pptx  # noqa: F401
        except ImportError:
            self.skipTest("python-pptx is not installed")
        buf = StringIO()
        with redirect_stdout(buf):
            code = doctor.main([])
        self.assertEqual(code, 0)
        self.assertIn("status: ok", buf.getvalue())
        self.assertIn("python-pptx", buf.getvalue())


class SampleWriterTests(unittest.TestCase):
    def test_sample_generation_does_not_clobber_checked_in_decks(self):
        try:
            import pptx  # noqa: F401
        except ImportError:
            self.skipTest("python-pptx is not installed")
        diagrams = load("diagram_helpers")
        charts = load("chart_helpers")
        watched = [
            ROOT / "assets" / "demo.pptx",
            ROOT / "assets" / "diagram-samples" / "figure-first.pptx",
        ]
        before = {path: hashlib.sha256(path.read_bytes()).hexdigest() for path in watched}
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            diagrams.build_slide_from_spec({
                "title": "训练流程：四步示意",
                "layout": "row",
                "nodes": diagrams.FLOW_NODES,
                "edges": diagrams.FLOW_EDGES,
                "conclusion": "顺序是汇报归纳，不是原文算法。",
                "source": "来源：匿名示意，不是论文实验。",
                "page": "02",
            }, root / "one.pptx")
            spec = root / "chart.json"
            spec.write_text(json.dumps({
                "categories": ["甲", "乙"],
                "values": [1, 2],
                "region": [700000, 1600000, 6000000, 2800000],
                "highlight": 0,
                "source": "来源：测试，不是论文。",
            }), encoding="utf-8")
            self.assertEqual(charts.main([
                "--spec", str(spec), "--output", str(root / "bars.pptx"),
            ]), 0)
            self.assertTrue((root / "one.pptx").is_file())
            self.assertTrue((root / "bars.pptx").is_file())
            try:
                import matplotlib  # noqa: F401
            except ImportError:
                pass
            else:
                diagrams.write_sample_deck(root / "samples.pptx")
                self.assertTrue((root / "samples.pptx").is_file())
        for path, digest in before.items():
            self.assertEqual(hashlib.sha256(path.read_bytes()).hexdigest(), digest, path.name)


if __name__ == "__main__":
    unittest.main()
