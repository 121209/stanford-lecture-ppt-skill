"""TeX picture fallback: ladder selection does not need a TeX install."""
import hashlib
import importlib.util
import os
import re
import sys
import tempfile
import unittest
import zipfile
from io import BytesIO
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
import test_helpers


def load(name):
    path = ROOT / "scripts" / f"{name}.py"
    existing = sys.modules.get(name)
    if existing is not None and getattr(existing, "__file__", None):
        if Path(existing.__file__).resolve() == path.resolve():
            return existing
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


VENDOR_SHA256 = {
    "formula_ast.py": "67c65000fdf8682f9542fb03a39a75ccf9352a7a9efb53152dfe3383cbf17878",
    "formula_compiler.py": "4b487806c1350f9f373bef8815206ba37dc02c5db158de72370c26d1b9ad92ba",
    "formula_omml.py": "07e0db594546646b88772ab772820e959b0efdcc03b51bbaa9df0ce1b990ce6d",
    "formula_parser.py": "4636546110b62a105f11b0554058ed8ccf6154db0bc631147777f9b3b9384088",
    "formula_profile.py": "bedc1382bfc5f0c771f0137d75a3df246fd7d896f7610ff519b14c073b27bbd3",
}


def _png(width, height, rgba):
    from PIL import Image
    image = Image.new("RGBA", (width, height), rgba)
    buffer = BytesIO()
    image.save(buffer, format="PNG")
    return buffer.getvalue()


class TexFallbackTests(unittest.TestCase):
    def test_vendor_formula_files_stay_byte_identical(self):
        folder = ROOT / "scripts" / "vendor" / "pptmaster_formula"
        for name, digest in VENDOR_SHA256.items():
            blob = (folder / name).read_bytes()
            self.assertEqual(hashlib.sha256(blob).hexdigest(), digest, name)

    def test_degradation_ladder_selection(self):
        tex = load("tex_fallback")
        self.assertEqual(tex.PICTURE_RUNGS, ("tex-png", "mathtext"))
        self.assertEqual(tex.OMML_RUNGS, ("vendor", "pandoc"))
        self.assertEqual(tex.DEFAULT_ENGINE_CANDIDATES, ("xelatex", "lualatex", "pdflatex"))
        self.assertGreaterEqual(tex.MIN_DPI, 300)
        self.assertGreaterEqual(tex.DEFAULT_DPI, tex.MIN_DPI)

        def boom(*_args, **_kwargs):
            raise AssertionError("select_degradation must not probe PATH")

        original = tex.shutil.which
        tex.shutil.which = boom
        try:
            quiet = tex.select_degradation(
                compiler_ok=True,
                pandoc_present=True,
                has_latex=True,
                engine_present=False,
                png_converter_present=False,
                mathtext_present=True,
            )
            pretty = tex.select_degradation(
                compiler_ok=True,
                pandoc_present=False,
                has_latex=True,
                engine_present=True,
                png_converter_present=True,
                svg_converter_present=True,
            )
            rescue = tex.select_degradation(
                compiler_ok=False,
                pandoc_present=True,
                has_latex=True,
                engine_present=False,
                png_converter_present=False,
            )
            blocked = tex.select_degradation(
                compiler_ok=False,
                pandoc_present=False,
                has_latex=True,
                engine_present=False,
                png_converter_present=False,
                mathtext_present=False,
            )
            svg_only = tex.select_degradation(
                compiler_ok=True,
                has_latex=True,
                engine_present=True,
                png_converter_present=False,
                svg_converter_present=True,
                mathtext_present=True,
            )
            runs_only = tex.select_degradation(
                compiler_ok=True,
                has_latex=False,
                engine_present=True,
                png_converter_present=True,
            )
        finally:
            tex.shutil.which = original

        self.assertEqual(quiet["omml_attempt"], "vendor")
        self.assertEqual(quiet["omml"], "vendor")
        self.assertEqual(quiet["omml_if_compiler_raises"], "pandoc")
        self.assertEqual(quiet["picture"], "mathtext")
        self.assertIsNone(quiet["svg_sidecar"])
        self.assertNotEqual(quiet["picture"], "pandoc")

        self.assertEqual(pretty["omml"], "vendor")
        self.assertEqual(pretty["picture"], "tex-png")
        self.assertEqual(pretty["svg_sidecar"], "tex-svg")
        self.assertEqual(pretty["omml_if_compiler_raises"], "fail")

        self.assertEqual(rescue["omml_attempt"], "vendor")
        self.assertEqual(rescue["omml"], "pandoc")
        self.assertEqual(rescue["picture"], "mathtext")

        self.assertEqual(blocked["omml"], "fail")
        self.assertEqual(blocked["picture"], "fail")

        self.assertEqual(svg_only["picture"], "mathtext")
        self.assertEqual(svg_only["svg_sidecar"], "tex-svg")
        self.assertEqual(runs_only["picture"], "mathtext")
        self.assertIsNone(runs_only["svg_sidecar"])

    def test_engine_probe_order_and_missing_toolchain(self):
        tex = load("tex_fallback")
        seen = []

        def which(cmd):
            seen.append(cmd)
            if cmd == "pdflatex":
                return "/usr/bin/pdflatex"
            return None

        original = tex.shutil.which
        tex.shutil.which = which
        try:
            found = tex.find_engine()
            status = tex.toolchain_status()
        finally:
            tex.shutil.which = original
        self.assertEqual(seen[:3], ["xelatex", "lualatex", "pdflatex"])
        self.assertEqual(found, ("pdflatex", "/usr/bin/pdflatex"))
        self.assertEqual(status["engine"], "pdflatex")
        self.assertFalse(status["tex_png"])
        self.assertIsNone(status["png_converter"])

        tex.shutil.which = lambda cmd: None
        try:
            self.assertIsNone(tex.find_engine())
            empty = tex.toolchain_status()
            self.assertFalse(empty["tex_png"])
            self.assertFalse(empty["tex_svg"])
            with self.assertRaises(tex.TexFallbackError) as caught:
                tex.require_engine()
            self.assertIn("xelatex", str(caught.exception))
        finally:
            tex.shutil.which = original

    def test_document_wrapper_matches_nature_and_keeps_shell_escape_off(self):
        tex = load("tex_fallback")
        document = tex.build_latex_document(r"\frac{a}{b}", display=True, font_size=28)
        self.assertIn(r"\documentclass[border=2pt]{standalone}", document)
        self.assertNotIn("varwidth", document)
        self.assertIn(r"\usepackage{lmodern}", document)
        self.assertIn(r"\usepackage{amsmath,amssymb,mathtools,bm}", document)
        self.assertIn(r"\usepackage{xcolor}", document)
        self.assertIn(r"\color[HTML]{222222}", document)
        self.assertIn(r"\fontsize{28.00}", document)
        self.assertIn("$\\displaystyle\n\\frac{a}{b}\n$", document)
        self.assertNotIn("\\[", document)
        self.assertNotIn("editable", document)
        inline = tex.build_latex_document(r"$d_k$", display=False, font_size=22)
        self.assertIn("d_k", inline)
        self.assertIn(tex.INLINE_STRUT, inline)
        self.assertNotIn("\\[", inline)
        self.assertTrue(inline.strip().endswith("\\end{document}"))
        command = tex.latex_command("xelatex", "formula.tex")
        self.assertEqual(
            command,
            [
                "xelatex",
                "-interaction=nonstopmode",
                "-halt-on-error",
                "-file-line-error",
                "formula.tex",
            ],
        )
        self.assertNotIn("-shell-escape", command)
        magick = tex.imagemagick_png_command("magick", "formula.pdf", "formula.png", 300)
        self.assertEqual(magick[0:6], ["magick", "-density", "300", "formula.pdf", "-trim", "+repage"])
        self.assertEqual(magick[-1], "PNG32:formula.png")
        svg = tex.dvisvgm_svg_command("dvisvgm", "formula.pdf", "formula.svg")
        self.assertEqual(
            svg,
            ["dvisvgm", "--pdf", "--no-fonts", "--exact", "--output", "formula.svg", "formula.pdf"],
        )
        with self.assertRaises(tex.TexFallbackError):
            tex._assert_dpi(299)
        self.assertEqual(tex._assert_dpi(300), 300)

    def test_white_paper_becomes_transparent_rgba(self):
        from PIL import Image
        tex = load("tex_fallback")
        image = Image.new("RGB", (40, 20), (255, 255, 255))
        for x in range(12, 28):
            for y in range(6, 14):
                image.putpixel((x, y), (10, 10, 10))
        buffer = BytesIO()
        image.save(buffer, format="PNG")
        blob = tex._transparent_png(buffer.getvalue(), pad=2, dpi=300)
        self.assertEqual(blob[25], 6)
        painted = Image.open(BytesIO(blob)).convert("RGBA")
        pixels = painted.load()
        width, height = painted.size
        self.assertEqual(pixels[0, 0][3], 0)
        self.assertEqual(pixels[width - 1, height - 1][3], 0)
        dark = 0
        near_white = 0
        for y in range(height):
            for x in range(width):
                red, green, blue, alpha = pixels[x, y]
                if alpha < 16:
                    continue
                if red >= 248 and green >= 248 and blue >= 248:
                    near_white += 1
                elif red < 80 and green < 80 and blue < 80:
                    dark += 1
        self.assertEqual(near_white, 0)
        self.assertGreater(dark, 10)

        translucent = Image.new("RGBA", (8, 8), (0, 0, 0, 0))
        translucent.putpixel((3, 3), (0, 0, 0, 80))
        raw = BytesIO()
        translucent.save(raw, format="PNG")
        kept = Image.open(BytesIO(tex._transparent_png(raw.getvalue(), pad=0))).convert("RGBA")
        flat = kept.get_flattened_data() if hasattr(kept, "get_flattened_data") else kept.getdata()
        alphas = [pixel[3] for pixel in flat if pixel[3]]
        self.assertEqual(alphas, [80])
        clear = Image.new("RGBA", (10, 10), (0, 0, 0, 0))
        clear.putpixel((4, 4), (34, 34, 34, 255))
        raw_clear = BytesIO()
        clear.save(raw_clear, format="PNG")
        preserved = Image.open(
            BytesIO(tex._transparent_png(raw_clear.getvalue(), pad=0, trim=False))
        )
        self.assertEqual(preserved.getpixel((4, 4)), (34, 34, 34, 255))
        self.assertEqual(preserved.getpixel((0, 0))[3], 0)

    def test_inline_natural_size_widens_a_narrow_softmax_box(self):
        inject = load("inject_omml")
        # 0.70 in by 0.42 in, the narrow slot from the visual deck.
        xfrm = (
            '<a:xfrm><a:off x="7589520" y="3657600"/>'
            '<a:ext cx="640080" cy="384048"/></a:xfrm>'
        )
        # 22pt softmax is wider than 0.70 in. 900 px at 400 dpi is 2.25 in.
        placed = inject._place_inline_natural(xfrm, 900, 180, 400)
        ext = dict(re.findall(r'([A-Za-z0-9:]+)="([^"]*)"', placed))
        self.assertGreater(int(ext["cx"]), 640080)
        self.assertGreater(int(ext["cy"]), 0)
        off_y = int(ext["y"])
        self.assertLess(off_y, 3657600)
        self.assertGreater(off_y, 0)
        fitted = inject._fit_xfrm(xfrm, 900, 180)
        fitted_ext = dict(re.findall(r'([A-Za-z0-9:]+)="([^"]*)"', fitted))
        self.assertLessEqual(int(fitted_ext["cx"]), 640080)

    def test_stack_pngs_leaves_a_single_blob_unchanged(self):
        tex = load("tex_fallback")
        one = _png(8, 6, (34, 34, 34, 255))
        self.assertEqual(tex.stack_pngs([one]), one)
        other = _png(8, 4, (34, 34, 34, 255))
        from PIL import Image
        stacked = Image.open(BytesIO(tex.stack_pngs([one, other])))
        self.assertGreater(stacked.height, 6 + 4)
        self.assertEqual(stacked.getpixel((0, 0))[3], 255)

    def test_live_tex_render_is_optional(self):
        tex = load("tex_fallback")
        status = tex.toolchain_status()
        if not status["tex_png"]:
            with self.assertRaises(tex.TexFallbackError):
                tex.render_tex_png(r"\lambda", display=False, dpi=300)
        else:
            blob = tex.render_tex_png(
                r"\lambda", display=False, dpi=max(300, tex.MIN_DPI), font_size=22, trim=False,
            )
            self.assertEqual(blob[:8], b"\x89PNG\r\n\x1a\n")
            self.assertEqual(blob[25], 6)
            from PIL import Image
            image = Image.open(BytesIO(blob)).convert("RGBA")
            self.assertEqual(image.getpixel((0, 0))[3], 0)
            rows = image.get_flattened_data() if hasattr(image, "get_flattened_data") else image.getdata()
            solid = [
                pixel for pixel in rows
                if pixel[3] == 255 and pixel[0] < 80 and pixel[1] < 80 and pixel[2] < 80
            ]
            self.assertTrue(solid, "TeX ink should stay opaque, not washed out")
            display = tex.render_tex_png(
                r"\sum_{i=1}^{n} i", display=True, dpi=300, font_size=22, trim=True,
            )
            self.assertEqual(display[25], 6)
            self.assertGreater(len(display), 500)
        if not status["tex_svg"]:
            with tempfile.TemporaryDirectory() as folder:
                with self.assertRaises(tex.TexFallbackError) as caught:
                    tex.render_tex_svg("a", Path(folder) / "a.svg")
            message = str(caught.exception)
            self.assertTrue(
                "dvisvgm" in message or "xelatex" in message or "LaTeX engine" in message,
                message,
            )

    def _patched_tex(self, status, render):
        inject = load("inject_omml")
        tex = inject._load_tex_fallback()
        originals = (tex.toolchain_status, tex.render_tex_png, inject.render_equation_png, inject._note_tex_degrade)
        tex.toolchain_status = lambda: status
        tex.render_tex_png = render
        return inject, tex, originals

    def _restore(self, tex, inject, originals):
        tex.toolchain_status, tex.render_tex_png, inject.render_equation_png, inject._note_tex_degrade = originals

    def test_tex_png_is_the_fallback_and_omml_stays_editable(self):
        known = _png(80, 20, (0, 0, 0, 0))
        from PIL import Image
        image = Image.open(BytesIO(known))
        draw = image.load()
        for x in range(10, 70):
            for y in range(4, 16):
                draw[x, y] = (34, 34, 34, 255)
        buffer = BytesIO()
        image.save(buffer, format="PNG")
        known = buffer.getvalue()
        self.assertEqual(known[25], 6)
        calls = []

        def fake_render(source, **kwargs):
            calls.append((source, kwargs.get("display"), kwargs.get("dpi")))
            return known

        def fail_math(*_args, **_kwargs):
            raise AssertionError("mathtext should not run when TeX PNG succeeds")

        status = {
            "engine": "xelatex",
            "engine_path": "/usr/bin/xelatex",
            "png_converter": "pdftocairo",
            "svg_converter": "dvisvgm",
            "tex_png": True,
            "tex_svg": True,
        }
        inject, tex, originals = self._patched_tex(status, fake_render)
        inject.render_equation_png = fail_math
        try:
            with tempfile.TemporaryDirectory() as folder:
                folder = Path(folder)
                src = folder / "in.pptx"
                test_helpers.write_deck(src, [[{"text": "{{MATH:attn}}", "sz": 2200, "y": 1800000}]])
                dest = folder / "out.pptx"
                inject.inject_pptx(src, {"attn": {"latex": r"\frac{QK^{\top}}{\sqrt{d_k}}"}}, dest)
                with zipfile.ZipFile(dest) as package:
                    slide = package.read("ppt/slides/slide1.xml")
                    media = [name for name in package.namelist() if name.endswith(".png")]
                    self.assertEqual(package.read(media[0]), known)
                    self.assertTrue(all(not name.endswith(".tex") for name in package.namelist()))
                report = test_helpers.audit.audit(dest)
        finally:
            self._restore(tex, inject, originals)
        self.assertEqual(calls, [(r"\frac{QK^{\top}}{\sqrt{d_k}}", True, 600)])
        self.assertGreaterEqual(calls[0][2], 300)
        self.assertIn(b"<m:f>", slide)
        self.assertIn(b"mc:AlternateContent", slide)
        self.assertIn(b"mc:Fallback", slide)
        self.assertNotIn(b"dfrac", slide)
        self.assertNotIn(b"\\frac", slide)
        self.assertNotIn(b"editable", slide)
        self.assertEqual(report["office_math_count"], 1)
        self.assertEqual(report["slides_with_invalid_office_math"], [])
        self.assertEqual(report["slides_with_raw_latex"], [])

    def test_inline_tex_fallback_is_not_display(self):
        known = _png(30, 16, (34, 34, 34, 255))
        calls = []

        def fake_render(source, **kwargs):
            calls.append(kwargs.get("display"))
            return known

        status = {
            "engine": "xelatex",
            "engine_path": "/usr/bin/xelatex",
            "png_converter": "pdftocairo",
            "svg_converter": None,
            "tex_png": True,
            "tex_svg": False,
        }
        inject, tex, originals = self._patched_tex(status, fake_render)
        inject.render_equation_png = lambda *_a, **_k: (_ for _ in ()).throw(
            AssertionError("mathtext")
        )
        try:
            png = inject._fallback_png(
                {"dk": {"latex": "d_k", "inline": True}},
                ["dk"],
                ["$d_k$"],
                2.0,
                0.4,
                "dk",
                True,
                True,
            )
        finally:
            self._restore(tex, inject, originals)
        self.assertEqual(png, known)
        self.assertEqual(calls, [False])

    def test_tex_failure_degrades_to_mathtext_and_keeps_omml(self):
        from PIL import Image
        notes = []

        def boom(_source, **_kwargs):
            raise load("tex_fallback").TexFallbackError("engine missing")

        status = {
            "engine": "xelatex",
            "engine_path": "/usr/bin/xelatex",
            "png_converter": "pdftocairo",
            "svg_converter": None,
            "tex_png": True,
            "tex_svg": False,
        }
        inject, tex, originals = self._patched_tex(status, boom)
        inject._note_tex_degrade = lambda ids, exc: notes.append((list(ids), str(exc)))
        try:
            with tempfile.TemporaryDirectory() as folder:
                folder = Path(folder)
                src = folder / "in.pptx"
                test_helpers.write_deck(src, [[{"text": "{{MATH:rate}}", "sz": 2000}]])
                dest = folder / "out.pptx"
                inject.inject_pptx(src, {"rate": {"latex": r"\lambda(t)=\alpha"}}, dest)
                with zipfile.ZipFile(dest) as package:
                    slide = package.read("ppt/slides/slide1.xml")
                    media = [name for name in package.namelist() if name.endswith(".png")]
                    blob = package.read(media[0])
                report = test_helpers.audit.audit(dest)
        finally:
            self._restore(tex, inject, originals)
        self.assertEqual(blob[25], 6)
        image = Image.open(BytesIO(blob)).convert("RGBA")
        self.assertEqual(image.getpixel((0, 0))[3], 0)
        self.assertIn(b"mc:AlternateContent", slide)
        self.assertNotIn(b"\\lambda", slide)
        self.assertEqual(report["office_math_count"], 1)
        self.assertEqual(report["slides_with_raw_latex"], [])
        self.assertEqual(notes[0][0], ["rate"])
        self.assertIn("engine missing", notes[0][1])

    def test_runs_only_stays_on_mathtext_even_if_tex_exists(self):
        def boom(*_args, **_kwargs):
            raise AssertionError("runs-only specs have no TeX source")

        status = {
            "engine": "xelatex",
            "engine_path": "/usr/bin/xelatex",
            "png_converter": "magick",
            "svg_converter": None,
            "tex_png": True,
            "tex_svg": False,
        }
        inject, tex, originals = self._patched_tex(status, boom)
        try:
            with tempfile.TemporaryDirectory() as folder:
                folder = Path(folder)
                src = folder / "in.pptx"
                test_helpers.write_deck(src, [[{"text": "{{MATH:rate}}", "sz": 2000}]])
                dest = folder / "out.pptx"
                inject.inject_pptx(src, {"rate": {"runs": [{"t": "λ"}]}}, dest)
                with zipfile.ZipFile(dest) as package:
                    blob = package.read(
                        next(name for name in package.namelist() if name.endswith(".png"))
                    )
        finally:
            self._restore(tex, inject, originals)
        self.assertEqual(blob[25], 6)

    def test_absent_tex_does_not_warn_and_sidecar_stays_out_of_pptx(self):
        notes = []
        status = {
            "engine": None,
            "engine_path": None,
            "png_converter": None,
            "svg_converter": None,
            "tex_png": False,
            "tex_svg": False,
        }
        inject, tex, originals = self._patched_tex(status, lambda *_a, **_k: b"")
        inject._note_tex_degrade = lambda *_args: notes.append("warned")
        previous = os.environ.get("STANFORD_FORMULA_TEX_DIR")
        try:
            with tempfile.TemporaryDirectory() as folder:
                folder = Path(folder)
                side = folder / "tex"
                os.environ["STANFORD_FORMULA_TEX_DIR"] = str(side)
                src = folder / "in.pptx"
                test_helpers.write_deck(src, [[{"text": "{{MATH:rate}}", "sz": 2000}]])
                dest = folder / "out.pptx"
                inject.inject_pptx(src, {"rate": {"latex": r"\lambda(t)=\alpha"}}, dest)
                document = (side / "rate.tex").read_text(encoding="utf-8")
                blob = dest.read_bytes()
                with zipfile.ZipFile(dest) as package:
                    names = package.namelist()
        finally:
            if previous is None:
                os.environ.pop("STANFORD_FORMULA_TEX_DIR", None)
            else:
                os.environ["STANFORD_FORMULA_TEX_DIR"] = previous
            self._restore(tex, inject, originals)
        self.assertEqual(notes, [])
        self.assertIn("standalone", document)
        self.assertIn(r"\lambda", document)
        self.assertNotIn("editable: false", document)
        self.assertNotIn(b"\\lambda", blob)
        self.assertNotIn(b"\\documentclass", blob)
        self.assertTrue(all(not name.endswith(".tex") for name in names))

    def test_two_latex_sources_stack_before_mathtext(self):
        from PIL import Image
        first = _png(12, 10, (34, 34, 34, 255))
        second = _png(12, 8, (34, 34, 34, 255))
        calls = []

        def fake_render(source, **_kwargs):
            calls.append(source)
            return first if len(calls) == 1 else second

        status = {
            "engine": "lualatex",
            "engine_path": "/usr/bin/lualatex",
            "png_converter": "ghostscript",
            "svg_converter": None,
            "tex_png": True,
            "tex_svg": False,
        }
        inject, tex, originals = self._patched_tex(status, fake_render)
        try:
            blob = inject._fallback_png(
                {"a": {"latex": r"\alpha"}, "b": {"latex": r"\beta"}},
                ["a", "b"],
                ["$a$", "$b$"],
                4.0,
                1.0,
                "ab",
                True,
                False,
            )
        finally:
            self._restore(tex, inject, originals)
        self.assertEqual(calls, [r"\alpha", r"\beta"])
        stacked = Image.open(BytesIO(blob))
        self.assertGreater(stacked.height, 10 + 8)
        self.assertEqual(blob[25], 6)


if __name__ == "__main__":
    unittest.main()
