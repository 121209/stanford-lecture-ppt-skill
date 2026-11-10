#!/usr/bin/env python3
"""Embed an opt-in ManimGX clip in an editable PPTX.

ManimGX writes an MP4 (``manimgx render``) and a storyboard PNG
(``manimgx inspect``). It does not write PPTX, and this script does not
invent a still-image export flag. The clip is optional: formulas stay on
the OMML path, and algorithm pages stay native shapes.

``--check`` only looks for ``manimgx`` on PATH. It does not install it and
it exits 0 when the tool is missing. ``--commands`` prints the inspect-then-
render argv and does not run them. Embedding needs python-pptx and an MP4
the caller has already rendered.
"""
import argparse
import json
import shutil
import struct
import subprocess
import sys
import zlib
from io import BytesIO
from pathlib import Path

SLIDE_W = 12192000
SLIDE_H = 6858000
EMU_IN = 914400
FONT = "Microsoft YaHei"
INK = (0x22, 0x22, 0x22)
ACCENT = (0xC0, 0x00, 0x00)
# 16:9 clip on the standard 13.333 x 7.5 inch slide. About 39% of the canvas,
# crossing the lower content band, under the 75% large-picture hint.
MEDIA_BOX = (2286000, 1097280, 7620000, 4286250)
TITLE_BOX = (640080, 347472, 10911840, 566928)
CAPTION_BOX = (640080, 5486400, 10911840, 365760)

PYTHON_REQUIRED = "3.13 or 3.14"
INSTALL = "uv tool install --python 3.14 manimgx"
# Skill recommendation. Omitting these flags leaves ManimGX at 1920x1080, 60 fps.
RENDER_RESOLUTION = "1920x1080"
RENDER_FPS = 30


def probe_manimgx(executable=None, runner=None):
    """Report whether the manimgx CLI can be run. Never installs it.

    ``executable`` skips PATH lookup. A missing tool is ``available: false``,
    not an exception. ``runner`` replaces subprocess for tests.
    """
    info = {
        "available": False,
        "executable": None,
        "version": None,
        "returncode": None,
        "python_required": PYTHON_REQUIRED,
        "install": INSTALL,
        "reason": None,
    }
    if executable is None:
        executable = shutil.which("manimgx")
    if not executable:
        info["reason"] = "manimgx is not on PATH"
        return info
    path = Path(executable)
    info["executable"] = str(path)
    if not path.is_file():
        info["reason"] = "manimgx executable was not found"
        return info
    run = runner or subprocess.run
    try:
        completed = run(
            [str(path), "--version"],
            capture_output=True,
            text=True,
            timeout=15,
        )
    except (OSError, subprocess.TimeoutExpired) as exc:
        info["reason"] = str(exc)
        return info
    info["returncode"] = completed.returncode
    version = (completed.stdout or "").strip() or (completed.stderr or "").strip()
    if completed.returncode != 0:
        info["reason"] = version or "manimgx --version failed"
        return info
    info["available"] = True
    info["version"] = version or None
    return info


def planned_commands(
    scene_file,
    scene_name=None,
    video_path=None,
    storyboard_path=None,
    resolution=RENDER_RESOLUTION,
    fps=RENDER_FPS,
):
    """Argv for inspect, then render. Does not execute either command.

    Flags are the ones documented by ``manimgx render`` / ``manimgx inspect``:
    ``-o``, ``-r``, ``--fps``. Inspect does not take ``-r`` or ``--fps``.
    """
    scene = str(scene_file)
    if not scene or scene.startswith("-"):
        raise ValueError("scene file must be a path, not a flag")
    if scene_name is not None and not str(scene_name).isidentifier():
        raise ValueError("scene name must be a Python identifier")
    _check_resolution(resolution)
    _check_fps(fps)
    inspect = ["manimgx", "inspect"]
    if storyboard_path:
        inspect.extend(["-o", str(storyboard_path)])
    inspect.append(scene)
    render = ["manimgx", "render"]
    if video_path:
        render.extend(["-o", str(video_path)])
    if resolution:
        render.extend(["-r", resolution])
    if fps is not None:
        render.extend(["--fps", str(int(fps))])
    render.append(scene)
    if scene_name:
        inspect.append(str(scene_name))
        render.append(str(scene_name))
    return {"inspect": inspect, "render": render}


def _check_resolution(resolution):
    if resolution is None:
        return
    text = str(resolution)
    parts = text.split("x")
    if len(parts) != 2 or not all(part.isdigit() and int(part) > 0 for part in parts):
        raise ValueError("resolution must look like 1920x1080")
    width, height = (int(part) for part in parts)
    if width % 2 or height % 2:
        raise ValueError("resolution width and height must be even")


def _check_fps(fps):
    if fps is None:
        return
    if isinstance(fps, bool) or not isinstance(fps, int) or fps < 1:
        raise ValueError("fps must be an integer >= 1")


def _pptx():
    try:
        import pptx
        from pptx.dml.color import RGBColor
        from pptx.enum.text import PP_ALIGN
        from pptx.oxml.ns import qn
        from pptx.util import Emu, Pt
    except ImportError as exc:
        raise RuntimeError("python-pptx is required to embed a motion clip") from exc
    return {
        "Presentation": pptx.Presentation,
        "RGBColor": RGBColor,
        "PP_ALIGN": PP_ALIGN,
        "qn": qn,
        "Emu": Emu,
        "Pt": Pt,
    }


def _blank_layout(prs):
    for layout in prs.slide_layouts:
        if (layout.name or "").lower() == "blank":
            return layout
    return prs.slide_layouts[len(prs.slide_layouts) - 1]


def _white_png(width=16, height=9):
    """Opaque white PNG. A placeholder poster, not a ManimGX frame."""
    def chunk(tag, data):
        crc = zlib.crc32(tag + data) & 0xFFFFFFFF
        return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", crc)

    raw = b"".join(b"\x00" + (b"\xff" * (width * 3)) for _ in range(height))
    ihdr = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", ihdr)
        + chunk(b"IDAT", zlib.compress(raw, 9))
        + chunk(b"IEND", b"")
    )


def _scale(box, slide_w, slide_h):
    if slide_w == SLIDE_W and slide_h == SLIDE_H:
        return box
    sx = slide_w / SLIDE_W
    sy = slide_h / SLIDE_H
    x, y, w, h = box
    return (int(x * sx), int(y * sy), max(1, int(w * sx)), max(1, int(h * sy)))


def _add_text(slide, box, text, size, color, api, bold=False):
    qn = api["qn"]
    x, y, w, h = box
    shape = slide.shapes.add_textbox(api["Emu"](x), api["Emu"](y), api["Emu"](w), api["Emu"](h))
    style = shape._element.find(qn("p:style"))
    if style is not None:
        shape._element.remove(style)
    frame = shape.text_frame
    frame.word_wrap = True
    body = frame._txBody.find(qn("a:bodyPr"))
    body.set("anchor", "ctr")
    paragraph = frame.paragraphs[0]
    paragraph.alignment = api["PP_ALIGN"].LEFT
    run = paragraph.add_run()
    run.text = text
    run.font.size = api["Pt"](size)
    run.font.bold = bold
    run.font.color.rgb = api["RGBColor"](*color)
    run.font.name = FONT
    r_pr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        element = r_pr.find(qn(tag))
        if element is None:
            element = r_pr.makeelement(qn(tag), {})
            r_pr.append(element)
        element.set("typeface", FONT)
    return shape


def _append_notes(slide, text):
    if not text:
        return
    frame = slide.notes_slide.notes_text_frame
    existing = (frame.text or "").strip()
    frame.text = existing + "\n" + text if existing else text


def _open_deck(api, deck):
    prs = api["Presentation"](str(deck)) if deck else api["Presentation"]()
    if deck is None:
        prs.slide_width = api["Emu"](SLIDE_W)
        prs.slide_height = api["Emu"](SLIDE_H)
    return prs


def _target_slide(prs, slide_number):
    slides = list(prs.slides)
    if slide_number is None:
        slide = prs.slides.add_slide(_blank_layout(prs))
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = _pptx()["RGBColor"](0xFF, 0xFF, 0xFF)
        return slide, len(list(prs.slides))
    if slide_number < 1 or slide_number > len(slides):
        raise ValueError("slide %s is outside 1..%s" % (slide_number, len(slides)))
    return slides[slide_number - 1], slide_number


def embed_motion(
    output,
    video=None,
    poster=None,
    deck=None,
    slide=None,
    title=None,
    caption=None,
    notes=None,
    box=None,
):
    """Write ``output`` with one MP4 and/or a still, plus editable text.

    ``video`` is an existing file from ``manimgx render`` (``.mp4``).
    ``poster`` is an existing PNG or JPEG, usually the inspect storyboard.
    With a video and no poster, the still is a white placeholder, not a
    rendered frame. With a poster and no video, the still is an ordinary
    picture. At least one of the two is required.
    """
    video_path = Path(video) if video else None
    poster_path = Path(poster) if poster else None
    destination = Path(output)
    if deck is not None and destination.resolve() == Path(deck).resolve():
        raise ValueError("output must be a different file from the input deck")
    if video_path is None and poster_path is None:
        raise ValueError("pass a rendered .mp4, a poster image, or both")
    if video_path is not None:
        if video_path.suffix.lower() != ".mp4":
            raise ValueError("video must be an .mp4 from manimgx render")
        if not video_path.is_file():
            raise ValueError("video file not found: " + str(video_path))
    if poster_path is not None:
        if poster_path.suffix.lower() not in {".png", ".jpg", ".jpeg"}:
            raise ValueError("poster must be a .png or .jpeg file")
        if not poster_path.is_file():
            raise ValueError("poster file not found: " + str(poster_path))
    api = _pptx()
    prs = _open_deck(api, deck)
    target, number = _target_slide(prs, slide)
    had_timing = target._element.find(api["qn"]("p:timing")) is not None
    slide_w = int(prs.slide_width)
    slide_h = int(prs.slide_height)
    media_box = box if box is not None else _scale(MEDIA_BOX, slide_w, slide_h)
    if title:
        _add_text(target, _scale(TITLE_BOX, slide_w, slide_h), title, 32, ACCENT, api, bold=True)
    poster_kind = "supplied" if poster_path is not None else None
    if video_path is not None:
        if poster_path is None:
            poster_kind = "placeholder"
            poster_file = BytesIO(_white_png())
        else:
            poster_file = str(poster_path)
        target.shapes.add_movie(
            str(video_path),
            api["Emu"](media_box[0]),
            api["Emu"](media_box[1]),
            api["Emu"](media_box[2]),
            api["Emu"](media_box[3]),
            poster_frame_image=poster_file,
            mime_type="video/mp4",
        )
    else:
        poster_kind = "still"
        target.shapes.add_picture(
            str(poster_path),
            api["Emu"](media_box[0]),
            api["Emu"](media_box[1]),
            api["Emu"](media_box[2]),
            api["Emu"](media_box[3]),
        )
    if caption:
        _add_text(target, _scale(CAPTION_BOX, slide_w, slide_h), caption, 20, INK, api)
    _append_notes(target, notes)
    destination.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(destination))
    if had_timing and video_path is not None:
        print(
            "warning: slide %s already had timing; python-pptx may have replaced it while adding the video"
            % number,
            file=sys.stderr,
        )
    return {
        "output": str(destination),
        "slide": number,
        "video": video_path is not None,
        "poster": poster_kind,
        "timing_replaced_warning": bool(had_timing and video_path is not None),
    }


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("deck", nargs="?", help="Existing PPTX. Omit to start a blank 16:9 deck.")
    parser.add_argument("-o", "--output", help="Destination PPTX. Required when embedding.")
    parser.add_argument("--video", help="MP4 from manimgx render.")
    parser.add_argument("--poster", help="PNG or JPEG still, usually the inspect storyboard.")
    parser.add_argument("--slide", type=int, help="1-based slide to add onto. Default: append a slide.")
    parser.add_argument("--title", help="Editable slide title.")
    parser.add_argument("--caption", help="Editable caption under the clip.")
    parser.add_argument("--notes", help="Speaker notes. Appended when the slide already has notes.")
    parser.add_argument("--left", type=float, help="Media left edge, inches. Pass all four box edges together.")
    parser.add_argument("--top", type=float)
    parser.add_argument("--width", type=float)
    parser.add_argument("--height", type=float)
    parser.add_argument("--check", action="store_true", help="Print manimgx availability as JSON and exit 0.")
    parser.add_argument("--commands", metavar="SCENE.py", help="Print inspect and render argv as JSON. Does not run them.")
    parser.add_argument("--scene", help="Scene class name, only if SCENE.py defines several.")
    parser.add_argument("--storyboard", help="Passed to inspect -o when printing --commands.")
    parser.add_argument("--fps", type=int, default=RENDER_FPS)
    parser.add_argument("--resolution", default=RENDER_RESOLUTION)
    args = parser.parse_args(argv)
    try:
        if args.check:
            print(json.dumps(probe_manimgx(), ensure_ascii=False, indent=2))
            return 0
        if args.commands:
            print(json.dumps(planned_commands(
                args.commands,
                scene_name=args.scene,
                video_path=args.video,
                storyboard_path=args.storyboard,
                resolution=args.resolution,
                fps=args.fps,
            ), ensure_ascii=False, indent=2))
            return 0
        if not args.output:
            parser.error("embedding needs -o/--output, or pass --check / --commands")
        edges = (args.left, args.top, args.width, args.height)
        if any(edge is not None for edge in edges) and any(edge is None for edge in edges):
            raise ValueError("pass --left, --top, --width, and --height together")
        box = None
        if args.left is not None:
            box = tuple(int(round(edge * EMU_IN)) for edge in edges)
        result = embed_motion(
            args.output,
            video=args.video,
            poster=args.poster,
            deck=args.deck,
            slide=args.slide,
            title=args.title,
            caption=args.caption,
            notes=args.notes,
            box=box,
        )
        print(result["output"])
    except (OSError, ValueError, RuntimeError) as exc:
        print(type(exc).__name__ + ": " + str(exc), file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
