#!/usr/bin/env python3
"""Report the local lecture-deck toolchain. Does not author slides.

Required: Python >= 3.10 and an importable python-pptx.
Optional: a TeX engine, pdftocairo, LibreOffice, and ManimGX.

Exit 0 when the required pieces are present, even if every optional tool is
missing. The report names what is missing so callers can degrade the way the
skill already does: mathtext instead of a TeX PNG, skip render checks, and
leave motion on editable OMML or native shapes.

``--write`` stores a secret-free runtime profile. It contains no API keys and
no machine path that a caller must satisfy. PowerPoint, COM, WPS, and paid
APIs are not probed and are not required.
"""
import argparse
import importlib.util
import json
import os
import shutil
import sys
from pathlib import Path

SCHEMA = "stanford-lecture-ppt-runtime-profile"
SCHEMA_VERSION = 1
MIN_PYTHON = (3, 10)
ENGINES = ("xelatex", "lualatex", "pdflatex")
PNG_CONVERTERS = ("pdftocairo", "ghostscript", "magick", "convert")
OFFICE_COMMANDS = ("soffice", "libreoffice")


def _sibling(name):
    cached = getattr(_sibling, name, None)
    if cached is not None:
        return cached
    path = Path(__file__).resolve().parent / (name + ".py")
    spec = importlib.util.spec_from_file_location("_doctor_" + name, path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    setattr(_sibling, name, module)
    return module


def python_status():
    version = ".".join(str(part) for part in sys.version_info[:3])
    return {
        "version": version,
        "minimum": "%d.%d" % MIN_PYTHON[:2],
        "ok": sys.version_info >= MIN_PYTHON,
    }


def pptx_status():
    try:
        import pptx
    except ImportError:
        return {"available": False, "version": None}
    version = getattr(pptx, "__version__", None)
    try:
        from importlib.metadata import version as package_version
        version = package_version("python-pptx")
    except Exception:
        pass
    if not isinstance(version, str) or not version or "/" in version or "\\" in version:
        version = None
    return {"available": True, "version": version}


def tex_status():
    status = _sibling("tex_fallback").toolchain_status()
    engine = status.get("engine")
    if engine not in ENGINES:
        engine = None
    converter = status.get("png_converter")
    if converter not in PNG_CONVERTERS:
        converter = None
    return {
        "available": engine is not None,
        "engine": engine,
        "png_converter": converter,
        "tex_png": bool(status.get("tex_png")) and engine is not None and converter is not None,
        "disabled_by_env": os.environ.get("STANFORD_FORMULA_DISABLE_TEX") == "1",
    }


def pdftocairo_status():
    return {"available": shutil.which("pdftocairo") is not None}


def libreoffice_status():
    for name in OFFICE_COMMANDS:
        if shutil.which(name):
            return {"available": True, "command": name}
    return {"available": False, "command": None}


def manimgx_status():
    info = _sibling("embed_motion").probe_manimgx()
    version = info.get("version")
    if not isinstance(version, str):
        version = None
    else:
        version = " ".join(version.split())
        if not version or len(version) > 80 or "/" in version or "\\" in version:
            version = None
    return {
        "available": bool(info.get("available")),
        "version": version,
        "python_required": "3.13 or 3.14",
    }


def _missing(optional):
    missing = []
    if not optional["tex"]["available"]:
        missing.append("tex")
    if not optional["pdftocairo"]["available"]:
        missing.append("pdftocairo")
    if not optional["libreoffice"]["available"]:
        missing.append("libreoffice")
    if not optional["manimgx"]["available"]:
        missing.append("manimgx")
    return missing


def _degrade(optional):
    return {
        "formula_picture": "tex-png" if optional["tex"]["tex_png"] else "mathtext",
        "render_checks": "libreoffice" if optional["libreoffice"]["available"] else "skip",
        "motion": "manimgx" if optional["manimgx"]["available"] else "editable-only",
    }


def collect_profile():
    """Build the runtime profile from the current process. No secrets, no paths."""
    python = python_status()
    pptx = pptx_status()
    optional = {
        "tex": tex_status(),
        "pdftocairo": pdftocairo_status(),
        "libreoffice": libreoffice_status(),
        "manimgx": manimgx_status(),
    }
    return {
        "schema": SCHEMA,
        "schema_version": SCHEMA_VERSION,
        "secret_free": True,
        "ok": bool(python["ok"] and pptx["available"]),
        "python": python,
        "required": {"python_pptx": pptx},
        "optional": optional,
        "missing_optional": _missing(optional),
        "degrade": _degrade(optional),
    }


def render_text(profile):
    python = profile["python"]
    pptx = profile["required"]["python_pptx"]
    tex = profile["optional"]["tex"]
    cairo = profile["optional"]["pdftocairo"]
    office = profile["optional"]["libreoffice"]
    motion = profile["optional"]["manimgx"]
    lines = [
        "stanford-lecture-ppt doctor",
        "required:",
        "  python %s (>= %s) %s" % (
            python["version"], python["minimum"], "ok" if python["ok"] else "missing",
        ),
    ]
    if pptx["available"]:
        lines.append("  python-pptx %s ok" % (pptx["version"] or "unknown"))
    else:
        lines.append("  python-pptx missing")
    lines.append("optional:")
    if tex["available"]:
        lines.append("  tex: %s (png converter: %s)" % (
            tex["engine"], tex["png_converter"] or "none",
        ))
    else:
        lines.append("  tex: missing")
    lines.append("  pdftocairo: %s" % ("ok" if cairo["available"] else "missing"))
    if office["available"]:
        lines.append("  libreoffice: %s" % office["command"])
    else:
        lines.append("  libreoffice: missing")
    if motion["available"]:
        lines.append("  manimgx: %s" % (motion["version"] or "ok"))
    else:
        lines.append("  manimgx: missing")
    missing = profile["missing_optional"]
    lines.append("missing optional: %s" % (", ".join(missing) if missing else "none"))
    degrade = profile["degrade"]
    lines.append(
        "degrade: formula picture=%s; render checks=%s; motion=%s" % (
            degrade["formula_picture"], degrade["render_checks"], degrade["motion"],
        )
    )
    lines.append("status: %s" % ("ok" if profile["ok"] else "required-missing"))
    return "\n".join(lines)


def _write_profile(path, encoded):
    target = Path(path)
    if target.suffix.lower() == ".pptx":
        raise ValueError("runtime profile cannot overwrite a pptx")
    target.parent.mkdir(parents=True, exist_ok=True)
    temporary = target.with_name(target.name + ".tmp")
    temporary.write_text(encoded + "\n", encoding="utf-8")
    temporary.replace(target)


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--json", action="store_true", help="Print the runtime profile as JSON")
    parser.add_argument("--write", metavar="PATH", help="Write the secret-free profile JSON to PATH")
    args = parser.parse_args(argv)
    profile = collect_profile()
    encoded = json.dumps(profile, ensure_ascii=False, indent=2)
    try:
        if args.write:
            _write_profile(args.write, encoded)
        if args.json:
            print(encoded)
        else:
            print(render_text(profile))
    except (OSError, ValueError) as exc:
        print(type(exc).__name__ + ": " + str(exc), file=sys.stderr)
        return 2
    return 0 if profile["ok"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
