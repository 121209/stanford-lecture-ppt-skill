#!/usr/bin/env python3
"""Put a short math letter into a small editable OMML box.

Full equations stay on an independent ``{{MATH:id}}`` shape and
``inject_omml.py``. This helper is for a symbol that sits beside a sentence:
``d_k``, ``\\lambda``, ``\\lambda(t)``, ``\\mathrm{softmax}``. The box is still
the same contract as display math: ``mc:AlternateContent``, an ``a14`` choice,
and a transparent picture fallback. Do not put a Chinese sentence in the box.

The placeholder run is the whole shape, for example ``{{MATH:dk}}``. Keep the
shape small (about 1.4 in by 0.42 in for one symbol) and leave the Chinese in
the neighboring text box.
"""
import argparse
import importlib.util
import json
import re
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

CJK = re.compile(r"[\u3400-\u9fff]")
# Long enough for ``\mathrm{softmax}`` or ``\lambda(t)=\alpha``, short of a display.
MAX_INLINE_CHARS = 40


def _inject():
    name = "inject_omml"
    loaded = sys.modules.get(name)
    if loaded is not None and hasattr(loaded, "inject_pptx"):
        return loaded
    path = Path(__file__).resolve().parent / "inject_omml.py"
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


def assert_short_inline(latex):
    """Return the math string, or raise if it is a sentence or a full equation."""
    if not isinstance(latex, str) or not latex.strip():
        raise ValueError("inline latex must be a non-empty string")
    text = latex.strip()
    if text.startswith("$$"):
        raise ValueError("full equations use {{MATH:}}, not an inline box")
    inject = _inject()
    body = inject._unwrap_math_dollars(text)
    if not body:
        raise ValueError("inline latex must be a non-empty string")
    if CJK.search(body):
        raise ValueError("do not put a Chinese sentence in a math box")
    if r"\begin" in body or len(body) > MAX_INLINE_CHARS:
        raise ValueError(
            "inline math is only a short symbol; put full equations in {{MATH:}}"
        )
    return body


def prepare_inline_spec(latex):
    """Formula spec for one short symbol. Sets ``inline`` for the fallback size."""
    return {"latex": assert_short_inline(latex), "inline": True}


def normalize_inline_formulas(formulas):
    if not isinstance(formulas, dict) or not formulas:
        raise ValueError("inline formulas must be a non-empty object")
    prepared = {}
    for key, spec in formulas.items():
        if isinstance(spec, str):
            latex = spec
        elif isinstance(spec, dict) and isinstance(spec.get("latex"), str):
            latex = spec["latex"]
        else:
            raise ValueError("inline formula %s needs latex" % key)
        prepared[key] = prepare_inline_spec(latex)
    return prepared


def inject_inline_pptx(src, formulas, dest):
    """Write a new pptx. ``formulas`` maps id → latex string or ``{"latex": ...}``."""
    return _inject().inject_pptx(src, normalize_inline_formulas(formulas), dest)


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pptx", help="Source pptx containing {{MATH:id}} runs")
    parser.add_argument("formulas", help="JSON object of short latex specs")
    parser.add_argument("-o", "--output", required=True, help="New pptx path; never overwrites")
    args = parser.parse_args(argv)
    try:
        formulas = _inject().load_formulas(args.formulas)
        inject_inline_pptx(args.pptx, formulas, args.output)
    except (OSError, ValueError, ET.ParseError, zipfile.BadZipFile, json.JSONDecodeError) as exc:
        print(type(exc).__name__ + ": " + str(exc))
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
