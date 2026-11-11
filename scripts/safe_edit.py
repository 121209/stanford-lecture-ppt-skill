#!/usr/bin/env python3
"""Copy an existing PPTX before editing it. Stdlib only.

Opening a deck for edits saves a working copy and records the source sha256.
The source bytes stay untouched unless the caller passes in_place=True (CLI:
``--in-place``). Any other value, including an omitted flag, is not
authorization to overwrite.

Helpers that already write a separate output path stay on that contract.
This script does not call PowerPoint, COM, or WPS, and it does not read keys.
``run_state.py`` still only records production gates; it does not copy decks.
"""
import argparse
import hashlib
import json
import shutil
import sys
from pathlib import Path


def file_sha256(path):
    digest = hashlib.sha256()
    with Path(path).open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _unused_copy_path(source):
    stem = source.stem
    suffix = source.suffix
    parent = source.parent
    names = [stem + ".work" + suffix]
    names.extend("%s.work-%d%s" % (stem, number, suffix) for number in range(2, 1000))
    for name in names:
        candidate = parent / name
        if candidate.resolve() == source.resolve():
            continue
        if not candidate.exists():
            return candidate
    raise ValueError("No free working-copy name next to " + source.name)


def _write_record(path, info):
    target = Path(path)
    if target.resolve() == Path(info["source"]).resolve():
        raise ValueError("record path must not be the source pptx")
    if target.resolve() == Path(info["working_copy"]).resolve():
        raise ValueError("record path must not be the working pptx")
    target.parent.mkdir(parents=True, exist_ok=True)
    temporary = target.with_name(target.name + ".tmp")
    temporary.write_text(json.dumps(info, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    temporary.replace(target)


def prepare_working_copy(source, dest=None, in_place=False, record=None):
    """Return a working PPTX and the source sha256.

    ``in_place`` must be the boolean True. A missing flag, False, or any
    non-boolean does not authorize overwriting the source.
    """
    if not isinstance(in_place, bool):
        raise ValueError("in_place must be a boolean; only True authorizes overwrite")
    if in_place and dest is not None:
        raise ValueError("do not pass dest with in_place")
    path = Path(source)
    if path.suffix.lower() != ".pptx":
        raise ValueError("source must be a .pptx file")
    if not path.is_file():
        raise ValueError("Missing pptx: " + str(path))
    digest = file_sha256(path)
    if in_place:
        target = path
        copied = False
    else:
        target = Path(dest) if dest is not None else _unused_copy_path(path)
        if target.suffix.lower() != ".pptx":
            raise ValueError("working copy must be a .pptx file")
        if target.resolve() == path.resolve():
            raise ValueError("dest is the source; pass in_place=True only when overwrite is authorized")
        if target.exists():
            raise ValueError("working copy already exists: " + target.name)
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(path, target)
        copied_hash = file_sha256(target)
        if copied_hash != digest:
            target.unlink()
            raise ValueError("Working copy hash does not match the source")
        copied = True
    info = {
        "schema_version": 1,
        "source": str(path.resolve()),
        "sha256": digest,
        "working_copy": str(target.resolve()),
        "in_place": in_place,
        "copied": copied,
    }
    if record is not None:
        _write_record(record, info)
    return info


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", help="Existing PPTX. Not modified unless --in-place is set.")
    parser.add_argument("--dest", help="Working-copy path. Default: <name>.work.pptx beside the source.")
    parser.add_argument(
        "--in-place",
        action="store_true",
        help="Authorize overwrite of the source. Omit to save a working copy.",
    )
    parser.add_argument("--record", help="Write the source-hash JSON here.")
    args = parser.parse_args(argv)
    try:
        info = prepare_working_copy(
            args.source,
            dest=args.dest,
            in_place=args.in_place,
            record=args.record,
        )
    except (OSError, ValueError) as exc:
        print(type(exc).__name__ + ": " + str(exc), file=sys.stderr)
        return 2
    print(json.dumps(info, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
