#!/usr/bin/env python3
"""Read-only OOXML inventory. This is not a visual or scientific QA verdict."""
import argparse
import hashlib
import json
import posixpath
from pathlib import Path
import sys
import xml.etree.ElementTree as ET
import zipfile

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
}
EMU_PER_PX = 9525


def tag(prefix, local):
    return "{" + NS[prefix] + "}" + local


def resolve_part(owner, target):
    resolved = (target.lstrip("/") if target.startswith("/") else
                posixpath.normpath(posixpath.join(posixpath.dirname(owner), target)))
    if resolved.startswith("../") or resolved == "..":
        raise ValueError("Relationship escapes package root: " + target)
    return resolved


def relationships(zf, owner):
    relpart = posixpath.join(posixpath.dirname(owner), "_rels", posixpath.basename(owner) + ".rels")
    if relpart not in zf.namelist():
        return {}
    return {r.get("Id"): dict(r.attrib) for r in ET.fromstring(zf.read(relpart))}


def objects(container, nested=False):
    """Choose first AlternateContent Choice for inventory, not a renderer claim."""
    for item in container:
        if item.tag == tag("mc", "AlternateContent"):
            choice = item.find("mc:Choice", NS)
            if choice is None:
                choice = item.find("mc:Fallback", NS)
            if choice is not None:
                yield from objects(choice, nested)
        elif item.tag == tag("p", "grpSp"):
            yield from objects(item, True)
        elif item.tag in (tag("p", "sp"), tag("p", "pic"), tag("p", "graphicFrame"), tag("p", "cxnSp")):
            yield item, nested


def text_of(node):
    return "\n".join("".join(t.itertext()) for t in node.iter()
                     if t.tag in (tag("a", "t"), tag("m", "t")))


def geometry(node):
    xfrm = node.find("p:spPr/a:xfrm", NS)
    if xfrm is None:
        xfrm = node.find("p:xfrm", NS)
    if xfrm is None:
        return None
    off, ext = xfrm.find("a:off", NS), xfrm.find("a:ext", NS)
    if off is None or ext is None:
        return None
    return [int(off.get("x", 0)), int(off.get("y", 0)),
            int(ext.get("cx", 0)), int(ext.get("cy", 0))]


def notes_body(zf, slide_part):
    rels = relationships(zf, slide_part)
    rel = next((r for r in rels.values() if r.get("Type", "").endswith("/notesSlide")), None)
    if rel is None:
        return None, ""
    if rel.get("TargetMode") == "External":
        return None, ""
    part = resolve_part(slide_part, rel["Target"])
    notes = ET.fromstring(zf.read(part))
    texts = []
    for shape in notes.findall(".//p:sp", NS):
        ph = shape.find("p:nvSpPr/p:nvPr/p:ph", NS)
        if ph is not None and ph.get("type") in {"sldImg", "sldNum", "dt", "hdr", "ftr"}:
            continue
        text = text_of(shape).strip()
        if text:
            texts.append(text)
    return part, "\n".join(texts)


def audit(path):
    source = Path(path).resolve(strict=True)
    result = {
        "schema_version": 1,
        "file": str(source),
        "sha256": hashlib.sha256(source.read_bytes()).hexdigest(),
        "claim_boundary": (
            "Read-only package inventory. Counts and large-image hints do not prove editability, "
            "visual correctness, notes usefulness, or scientific accuracy. Render and inspect separately. "
            "AlternateContent uses its first Choice when present; app rendering may choose differently. "
            "Grouped and inherited coordinates are not geometrically validated."
        ),
        "warnings": [],
    }
    with zipfile.ZipFile(source) as zf:
        broken = zf.testzip()
        if broken:
            raise ValueError("Corrupt ZIP member: " + broken)
        names = zf.namelist()
        if len(names) != len(set(names)):
            result["warnings"].append({"code": "duplicate_zip_members"})
        # Check all relationships without following external targets or network requests.
        external = []
        for name in names:
            if not name.endswith(".rels"):
                continue
            for rel in ET.fromstring(zf.read(name)):
                if rel.get("TargetMode") == "External":
                    external.append({"part": name, "id": rel.get("Id"), "type": rel.get("Type"), "target": rel.get("Target")})
        result["external_relationships"] = external
        pres = ET.fromstring(zf.read("ppt/presentation.xml"))
        size = pres.find("p:sldSz", NS)
        if size is None:
            raise ValueError("Missing p:sldSz")
        sw, sh = int(size.get("cx")), int(size.get("cy"))
        result["slide_size_emu"] = [sw, sh]
        result["slide_size_css_px"] = [round(sw / EMU_PER_PX, 2), round(sh / EMU_PER_PX, 2)]
        rels = relationships(zf, "ppt/presentation.xml")
        pages = []
        for page, sid in enumerate(pres.findall("p:sldIdLst/p:sldId", NS), 1):
            rel = rels[sid.get(tag("r", "id"))]
            if rel.get("TargetMode") == "External":
                raise ValueError("Slide relationship must not be external")
            part = resolve_part("ppt/presentation.xml", rel["Target"])
            slide = ET.fromstring(zf.read(part))
            tree = slide.find("p:cSld/p:spTree", NS)
            if tree is None:
                raise ValueError("Missing shape tree: " + part)
            record = {"page": page, "part": part, "native_text_shape_count": 0,
                      "picture_count": 0, "native_table_count": 0, "native_chart_count": 0,
                      "office_math_count": 0, "native_texts": [], "large_picture_hints": [], "warnings": []}
            for shape, nested in objects(tree):
                text = text_of(shape)
                info = shape.find(".//p:cNvPr", NS)
                identity = {"id": info.get("id") if info is not None else None,
                            "name": info.get("name") if info is not None else None}
                box = geometry(shape)
                if shape.tag == tag("p", "sp") and text.strip():
                    record["native_text_shape_count"] += 1
                    entry = {**identity, "text": text, "nested_in_group": nested}
                    if box:
                        entry["box_css_px"] = [round(v / EMU_PER_PX, 2) for v in box]
                    record["native_texts"].append(entry)
                    if box and not nested:
                        x, y, w, h = box
                        if x < 0 or y < 0 or x+w > sw or y+h > sh:
                            record["warnings"].append({"code": "direct_text_frame_outside_canvas", **identity})
                record["native_table_count"] += len(shape.findall(".//a:tbl", NS))
                record["native_chart_count"] += len(shape.findall(".//c:chart", NS))
                record["office_math_count"] += len(shape.findall(".//m:oMath", NS))
                if shape.tag == tag("p", "pic"):
                    record["picture_count"] += 1
                    if box and not nested:
                        x, y, w, h = box
                        area = max(0, min(x+w, sw)-max(0, x)) * max(0, min(y+h, sh)-max(0, y)) / (sw*sh)
                        if area >= .75:
                            record["large_picture_hints"].append({**identity, "canvas_area_ratio": round(area, 4), "interpretation": "Inspect visually; may be legitimate source figure or flattened slide."})
            notes_part, notes = notes_body(zf, part)
            count = sum(not c.isspace() for c in notes)
            record.update({"notes_part": notes_part, "has_nonempty_body_notes": bool(count),
                           "notes_non_whitespace_char_count": count, "notes_text": notes})
            pages.append(record)
        result["slides"] = pages
        result["slide_count"] = len(pages)
        result["slides_without_body_notes"] = [p["page"] for p in pages if not p["has_nonempty_body_notes"]]
        result["slides_with_large_picture_hints"] = [p["page"] for p in pages if p["large_picture_hints"]]
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pptx", help="Existing source PPTX; never modified")
    parser.add_argument("--json", dest="json_path", help="Optional new report file (refuses overwrite)")
    args = parser.parse_args()
    try:
        report = audit(args.pptx)
        encoded = json.dumps(report, ensure_ascii=False, indent=2)
        if args.json_path:
            target = Path(args.json_path).resolve()
            if target == Path(args.pptx).resolve():
                raise ValueError("Report cannot overwrite the source PPTX")
            target.parent.mkdir(parents=True, exist_ok=True)
            with target.open("x", encoding="utf-8") as stream:
                stream.write(encoded + "\n")
            print(json.dumps({"report": str(target), "slide_count": report["slide_count"], "sha256": report["sha256"]}, ensure_ascii=False))
        else:
            print(encoded)
    except (OSError, ValueError, KeyError, ET.ParseError, zipfile.BadZipFile) as exc:
        print(json.dumps({"error": type(exc).__name__, "message": str(exc)}, ensure_ascii=False), file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
