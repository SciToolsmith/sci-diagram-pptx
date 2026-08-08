#!/usr/bin/env python3
"""Read-only OOXML audit for a sci-diagram-pptx deliverable.

The audit deliberately inspects the ZIP package and XML instead of relying on
PowerPoint automation.  It proves only properties visible in OOXML; anything
that cannot be established from the package is reported as WARN, never PASS.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import posixpath
import re
import sys
import unicodedata
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable
from urllib.parse import unquote
from xml.etree import ElementTree as ET


TOOL_VERSION = "1.0"
EMU_PER_INCH = 914400

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}
R_ID = f"{{{NS['r']}}}id"
R_EMBED = f"{{{NS['r']}}}embed"
R_LINK = f"{{{NS['r']}}}link"
REL_NS = NS["rel"]

IMAGE_REL_SUFFIX = "/image"
SLIDE_REL_SUFFIX = "/slide"
SLIDE_LAYOUT_REL_SUFFIX = "/slideLayout"
SLIDE_MASTER_REL_SUFFIX = "/slideMaster"
VECTOR_EXTENSIONS = {".svg", ".emf", ".wmf"}
FALSE_VALUES = {"0", "false", "off", "no"}
SHA256_RE = re.compile(r"^[0-9a-fA-F]{64}$")


class AuditError(RuntimeError):
    """An operational failure that prevents a trustworthy audit."""


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")


def load_json(path: Path, label: str) -> dict[str, Any]:
    try:
        value = json.loads(path.expanduser().resolve().read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise AuditError(f"Cannot read {label} JSON '{path}': {exc}") from exc
    if not isinstance(value, dict):
        raise AuditError(f"{label} must contain one JSON object")
    return value


def write_json(payload: dict[str, Any], output: Path) -> None:
    target = output.expanduser().resolve()
    target.parent.mkdir(parents=True, exist_ok=True)
    temporary = target.with_name(target.name + ".tmp")
    temporary.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    os.replace(temporary, target)


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def normalise_part_target(owner_part: str, target: str) -> str | None:
    """Resolve an OOXML relationship target to a package-relative part name."""
    target = unquote(target).replace("\\", "/")
    if target.startswith("/"):
        resolved = posixpath.normpath(target.lstrip("/"))
    else:
        resolved = posixpath.normpath(posixpath.join(posixpath.dirname(owner_part), target))
    if resolved == ".." or resolved.startswith("../"):
        return None
    return resolved


def relationship_part(owner_part: str) -> str:
    folder, name = posixpath.split(owner_part)
    return posixpath.join(folder, "_rels", name + ".rels")


def parse_xml(zf: zipfile.ZipFile, part: str) -> ET.Element:
    try:
        return ET.fromstring(zf.read(part))
    except KeyError as exc:
        raise AuditError(f"Required OOXML part is missing: {part}") from exc
    except ET.ParseError as exc:
        raise AuditError(f"Malformed XML in {part}: {exc}") from exc


def read_relationships(zf: zipfile.ZipFile, owner_part: str) -> tuple[dict[str, dict[str, Any]], list[str]]:
    rel_part = relationship_part(owner_part)
    if rel_part not in zf.namelist():
        return {}, [f"Relationship part is missing: {rel_part}"]
    root = parse_xml(zf, rel_part)
    relationships: dict[str, dict[str, Any]] = {}
    warnings: list[str] = []
    for rel in root.findall(f"{{{REL_NS}}}Relationship"):
        rid = rel.get("Id")
        target = rel.get("Target")
        rel_type = rel.get("Type")
        external = (rel.get("TargetMode") or "").lower() == "external"
        if not rid or not target or not rel_type:
            warnings.append(f"Incomplete Relationship entry in {rel_part}")
            continue
        resolved = None if external else normalise_part_target(owner_part, target)
        relationships[rid] = {
            "id": rid,
            "type": rel_type,
            "target": target,
            "resolved_target": resolved,
            "external": external,
            "relationship_part": rel_part,
        }
    return relationships, warnings


def package_external_relationships(zf: zipfile.ZipFile) -> tuple[list[dict[str, Any]], list[str]]:
    found: list[dict[str, Any]] = []
    warnings: list[str] = []
    for name in sorted(item for item in zf.namelist() if item.endswith(".rels")):
        try:
            root = parse_xml(zf, name)
        except AuditError as exc:
            warnings.append(str(exc))
            continue
        for rel in root.findall(f"{{{REL_NS}}}Relationship"):
            if (rel.get("TargetMode") or "").lower() == "external":
                found.append({
                    "relationship_part": name,
                    "id": rel.get("Id"),
                    "type": rel.get("Type"),
                    "target": rel.get("Target"),
                })
    return found, warnings


def package_relationship_type_hits(zf: zipfile.ZipFile, suffixes: tuple[str, ...]) -> tuple[list[dict[str, Any]], list[str]]:
    found: list[dict[str, Any]] = []
    warnings: list[str] = []
    for name in sorted(item for item in zf.namelist() if item.endswith(".rels")):
        try:
            root = parse_xml(zf, name)
        except AuditError as exc:
            warnings.append(str(exc))
            continue
        for rel in root.findall(f"{{{REL_NS}}}Relationship"):
            rel_type = rel.get("Type") or ""
            if any(rel_type.endswith(suffix) for suffix in suffixes):
                found.append({
                    "relationship_part": name,
                    "id": rel.get("Id"),
                    "type": rel_type,
                    "target": rel.get("Target"),
                    "external": (rel.get("TargetMode") or "").lower() == "external",
                })
    return found, warnings


def bool_is_false(value: str | None) -> bool:
    return value is not None and value.strip().lower() in FALSE_VALUES


def int_attr(node: ET.Element | None, name: str) -> int | None:
    if node is None:
        return None
    try:
        return int(node.get(name, ""))
    except ValueError:
        return None


def xfrm_box(xfrm: ET.Element | None) -> tuple[int, int, int, int] | None:
    if xfrm is None:
        return None
    off = xfrm.find("a:off", NS)
    ext = xfrm.find("a:ext", NS)
    values = (int_attr(off, "x"), int_attr(off, "y"), int_attr(ext, "cx"), int_attr(ext, "cy"))
    if any(value is None for value in values):
        return None
    x, y, width, height = values
    return int(x), int(y), int(width), int(height)


def box_dict(box: tuple[int, int, int, int] | None) -> dict[str, int] | None:
    if box is None:
        return None
    return {"x_emu": box[0], "y_emu": box[1], "width_emu": box[2], "height_emu": box[3]}


def picture_geometry(slide_root: ET.Element) -> tuple[list[dict[str, Any]], list[str]]:
    """Return picture boxes in slide coordinates, resolving scale/translation groups.

    Rotated/flipped groups and pictures are retained as uncertain evidence because
    their axis-aligned coverage cannot be proven by this lightweight parser.
    """
    sp_tree = slide_root.find("./p:cSld/p:spTree", NS)
    if sp_tree is None:
        return [], ["Slide has no p:cSld/p:spTree"]
    pictures: list[dict[str, Any]] = []
    warnings: list[str] = []

    def walk(parent: ET.Element, sx: float, sy: float, bx: float, by: float, uncertain: bool) -> None:
        for child in list(parent):
            local = child.tag.rsplit("}", 1)[-1]
            if local == "pic":
                xfrm = child.find("./p:spPr/a:xfrm", NS)
                box = xfrm_box(xfrm)
                item_uncertain = uncertain
                if xfrm is not None and (xfrm.get("rot") not in (None, "0") or xfrm.get("flipH") or xfrm.get("flipV")):
                    item_uncertain = True
                if box is None:
                    item_uncertain = True
                    warnings.append("A picture has no complete explicit transform; inherited geometry was not guessed")
                    slide_box = None
                else:
                    x, y, width, height = box
                    slide_box = {
                        "x_emu": round(sx * x + bx),
                        "y_emu": round(sy * y + by),
                        "width_emu": round(abs(sx) * width),
                        "height_emu": round(abs(sy) * height),
                    }
                blips = child.findall(".//a:blip", NS)
                relation_ids = sorted({rid for blip in blips for rid in (blip.get(R_EMBED), blip.get(R_LINK)) if rid})
                non_visual = child.find("./p:nvPicPr/p:cNvPr", NS)
                src_rect = child.find("./p:blipFill/a:srcRect", NS)
                crop = {
                    edge: int_attr(src_rect, edge) or 0
                    for edge in ("l", "t", "r", "b")
                }
                fill_rect_node = child.find("./p:blipFill/a:stretch/a:fillRect", NS)
                fill_rect = {
                    edge: int_attr(fill_rect_node, edge) or 0
                    for edge in ("l", "t", "r", "b")
                }
                alpha_values: list[int] = []
                for alpha in child.findall(".//a:alpha", NS) + child.findall(".//a:alphaModFix", NS):
                    value = int_attr(alpha, "val")
                    if value is None:
                        value = int_attr(alpha, "amt")
                    if value is not None:
                        alpha_values.append(value)
                pictures.append({
                    "name": non_visual.get("name") if non_visual is not None else None,
                    "box": slide_box,
                    "relationship_ids": relation_ids,
                    "geometry_uncertain": item_uncertain,
                    "crop": crop,
                    "fill_rect": fill_rect,
                    "cropped": any(crop.values()) or any(fill_rect.values()),
                    "tiled": child.find("./p:blipFill/a:tile", NS) is not None,
                    "hidden": bool_is_false(child.get("show")) or (
                        non_visual is not None
                        and ((non_visual.get("hidden") or "").lower() in {"1", "true", "on", "yes"} or bool_is_false(non_visual.get("show")))
                    ),
                    "fully_transparent": bool(alpha_values) and min(alpha_values) <= 0,
                })
            elif local == "grpSp":
                group_xfrm = child.find("./p:grpSpPr/a:xfrm", NS)
                group_uncertain = uncertain
                if group_xfrm is None:
                    warnings.append("A nested group has no transform; picture geometry inside it is uncertain")
                    walk(child, sx, sy, bx, by, True)
                    continue
                if group_xfrm.get("rot") not in (None, "0") or group_xfrm.get("flipH") or group_xfrm.get("flipV"):
                    group_uncertain = True
                    warnings.append("A rotated or flipped group contains content whose coverage is uncertain")
                off = group_xfrm.find("a:off", NS)
                ext = group_xfrm.find("a:ext", NS)
                child_off = group_xfrm.find("a:chOff", NS)
                child_ext = group_xfrm.find("a:chExt", NS)
                values = (
                    int_attr(off, "x"), int_attr(off, "y"), int_attr(ext, "cx"), int_attr(ext, "cy"),
                    int_attr(child_off, "x"), int_attr(child_off, "y"), int_attr(child_ext, "cx"), int_attr(child_ext, "cy"),
                )
                if any(value is None for value in values) or values[6] == 0 or values[7] == 0:
                    warnings.append("A group has an incomplete transform; picture geometry inside it is uncertain")
                    walk(child, sx, sy, bx, by, True)
                    continue
                ox, oy, ew, eh, cox, coy, cew, ceh = (float(value) for value in values)
                local_sx, local_sy = ew / cew, eh / ceh
                walk(
                    child,
                    sx * local_sx,
                    sy * local_sy,
                    sx * (ox - cox * local_sx) + bx,
                    sy * (oy - coy * local_sy) + by,
                    group_uncertain,
                )

    walk(sp_tree, 1.0, 1.0, 0.0, 0.0, False)
    return pictures, warnings


def native_shape_geometry(slide_root: ET.Element) -> tuple[list[dict[str, Any]], list[str]]:
    """Resolve p:sp boxes through non-rotated group transforms."""
    sp_tree = slide_root.find("./p:cSld/p:spTree", NS)
    if sp_tree is None:
        return [], ["Slide has no p:cSld/p:spTree for shape geometry"]
    shapes: list[dict[str, Any]] = []
    warnings: list[str] = []

    def walk(parent: ET.Element, sx: float, sy: float, bx: float, by: float, uncertain: bool) -> None:
        for child in list(parent):
            local = child.tag.rsplit("}", 1)[-1]
            if local == "sp":
                xfrm = child.find("./p:spPr/a:xfrm", NS)
                box = xfrm_box(xfrm)
                item_uncertain = uncertain or box is None
                if xfrm is not None and (xfrm.get("rot") not in (None, "0") or xfrm.get("flipH") or xfrm.get("flipV")):
                    item_uncertain = True
                slide_box = None
                if box is not None:
                    x, y, shape_width, shape_height = box
                    slide_box = {
                        "x_emu": round(sx * x + bx),
                        "y_emu": round(sy * y + by),
                        "width_emu": round(abs(sx) * shape_width),
                        "height_emu": round(abs(sy) * shape_height),
                    }
                non_visual = child.find("./p:nvSpPr/p:cNvPr", NS)
                blips = child.findall("./p:spPr/a:blipFill//a:blip", NS)
                relationship_ids = sorted({rid for blip in blips for rid in (blip.get(R_EMBED), blip.get(R_LINK)) if rid})
                alpha_values: list[int] = []
                for alpha in child.findall("./p:spPr//a:alpha", NS) + child.findall("./p:spPr//a:alphaModFix", NS):
                    value = int_attr(alpha, "val")
                    if value is None:
                        value = int_attr(alpha, "amt")
                    if value is not None:
                        alpha_values.append(value)
                shapes.append({
                    "name": non_visual.get("name") if non_visual is not None else None,
                    "box": slide_box,
                    "geometry_uncertain": item_uncertain,
                    "image_fill": bool(blips),
                    "relationship_ids": relationship_ids,
                    "no_fill": child.find("./p:spPr/a:noFill", NS) is not None,
                    "fully_transparent_fill": bool(alpha_values) and min(alpha_values) <= 0,
                    "hidden": bool_is_false(child.get("show")) or (
                        non_visual is not None
                        and ((non_visual.get("hidden") or "").lower() in {"1", "true", "on", "yes"} or bool_is_false(non_visual.get("show")))
                    ),
                })
            elif local == "grpSp":
                xfrm = child.find("./p:grpSpPr/a:xfrm", NS)
                group_uncertain = uncertain
                if xfrm is None or xfrm.get("rot") not in (None, "0") or xfrm.get("flipH") or xfrm.get("flipV"):
                    walk(child, sx, sy, bx, by, True)
                    continue
                off, ext = xfrm.find("a:off", NS), xfrm.find("a:ext", NS)
                child_off, child_ext = xfrm.find("a:chOff", NS), xfrm.find("a:chExt", NS)
                values = (
                    int_attr(off, "x"), int_attr(off, "y"), int_attr(ext, "cx"), int_attr(ext, "cy"),
                    int_attr(child_off, "x"), int_attr(child_off, "y"), int_attr(child_ext, "cx"), int_attr(child_ext, "cy"),
                )
                if any(value is None for value in values) or values[6] == 0 or values[7] == 0:
                    walk(child, sx, sy, bx, by, True)
                    continue
                ox, oy, ew, eh, cox, coy, cew, ceh = (float(value) for value in values)
                local_sx, local_sy = ew / cew, eh / ceh
                walk(child, sx * local_sx, sy * local_sy, sx * (ox - cox * local_sx) + bx, sy * (oy - coy * local_sy) + by, group_uncertain)

    walk(sp_tree, 1.0, 1.0, 0.0, 0.0, False)
    return shapes, warnings


def box_covers_slide(box: dict[str, int] | None, width: int, height: int, tolerance: float) -> bool:
    if not box or width <= 0 or height <= 0:
        return False
    tol_x, tol_y = width * tolerance, height * tolerance
    x, y = box["x_emu"], box["y_emu"]
    right, bottom = x + box["width_emu"], y + box["height_emu"]
    return x <= tol_x and y <= tol_y and right >= width - tol_x and bottom >= height - tol_y


def clipped_box(box: dict[str, int] | None, width: int, height: int) -> tuple[int, int, int, int] | None:
    if not box:
        return None
    x1 = max(0, box["x_emu"])
    y1 = max(0, box["y_emu"])
    x2 = min(width, box["x_emu"] + box["width_emu"])
    y2 = min(height, box["y_emu"] + box["height_emu"])
    return (x1, y1, x2, y2) if x2 > x1 and y2 > y1 else None


def rectangle_union_fraction(boxes: list[dict[str, int]], width: int, height: int) -> float:
    """Exact union area for axis-aligned picture boxes clipped to the slide."""
    clipped = [item for item in (clipped_box(box, width, height) for box in boxes) if item]
    if not clipped or width <= 0 or height <= 0:
        return 0.0
    x_values = sorted({coordinate for box in clipped for coordinate in (box[0], box[2])})
    area = 0
    for left, right in zip(x_values, x_values[1:]):
        intervals = sorted((top, bottom) for x1, top, x2, bottom in clipped if x1 < right and x2 > left)
        if not intervals:
            continue
        covered_y = 0
        start, end = intervals[0]
        for next_start, next_end in intervals[1:]:
            if next_start <= end:
                end = max(end, next_end)
            else:
                covered_y += end - start
                start, end = next_start, next_end
        covered_y += end - start
        area += (right - left) * covered_y
    return area / (width * height)


def background_for_part(zf: zipfile.ZipFile, owner_part: str, root: ET.Element) -> dict[str, Any] | None:
    bg = root.find("./p:cSld/p:bg", NS)
    if bg is None:
        return None
    rels, rel_warnings = read_relationships(zf, owner_part)
    blips = bg.findall(".//a:blip", NS)
    relation_ids = sorted({rid for blip in blips for rid in (blip.get(R_EMBED), blip.get(R_LINK)) if rid})
    image_relations = [rels[rid] for rid in relation_ids if rid in rels and rels[rid]["type"].endswith(IMAGE_REL_SUFFIX)]
    unresolved_ids = [rid for rid in relation_ids if rid not in rels]
    return {
        "owner_part": owner_part,
        "has_blip_fill": bool(bg.findall(".//a:blipFill", NS) or blips),
        "has_background_reference": bg.find("p:bgRef", NS) is not None,
        "image_relationships": image_relations,
        "unresolved_relationship_ids": unresolved_ids,
        "warnings": rel_warnings,
    }


def effective_background(zf: zipfile.ZipFile, slide_part: str, slide_root: ET.Element) -> dict[str, Any]:
    """Inspect explicit slide/layout/master background without inventing theme fills."""
    chain: list[tuple[str, ET.Element]] = [(slide_part, slide_root)]
    warnings: list[str] = []
    current_part, current_root = slide_part, slide_root
    for suffix in (SLIDE_LAYOUT_REL_SUFFIX, SLIDE_MASTER_REL_SUFFIX):
        rels, rel_warnings = read_relationships(zf, current_part)
        warnings.extend(rel_warnings)
        candidates = [rel for rel in rels.values() if rel["type"].endswith(suffix) and not rel["external"]]
        if not candidates:
            # A slide may be deliberately self-contained.  A missing layout link is
            # unusual but does not prove an image background.
            break
        target = candidates[0]["resolved_target"]
        if not target or target not in zf.namelist():
            warnings.append(f"Cannot resolve background inheritance target from {current_part}")
            break
        try:
            current_root = parse_xml(zf, target)
        except AuditError as exc:
            warnings.append(str(exc))
            break
        current_part = target
        chain.append((current_part, current_root))
    for owner_part, root in chain:
        background = background_for_part(zf, owner_part, root)
        if background is not None:
            background["source_level"] = "slide" if owner_part == slide_part else "layout" if "slideLayout" in owner_part else "master"
            background["warnings"].extend(warnings)
            return background
    return {
        "owner_part": None,
        "source_level": "default_or_theme",
        "has_blip_fill": False,
        "has_background_reference": False,
        "image_relationships": [],
        "unresolved_relationship_ids": [],
        "warnings": warnings,
    }


def slide_analysis(
    zf: zipfile.ZipFile,
    slide_part: str,
    index: int,
    width: int,
    height: int,
    full_bleed_tolerance: float,
) -> dict[str, Any]:
    root = parse_xml(zf, slide_part)
    rels, rel_warnings = read_relationships(zf, slide_part)
    pictures, geometry_warnings = picture_geometry(root)
    native_shapes, native_shape_warnings = native_shape_geometry(root)
    for picture in pictures:
        picture["full_bleed"] = False if picture["geometry_uncertain"] else box_covers_slide(picture["box"], width, height, full_bleed_tolerance)
        clip = clipped_box(picture["box"], width, height)
        if picture["box"] and width > 0 and height > 0:
            original_area = max(0, picture["box"]["width_emu"]) * max(0, picture["box"]["height_emu"])
            clipped_area = 0 if clip is None else (clip[2] - clip[0]) * (clip[3] - clip[1])
            picture["slide_coverage_fraction"] = clipped_area / (width * height)
            picture["visible_fraction_of_picture"] = clipped_area / original_area if original_area else 0.0
            picture["off_canvas"] = clip is None
            picture["partially_off_canvas"] = clip is not None and clipped_area < original_area
        else:
            picture["slide_coverage_fraction"] = None
            picture["visible_fraction_of_picture"] = None
            picture["off_canvas"] = None
            picture["partially_off_canvas"] = None
        picture["relationships"] = [rels[rid] for rid in picture["relationship_ids"] if rid in rels]
        picture["unresolved_relationship_ids"] = [rid for rid in picture["relationship_ids"] if rid not in rels]
    for shape in native_shapes:
        clip = clipped_box(shape["box"], width, height)
        if shape["box"] and width > 0 and height > 0:
            original_area = max(0, shape["box"]["width_emu"]) * max(0, shape["box"]["height_emu"])
            clipped_area = 0 if clip is None else (clip[2] - clip[0]) * (clip[3] - clip[1])
            shape["slide_coverage_fraction"] = clipped_area / (width * height)
            shape["off_canvas"] = clip is None
            shape["partially_off_canvas"] = clip is not None and clipped_area < original_area
        else:
            shape["slide_coverage_fraction"] = None
            shape["off_canvas"] = None
            shape["partially_off_canvas"] = None
    text_values = [node.text or "" for node in root.findall(".//a:t", NS)]
    named_objects: list[dict[str, Any]] = []
    object_paths = (
        ("shape", ".//p:sp", "./p:nvSpPr/p:cNvPr"),
        ("connector", ".//p:cxnSp", "./p:nvCxnSpPr/p:cNvPr"),
        ("group", ".//p:grpSp", "./p:nvGrpSpPr/p:cNvPr"),
        ("picture", ".//p:pic", "./p:nvPicPr/p:cNvPr"),
        ("graphic_frame", ".//p:graphicFrame", "./p:nvGraphicFramePr/p:cNvPr"),
    )
    for object_type, xpath, name_xpath in object_paths:
        for node in root.findall(xpath, NS):
            non_visual = node.find(name_xpath, NS)
            start_connection = node.find("./p:nvCxnSpPr/p:cNvCxnSpPr/a:stCxn", NS) if object_type == "connector" else None
            end_connection = node.find("./p:nvCxnSpPr/p:cNvCxnSpPr/a:endCxn", NS) if object_type == "connector" else None
            xfrm_paths = {
                "shape": "./p:spPr/a:xfrm",
                "connector": "./p:spPr/a:xfrm",
                "picture": "./p:spPr/a:xfrm",
                "graphic_frame": "./p:xfrm",
                "group": "./p:grpSpPr/a:xfrm",
            }
            paragraphs = []
            own_text_path = "./p:txBody/a:p" if object_type in {"shape", "connector"} else "./a:txBody/a:p" if object_type == "graphic_frame" else None
            for paragraph in node.findall(own_text_path, NS) if own_text_path else []:
                paragraphs.append("".join(text.text or "" for text in paragraph.findall(".//a:t", NS)))
            line = node.find("./p:spPr/a:ln", NS) if object_type == "connector" else None
            head = line.find("a:headEnd", NS) if line is not None else None
            tail = line.find("a:tailEnd", NS) if line is not None else None
            dash = line.find("a:prstDash", NS) if line is not None else None
            solid_fill = line.find("a:solidFill", NS) if line is not None else None
            srgb = solid_fill.find("a:srgbClr", NS) if solid_fill is not None else None
            scheme = solid_fill.find("a:schemeClr", NS) if solid_fill is not None else None
            shape_properties = node.find("./p:spPr", NS) if object_type in {"shape", "connector", "picture"} else None
            preset_geometry = shape_properties.find("a:prstGeom", NS) if shape_properties is not None else None
            custom_geometry = shape_properties.find("a:custGeom", NS) if shape_properties is not None else None
            shape_geometry: str | None = None
            shape_geometry_source: str | None = None
            if object_type == "shape":
                non_visual_shape = node.find("./p:nvSpPr/p:cNvSpPr", NS)
                is_textbox = non_visual_shape is not None and (non_visual_shape.get("txBox") or "").strip().lower() in {"1", "true", "on", "yes"}
                if is_textbox and preset_geometry is None and custom_geometry is None:
                    shape_geometry = "textbox"
                    shape_geometry_source = "p:cNvSpPr@txBox"
            if shape_geometry is None and preset_geometry is not None and preset_geometry.get("prst"):
                shape_geometry = preset_geometry.get("prst")
                shape_geometry_source = "a:prstGeom@prst"
            elif shape_geometry is None and custom_geometry is not None:
                shape_geometry = "custom"
                shape_geometry_source = "a:custGeom"
            elif shape_geometry is None and object_type == "group":
                shape_geometry = "group"
                shape_geometry_source = "p:grpSp"
            elif shape_geometry is None and object_type == "graphic_frame":
                shape_geometry = "graphicFrame"
                shape_geometry_source = "p:graphicFrame"
            named_objects.append({
                "type": object_type,
                "name": non_visual.get("name") if non_visual is not None else None,
                "id": non_visual.get("id") if non_visual is not None else None,
                "has_text": node.find(".//a:t", NS) is not None,
                "has_text_body": node.find("./p:txBody", NS) is not None,
                "has_omml": node.find(".//m:oMath", NS) is not None,
                "omml_text": "".join(math_text.text or "" for math_text in node.findall(".//m:t", NS)),
                "text": "\n".join(paragraphs),
                "box": box_dict(xfrm_box(node.find(xfrm_paths[object_type], NS))),
                "start_attached": start_connection is not None,
                "end_attached": end_connection is not None,
                "start_shape_id": start_connection.get("id") if start_connection is not None else None,
                "end_shape_id": end_connection.get("id") if end_connection is not None else None,
                "start_site_index": start_connection.get("idx") if start_connection is not None else None,
                "end_site_index": end_connection.get("idx") if end_connection is not None else None,
                "line_width_emu": int_attr(line, "w"),
                "head_arrow_type": head.get("type") if head is not None else None,
                "tail_arrow_type": tail.get("type") if tail is not None else None,
                "dash": dash.get("val") if dash is not None else "solid",
                "line_color_srgb": srgb.get("val") if srgb is not None else None,
                "line_color_scheme": scheme.get("val") if scheme is not None else None,
                "shape_geometry": shape_geometry,
                "normalized_shape_geometry": normalize_shape_geometry(shape_geometry),
                "shape_geometry_source": shape_geometry_source,
            })
    resolved_shape_boxes = {shape["name"]: shape["box"] for shape in native_shapes if shape.get("name") and shape.get("box")}
    resolved_picture_boxes = {picture["name"]: picture["box"] for picture in pictures if picture.get("name") and picture.get("box")}
    for item in named_objects:
        if item["type"] == "shape" and item["name"] in resolved_shape_boxes:
            item["box"] = resolved_shape_boxes[item["name"]]
        elif item["type"] == "picture" and item["name"] in resolved_picture_boxes:
            item["box"] = resolved_picture_boxes[item["name"]]
    connector_objects = [item for item in named_objects if item["type"] == "connector"]
    image_rels = [rel for rel in rels.values() if rel["type"].endswith(IMAGE_REL_SUFFIX)]
    used_image_rel_ids = {
        rid
        for picture in pictures
        for rid in picture["relationship_ids"]
        if rid in rels and rels[rid]["type"].endswith(IMAGE_REL_SUFFIX)
    }
    background = effective_background(zf, slide_part, root)
    image_fill_shapes = [shape for shape in native_shapes if shape["image_fill"]]
    for shape in image_fill_shapes:
        shape["relationships"] = [rels[rid] for rid in shape["relationship_ids"] if rid in rels]
        shape["unresolved_relationship_ids"] = [rid for rid in shape["relationship_ids"] if rid not in rels]
    image_fill_rel_ids = {rid for shape in image_fill_shapes for rid in shape["relationship_ids"]}
    all_blip_rel_ids = {
        rid
        for blip in root.findall(".//a:blip", NS)
        for rid in (blip.get(R_EMBED), blip.get(R_LINK))
        if rid
    }
    background_rel_ids = {rel["id"] for rel in background["image_relationships"]}
    unclassified_blip_rel_ids = sorted(all_blip_rel_ids - used_image_rel_ids - image_fill_rel_ids - background_rel_ids)
    all_used_image_rel_ids = used_image_rel_ids | image_fill_rel_ids | set(unclassified_blip_rel_ids) | background_rel_ids
    return {
        "index": index,
        "part": slide_part,
        "hidden": bool_is_false(root.get("show")),
        "shape_counts": {
            "auto_shapes": len(root.findall(".//p:sp", NS)),
            "connectors": len(root.findall(".//p:cxnSp", NS)),
            "connectors_attached_both_ends": sum(1 for item in connector_objects if item["start_attached"] and item["end_attached"]),
            "groups": len(root.findall(".//p:grpSp", NS)),
            "pictures": len(pictures),
            "graphic_frames": len(root.findall(".//p:graphicFrame", NS)),
            "text_shapes": sum(1 for shape in root.findall(".//p:sp", NS) if shape.find(".//a:t", NS) is not None),
            "text_runs": len(text_values),
            "nonempty_text_runs": sum(1 for value in text_values if value.strip()),
            "omml_objects": len(root.findall(".//m:oMath", NS)),
            "omml_paragraphs": len(root.findall(".//m:oMathPara", NS)),
        },
        "text_sample": [value for value in text_values if value.strip()][:12],
        "pictures": pictures,
        "native_shapes": native_shapes,
        "image_fill_shapes": image_fill_shapes,
        "unclassified_blip_relationship_ids": unclassified_blip_rel_ids,
        "named_objects": named_objects,
        "image_relationships": image_rels,
        "used_image_relationship_ids": sorted(all_used_image_rel_ids),
        "background": background,
        "warnings": rel_warnings + geometry_warnings + native_shape_warnings + background["warnings"],
    }


def presentation_slides(zf: zipfile.ZipFile) -> tuple[list[str], int | None, int | None, list[str]]:
    root = parse_xml(zf, "ppt/presentation.xml")
    rels, warnings = read_relationships(zf, "ppt/presentation.xml")
    parts: list[str] = []
    for slide_id in root.findall("./p:sldIdLst/p:sldId", NS):
        rid = slide_id.get(R_ID)
        relation = rels.get(rid or "")
        if not relation or not relation["type"].endswith(SLIDE_REL_SUFFIX) or not relation["resolved_target"]:
            warnings.append(f"Cannot resolve presentation slide relationship {rid!r}")
            continue
        parts.append(relation["resolved_target"])
    size = root.find("./p:sldSz", NS)
    return parts, int_attr(size, "cx"), int_attr(size, "cy"), warnings


def manifest_source(manifest: dict[str, Any]) -> dict[str, Any]:
    value = manifest.get("source")
    return value if isinstance(value, dict) else {}


def valid_sha(value: Any) -> str | None:
    return value.lower() if isinstance(value, str) and SHA256_RE.fullmatch(value) else None


def normalize_ooxml_kind(value: Any) -> str | None:
    if not isinstance(value, str) or not value.strip():
        return None
    compact = re.sub(r"[\s_]+", "-", value.strip().lower())
    if "cxnsp" in compact and "attach" in compact:
        return "attached_connector"
    aliases = {
        "p:sp": "shape",
        "shape": "shape",
        "auto-shape": "shape",
        "autoshape": "shape",
        "drawingml-shape": "shape",
        "text": "text",
        "text-box": "text",
        "textbox": "text",
        "p:sp-text": "text",
        "drawingml-text": "text",
        "p:cxnsp": "connector",
        "connector": "connector",
        "drawingml-connector": "connector",
        "p:cxnsp-attached": "attached_connector",
        "attached-connector": "attached_connector",
        "drawingml-attached-connector": "attached_connector",
        "p:grpsp": "group",
        "group": "group",
        "drawingml-group": "group",
        "p:pic": "picture",
        "picture": "picture",
        "image": "picture",
        "drawingml-picture": "picture",
        "p:graphicframe": "graphic_frame",
        "graphic-frame": "graphic_frame",
        "graphicframe": "graphic_frame",
        "m:omath": "omml",
        "m:omathpara": "omml",
        "omml": "omml",
        "office-math": "omml",
    }
    return aliases.get(compact)


def normalize_shape_geometry(value: Any) -> str | None:
    """Normalize a scene-plan geometry token without guessing its meaning.

    DrawingML preset names are compared case-insensitively with separators
    removed.  A small alias set covers common human spellings while preserving
    every other non-empty preset token for an exact normalized comparison.
    """
    if not isinstance(value, str) or not value.strip():
        return None
    compact = re.sub(r"[^a-z0-9]+", "", value.strip().casefold())
    aliases = {
        "rectangle": "rect",
        "roundedrectangle": "roundrect",
        "roundedrect": "roundrect",
        "roundrectangle": "roundrect",
        "oval": "ellipse",
        "circle": "ellipse",
        "text": "textbox",
        "textframe": "textbox",
        "textshape": "textbox",
        "textplaceholder": "textbox",
        "customgeometry": "custom",
        "freeform": "custom",
        "graphicframe": "graphicframe",
    }
    return aliases.get(compact, compact) if compact else None


def plan_expected_counts(plan: dict[str, Any]) -> dict[str, Any]:
    objects = plan.get("objects") if isinstance(plan.get("objects"), list) else []
    connections = plan.get("connections") if isinstance(plan.get("connections"), list) else []
    native_objects = 0
    raster_objects = 0
    text_objects = 0
    formula_objects = 0
    unsupported_objects = 0
    approximation_objects = 0
    object_ids: list[str] = []
    raster_object_ids: list[str] = []
    native_object_ids: list[str] = []
    mode_counts: dict[str, int] = {}
    ooxml_expectations: list[dict[str, Any]] = []
    connection_ooxml_expectations: list[dict[str, Any]] = []
    connection_mode_counts: dict[str, int] = {}
    raster_connections = 0
    raster_connection_ids: list[str] = []
    unsupported_connections = 0
    for item in objects:
        if not isinstance(item, dict):
            continue
        object_id = item.get("id")
        if isinstance(object_id, str) and object_id:
            object_ids.append(object_id)
        reconstruction = item.get("reconstruction") if isinstance(item.get("reconstruction"), dict) else {}
        implementation = item.get("implementation") if isinstance(item.get("implementation"), dict) else {}
        mode = str(implementation.get("classification") or reconstruction.get("mode") or "").lower()
        raw_ooxml_kind = implementation.get("ppt_object_kind") or reconstruction.get("expected_ooxml_kind")
        normalized_ooxml_kind = normalize_ooxml_kind(raw_ooxml_kind)
        ooxml_expectations.append({
            "object_id": object_id if isinstance(object_id, str) else None,
            "mode": mode,
            "expected_ooxml_kind": raw_ooxml_kind,
            "normalized_ooxml_kind": normalized_ooxml_kind,
        })
        mode_counts[mode or "unrecognized"] = mode_counts.get(mode or "unrecognized", 0) + 1
        if mode in {"native-exact", "native-approximation", "native", "vector", "hybrid"}:
            native_objects += 1
            if isinstance(object_id, str) and object_id:
                native_object_ids.append(object_id)
        if mode in {"isolated-raster", "raster"}:
            raster_objects += 1
            if isinstance(object_id, str) and object_id:
                raster_object_ids.append(object_id)
        if mode == "native-approximation":
            approximation_objects += 1
        if mode in {"unsupported", "omit"}:
            unsupported_objects += 1
        item_type = str(item.get("type", "")).lower()
        content = item.get("content") if isinstance(item.get("content"), dict) else {}
        if isinstance(content.get("text"), str) and content["text"] != "":
            text_objects += 1
        elif any(token in item_type for token in ("text", "label", "caption", "title")):
            text_objects += 1
        if normalized_ooxml_kind == "omml":
            formula_objects += 1
    for connection in connections:
        if not isinstance(connection, dict):
            continue
        connection_id = connection.get("id")
        reconstruction = connection.get("reconstruction") if isinstance(connection.get("reconstruction"), dict) else {}
        implementation = connection.get("implementation") if isinstance(connection.get("implementation"), dict) else {}
        mode = str(implementation.get("classification") or reconstruction.get("mode") or "").lower()
        raw_ooxml_kind = implementation.get("ppt_object_kind") or reconstruction.get("expected_ooxml_kind")
        connection_mode_counts[mode or "unrecognized"] = connection_mode_counts.get(mode or "unrecognized", 0) + 1
        if mode in {"isolated-raster", "raster"}:
            raster_connections += 1
            if isinstance(connection_id, str) and connection_id:
                raster_connection_ids.append(connection_id)
        if mode in {"unsupported", "omit"}:
            unsupported_connections += 1
        connection_ooxml_expectations.append({
            "connection_id": connection_id if isinstance(connection_id, str) else None,
            "mode": mode,
            "expected_ooxml_kind": raw_ooxml_kind,
            "normalized_ooxml_kind": normalize_ooxml_kind(raw_ooxml_kind),
        })
    return {
        "objects": len(objects),
        "native_objects": native_objects,
        "raster_objects": raster_objects,
        "text_objects": text_objects,
        "formula_objects": formula_objects,
        "connections": len(connections),
        "unsupported_objects": unsupported_objects,
        "unsupported_connections": unsupported_connections,
        "native_approximation_objects": approximation_objects,
        "object_ids": object_ids,
        "native_object_ids": native_object_ids,
        "raster_object_ids": raster_object_ids,
        "mode_counts": mode_counts,
        "connection_mode_counts": connection_mode_counts,
        "ooxml_expectations": ooxml_expectations,
        "connection_ooxml_expectations": connection_ooxml_expectations,
        "raster_connections": raster_connections,
        "raster_connection_ids": raster_connection_ids,
        "raster_entity_ids": raster_object_ids + raster_connection_ids,
        "raster_entities": raster_objects + raster_connections,
    }


def issue_from_check(check: dict[str, Any]) -> dict[str, Any]:
    return {"id": check["id"], "message": check["message"], "severity": check["severity"]}


class ReportBuilder:
    def __init__(self) -> None:
        self.checks: list[dict[str, Any]] = []

    def add(self, check_id: str, status: str, severity: str, message: str, **evidence: Any) -> None:
        normalized_status = status.upper()
        normalized_severity = severity.upper()
        if normalized_status not in {"PASS", "WARN", "FAIL"}:
            raise ValueError(f"Invalid check status: {status}")
        if normalized_severity not in {"HARD", "SOFT", "INFO"}:
            raise ValueError(f"Invalid check severity: {severity}")
        check: dict[str, Any] = {
            "id": check_id,
            "status": normalized_status,
            "severity": normalized_severity,
            "message": message,
        }
        if evidence:
            check["evidence"] = evidence
        self.checks.append(check)

    def finish(self, **payload: Any) -> dict[str, Any]:
        hard = [issue_from_check(c) for c in self.checks if c["severity"] == "HARD" and c["status"] == "FAIL"]
        warnings = [issue_from_check(c) for c in self.checks if c["status"] == "WARN"]
        status = "FAIL" if hard else "WARN" if warnings else "PASS"
        return {
            "schema_version": "1.0",
            "kind": "sci-diagram-pptx-pptx-audit",
            "tool": "sci-diagram-pptx/audit_pptx.py",
            "tool_version": TOOL_VERSION,
            "generated_at": utc_now(),
            "status": status,
            "hard_failures": hard,
            "warnings": warnings,
            "checks": self.checks,
            **payload,
        }


def audit_pptx(
    pptx_path: Path,
    source_manifest_path: Path,
    scene_plan_path: Path,
    full_bleed_tolerance: float = 0.01,
    aspect_tolerance: float = 0.005,
) -> dict[str, Any]:
    report = ReportBuilder()
    pptx = pptx_path.expanduser().resolve()
    manifest = load_json(source_manifest_path, "source manifest")
    plan = load_json(scene_plan_path, "scene plan")
    source = manifest_source(manifest)
    plan_source = plan.get("source") if isinstance(plan.get("source"), dict) else {}
    expected = plan_expected_counts(plan)

    manifest_errors = manifest.get("errors") if isinstance(manifest.get("errors"), list) else []
    manifest_hard = manifest.get("hard_failures") if isinstance(manifest.get("hard_failures"), list) else []
    manifest_warnings = manifest.get("warnings") if isinstance(manifest.get("warnings"), list) else []
    if manifest.get("ok") is False or manifest.get("status") == "FAIL" or manifest_errors or manifest_hard:
        report.add("inputs.source_manifest", "FAIL", "HARD", "Source preflight report contains a hard failure", error_count=len(manifest_errors), hard_failure_count=len(manifest_hard))
    elif not source:
        report.add("inputs.source_manifest", "WARN", "HARD", "Source manifest schema is unrecognized; source identity cannot be proven")
    elif manifest_warnings or str(manifest.get("status", "")).upper() == "WARN":
        report.add("inputs.source_manifest", "WARN", "HARD", "Source manifest is readable but contains unresolved warnings", warning_count=len(manifest_warnings))
    else:
        report.add("inputs.source_manifest", "PASS", "HARD", "Source manifest is readable and reports no hard failure")
    report.add(
        "inputs.source_manifest_kind",
        "PASS" if manifest.get("kind") == "sci-diagram-pptx-source-manifest" else "WARN",
        "HARD",
        "Source manifest kind is recognized" if manifest.get("kind") == "sci-diagram-pptx-source-manifest" else "Source manifest kind is missing or unexpected; it was not treated as fully trusted",
        expected="sci-diagram-pptx-source-manifest",
        actual=manifest.get("kind"),
    )

    required_plan_keys = {"schema_version", "source", "canvas", "objects", "connections", "approvals"}
    missing_plan_keys = sorted(required_plan_keys - set(plan))
    if missing_plan_keys:
        report.add("inputs.scene_plan", "FAIL", "HARD", "Scene plan is missing required top-level fields", missing=missing_plan_keys)
    elif not isinstance(plan.get("objects"), list) or not isinstance(plan.get("connections"), list):
        report.add("inputs.scene_plan", "FAIL", "HARD", "Scene plan objects and connections must be arrays")
    else:
        report.add("inputs.scene_plan", "PASS", "HARD", "Scene plan has the expected top-level structure")
    report.add(
        "inputs.scene_plan_kind",
        "PASS" if plan.get("kind") == "sci-diagram-pptx-scene-plan" else "WARN",
        "HARD",
        "Scene plan kind is recognized" if plan.get("kind") == "sci-diagram-pptx-scene-plan" else "Scene plan kind is missing or unexpected; it was not treated as fully trusted",
        expected="sci-diagram-pptx-scene-plan",
        actual=plan.get("kind"),
    )

    source_sha = valid_sha(source.get("sha256"))
    plan_sha = valid_sha(plan_source.get("sha256"))
    if source_sha and plan_sha:
        report.add(
            "inputs.source_identity",
            "PASS" if source_sha == plan_sha else "FAIL",
            "HARD",
            "Scene plan is bound to the locked source" if source_sha == plan_sha else "Scene plan source SHA-256 differs from the source manifest",
            source_manifest_sha256=source_sha,
            scene_plan_sha256=plan_sha,
        )
    else:
        report.add("inputs.source_identity", "WARN", "HARD", "A valid SHA-256 is missing from the source manifest or scene plan; source identity is unproven")

    if not pptx.exists() or not pptx.is_file():
        raise AuditError(f"PPTX does not exist or is not a regular file: {pptx}")
    if not zipfile.is_zipfile(pptx):
        raise AuditError(f"File is not a readable ZIP-based PPTX: {pptx}")

    try:
        with zipfile.ZipFile(pptx) as zf:
            bad_member = zf.testzip()
            if bad_member:
                raise AuditError(f"PPTX ZIP integrity check failed at {bad_member}")
            names = set(zf.namelist())
            report.add("package.readable", "PASS", "HARD", "PPTX is a readable ZIP package with valid member CRCs", member_count=len(names))
            slide_parts, width, height, presentation_warnings = presentation_slides(zf)
            if width is None or height is None or width <= 0 or height <= 0:
                report.add("presentation.slide_size", "FAIL", "HARD", "Presentation slide size is absent or invalid")
                width = height = 0
            else:
                report.add("presentation.slide_size", "PASS", "HARD", "Presentation slide size is explicit", width_emu=width, height_emu=height, width_inches=width / EMU_PER_INCH, height_inches=height / EMU_PER_INCH)

            if len(slide_parts) == 2:
                report.add("presentation.slide_count", "PASS", "HARD", "Presentation contains the default two-slide deliverable", slide_count=2)
            else:
                report.add("presentation.slide_count", "FAIL", "HARD", "Presentation must contain exactly two slides for the scientific-diagram reconstruction contract", slide_count=len(slide_parts), expected=2)

            orphan_slide_parts = sorted(name for name in names if re.fullmatch(r"ppt/slides/slide\d+\.xml", name) and name not in set(slide_parts))
            report.add(
                "package.orphan_slide_parts",
                "PASS" if not orphan_slide_parts else "FAIL",
                "HARD",
                "No orphaned slide XML parts were found" if not orphan_slide_parts else "Slide XML parts exist outside the presentation slide list",
                orphan_parts=orphan_slide_parts,
            )

            slides: list[dict[str, Any]] = []
            for index, part in enumerate(slide_parts, start=1):
                if part not in names:
                    report.add(f"slide{index}.part_exists", "FAIL", "HARD", "Slide relationship target is missing", part=part)
                    continue
                try:
                    slides.append(slide_analysis(zf, part, index, width, height, full_bleed_tolerance))
                except AuditError as exc:
                    report.add(f"slide{index}.xml", "FAIL", "HARD", str(exc))

            hidden = [slide["index"] for slide in slides if slide["hidden"]]
            report.add(
                "presentation.visible_slides",
                "PASS" if not hidden and len(slides) == len(slide_parts) else "WARN" if hidden else "FAIL",
                "HARD",
                "All listed slides are visible" if not hidden and len(slides) == len(slide_parts) else "One or more slides are hidden; visibility may violate the active contract" if hidden else "Not every listed slide could be inspected",
                hidden_slides=hidden,
            )

            external, rel_parse_warnings = package_external_relationships(zf)
            report.add(
                "package.no_external_relationships",
                "PASS" if not external and not rel_parse_warnings else "WARN" if rel_parse_warnings and not external else "FAIL",
                "HARD",
                "No external relationships were found" if not external and not rel_parse_warnings else "Relationship parsing was incomplete; absence of external links is unproven" if not external else "External relationships make the deck non-self-contained",
                external_relationships=external,
                parse_warnings=rel_parse_warnings,
            )

            content_types_text = zf.read("[Content_Types].xml").decode("utf-8", "replace").lower() if "[Content_Types].xml" in names else ""
            macro_parts = sorted(name for name in names if "vbaproject" in name.lower() or name.lower().endswith(".vba"))
            macro_enabled = "macroenabled" in content_types_text or "application/vnd.ms-office.vbaproject" in content_types_text
            report.add(
                "package.no_macros",
                "PASS" if not macro_parts and not macro_enabled else "FAIL",
                "HARD",
                "No VBA macro payload or macro-enabled content type was found" if not macro_parts and not macro_enabled else "Macro content is present",
                macro_parts=macro_parts,
                macro_enabled_content_type=macro_enabled,
            )

            embedding_parts = sorted(name for name in names if name.startswith("ppt/embeddings/") and not name.endswith("/"))
            embedded_relationships, embedded_rel_warnings = package_relationship_type_hits(zf, ("/oleObject", "/package"))
            ole_xml = []
            for slide in slides:
                try:
                    root = parse_xml(zf, slide["part"])
                    if root.findall(".//p:oleObj", NS):
                        ole_xml.append(slide["part"])
                except AuditError:
                    pass
            report.add(
                "package.no_ole_or_embedded_packages",
                "PASS" if not embedding_parts and not ole_xml and not embedded_relationships and not embedded_rel_warnings else "WARN" if embedded_rel_warnings and not embedding_parts and not ole_xml and not embedded_relationships else "FAIL",
                "HARD",
                "No OLE objects or embedded packages were found" if not embedding_parts and not ole_xml and not embedded_relationships and not embedded_rel_warnings else "Relationship parsing was incomplete; absence of OLE/package links is unproven" if embedded_rel_warnings and not embedding_parts and not ole_xml and not embedded_relationships else "OLE objects or embedded package parts are present",
                embedding_parts=embedding_parts,
                slides_with_ole=ole_xml,
                embedded_relationships=embedded_relationships,
                parse_warnings=embedded_rel_warnings,
            )

            media = sorted(name for name in names if name.startswith("ppt/media/") and not name.endswith("/"))
            vector_media = [name for name in media if Path(name).suffix.lower() in VECTOR_EXTENSIONS]
            report.add(
                "package.no_svg_emf_wmf_shortcuts",
                "PASS" if not vector_media else "FAIL",
                "HARD",
                "No SVG, EMF, or WMF media parts were found" if not vector_media else "Vector media shortcuts are present instead of native PowerPoint objects",
                vector_media=vector_media,
            )

            if slides:
                slide1 = slides[0]
                image_fill_shapes = slide1["image_fill_shapes"]
                full_bleed = [picture for picture in slide1["pictures"] if picture["full_bleed"]]
                full_bleed_image_fills = [
                    shape for shape in image_fill_shapes
                    if not shape["geometry_uncertain"] and box_covers_slide(shape["box"], width, height, full_bleed_tolerance)
                ]
                page_covering = [
                    item for item in [*slide1["pictures"], *image_fill_shapes]
                    if isinstance(item.get("slide_coverage_fraction"), (int, float))
                    and item["slide_coverage_fraction"] >= 0.80
                ]
                known_picture_boxes = [
                    item["box"] for item in [*slide1["pictures"], *image_fill_shapes]
                    if item["box"] is not None and not item["geometry_uncertain"]
                ]
                picture_union_coverage = rectangle_union_fraction(known_picture_boxes, width, height)
                uncertain_pictures = [item for item in [*slide1["pictures"], *image_fill_shapes] if item["geometry_uncertain"]]
                hidden_pictures = [
                    item for item in [*slide1["pictures"], *image_fill_shapes]
                    if item["hidden"] or item.get("fully_transparent") or item.get("fully_transparent_fill")
                ]
                off_canvas_pictures = [item for item in [*slide1["pictures"], *image_fill_shapes] if item["off_canvas"] is True]
                partially_off_canvas_pictures = [item for item in [*slide1["pictures"], *image_fill_shapes] if item["partially_off_canvas"] is True]
                background = slide1["background"]
                background_image = bool(background["has_blip_fill"] or background["image_relationships"])
                # A theme/master bgRef alone is not evidence of an image.  Treat
                # only unresolved image relationships or parse warnings as
                # uncertainty; direct blip fills are handled as failures above.
                background_uncertain = bool(background["unresolved_relationship_ids"] or background["warnings"])
                if full_bleed or full_bleed_image_fills or page_covering or picture_union_coverage >= 0.90 or background_image:
                    report.add(
                        "slide1.no_full_slide_image_or_image_background",
                        "FAIL",
                        "HARD",
                        "Slide 1 contains a full/near-page picture, image tiling, or image background",
                        full_bleed_picture_count=len(full_bleed),
                        full_bleed_image_fill_shape_count=len(full_bleed_image_fills),
                        page_covering_picture_count=len(page_covering),
                        picture_union_coverage=picture_union_coverage,
                        background=background,
                    )
                elif uncertain_pictures or background_uncertain:
                    report.add("slide1.no_full_slide_image_or_image_background", "WARN", "HARD", "Picture/background parsing is incomplete; absence of a page-covering raster cannot be proven", uncertain_picture_count=len(uncertain_pictures), picture_union_coverage=picture_union_coverage, background=background)
                else:
                    report.add("slide1.no_full_slide_image_or_image_background", "PASS", "HARD", "Slide 1 has no detected page-covering picture, image tiling, or image background", picture_union_coverage=picture_union_coverage)

                page_sized_shapes = [
                    shape for shape in slide1["native_shapes"]
                    if not shape["geometry_uncertain"] and box_covers_slide(shape["box"], width, height, full_bleed_tolerance)
                ]
                uncertain_native_shapes = [shape for shape in slide1["native_shapes"] if shape["geometry_uncertain"]]
                if page_sized_shapes:
                    report.add(
                        "slide1.no_page_sized_background_shapes",
                        "FAIL",
                        "HARD",
                        "Slide 1 contains page-sized native shapes, including possible transparent placeholders or artificial backgrounds",
                        shapes=page_sized_shapes,
                    )
                elif uncertain_native_shapes:
                    report.add("slide1.no_page_sized_background_shapes", "WARN", "HARD", "Some native-shape geometry is unresolved; absence of a page-sized placeholder is unproven", uncertain_shape_count=len(uncertain_native_shapes))
                else:
                    report.add("slide1.no_page_sized_background_shapes", "PASS", "HARD", "Slide 1 contains no detected page-sized native background shape or placeholder")

                slide1_image_hashes: list[dict[str, Any]] = []
                unresolved_slide1_images: list[dict[str, Any]] = []
                for relation in slide1["image_relationships"]:
                    target = relation["resolved_target"]
                    if relation["external"] or not target or target not in names:
                        unresolved_slide1_images.append(relation)
                    else:
                        slide1_image_hashes.append({"relationship_id": relation["id"], "target": target, "sha256": sha256_bytes(zf.read(target))})
                source_image_occurrences = [item for item in slide1_image_hashes if source_sha and item["sha256"] == source_sha]
                if source_image_occurrences or hidden_pictures or off_canvas_pictures:
                    report.add(
                        "slide1.no_hidden_or_offcanvas_source_images",
                        "FAIL",
                        "HARD",
                        "Slide 1 contains the locked source image, a hidden/transparent picture, or an off-canvas picture",
                        locked_source_occurrences=source_image_occurrences,
                        hidden_or_transparent_picture_count=len(hidden_pictures),
                        off_canvas_picture_count=len(off_canvas_pictures),
                    )
                elif unresolved_slide1_images or uncertain_pictures:
                    report.add("slide1.no_hidden_or_offcanvas_source_images", "WARN", "HARD", "Some Slide 1 image relationships or geometries are unresolved; hidden-source absence is unproven", unresolved_relationships=unresolved_slide1_images, uncertain_picture_count=len(uncertain_pictures))
                elif partially_off_canvas_pictures:
                    report.add("slide1.no_hidden_or_offcanvas_source_images", "WARN", "HARD", "A Slide 1 picture extends outside the canvas; inspect whether this is an approved local crop", partially_off_canvas_picture_count=len(partially_off_canvas_pictures))
                else:
                    report.add("slide1.no_hidden_or_offcanvas_source_images", "PASS", "HARD", "No locked-source, hidden, transparent, or off-canvas Slide 1 picture was detected")

                unclassified_image_bearer_count = len(slide1["unclassified_blip_relationship_ids"])
                actual_raster_structure_count = slide1["shape_counts"]["pictures"] + len(image_fill_shapes) + unclassified_image_bearer_count
                if actual_raster_structure_count != expected["raster_entities"]:
                    report.add(
                        "slide1.raster_objects_against_plan",
                        "FAIL",
                        "HARD",
                        "Slide 1 picture count differs from isolated-raster objects declared by the scene plan",
                        planned_isolated_raster_objects=expected["raster_objects"],
                        planned_isolated_raster_connections=expected["raster_connections"],
                        actual_picture_objects=slide1["shape_counts"]["pictures"],
                        actual_image_fill_shapes=len(image_fill_shapes),
                        unclassified_image_bearers=unclassified_image_bearer_count,
                        planned_ids=expected["raster_object_ids"],
                    )
                else:
                    report.add("slide1.raster_objects_against_plan", "PASS", "HARD", "Slide 1 image-bearing object count exactly matches planned isolated-raster exceptions", planned=expected["raster_entities"], actual=actual_raster_structure_count)

                raster_entity_names = set(expected["raster_entity_ids"])
                image_bearer_names = [item.get("name") for item in [*slide1["pictures"], *image_fill_shapes]]
                unexpected_image_bearers = [name for name in image_bearer_names if not isinstance(name, str) or name not in raster_entity_names]
                missing_named_rasters = sorted(raster_entity_names - {name for name in image_bearer_names if isinstance(name, str)})
                if unexpected_image_bearers or missing_named_rasters or slide1["unclassified_blip_relationship_ids"]:
                    report.add(
                        "slide1.image_bearers_match_isolated_raster_plan",
                        "FAIL",
                        "HARD",
                        "Every Slide 1 image-bearing picture/shape must map by cNvPr name to one approved isolated-raster entity",
                        unexpected_image_bearer_names=unexpected_image_bearers,
                        missing_named_raster_entities=missing_named_rasters,
                        unclassified_blip_relationship_ids=slide1["unclassified_blip_relationship_ids"],
                    )
                else:
                    report.add("slide1.image_bearers_match_isolated_raster_plan", "PASS", "HARD", "Every Slide 1 image-bearing object maps by name to an isolated-raster plan entity", entity_ids=sorted(raster_entity_names))

                if expected["unsupported_objects"] or expected["unsupported_connections"]:
                    report.add("plan.no_unsupported_objects", "FAIL", "HARD", "Scene plan contains unsupported objects or connections and cannot authorize a deliverable", object_count=expected["unsupported_objects"], connection_count=expected["unsupported_connections"])
                else:
                    report.add("plan.no_unsupported_objects", "PASS", "HARD", "Scene plan contains no unsupported object or connection classifications")

                editable_count = sum(slide1["shape_counts"][key] for key in ("auto_shapes", "connectors", "graphic_frames"))
                report.add(
                    "slide1.has_native_objects",
                    "PASS" if editable_count > 0 else "FAIL",
                    "HARD",
                    "Slide 1 contains native object structures" if editable_count > 0 else "Slide 1 contains no detected native shapes, connectors, or graphic frames",
                    editable_structure_count=editable_count,
                )

            if slides:
                actual = slides[0]["shape_counts"]
                accepted_modes = {"native-exact", "native-approximation", "isolated-raster", "unsupported", "native", "vector", "hybrid", "raster", "omit"}
                unknown_modes = {
                    "objects": {key: value for key, value in expected["mode_counts"].items() if key not in accepted_modes},
                    "connections": {key: value for key, value in expected["connection_mode_counts"].items() if key not in accepted_modes},
                }
                has_unknown_modes = any(unknown_modes.values())
                if has_unknown_modes:
                    report.add("plan.reconstruction_modes", "WARN", "HARD", "Unknown implementation classifications were not treated as native or raster", unknown_mode_counts=unknown_modes)
                else:
                    report.add("plan.reconstruction_modes", "PASS", "HARD", "All object and connection implementation classifications are recognized", object_mode_counts=expected["mode_counts"], connection_mode_counts=expected["connection_mode_counts"])

                actual_kind_counts = {
                    "shape": actual["auto_shapes"],
                    "text": actual["text_shapes"],
                    "connector": actual["connectors"],
                    "attached_connector": actual["connectors_attached_both_ends"],
                    "group": actual["groups"],
                    "picture": actual["pictures"],
                    "graphic_frame": actual["graphic_frames"],
                    "omml": actual["omml_objects"],
                }
                expected_kind_counts: dict[str, int] = {}
                unknown_kinds: list[dict[str, Any]] = []
                all_ooxml_expectations = expected["ooxml_expectations"] + expected["connection_ooxml_expectations"]
                for expectation in all_ooxml_expectations:
                    kind = expectation["normalized_ooxml_kind"]
                    if kind is None:
                        unknown_kinds.append(expectation)
                    else:
                        expected_kind_counts[kind] = expected_kind_counts.get(kind, 0) + 1
                deficient_kinds = {
                    kind: {"expected": count, "actual": actual_kind_counts.get(kind, 0)}
                    for kind, count in expected_kind_counts.items()
                    if actual_kind_counts.get(kind, 0) < count
                }
                if deficient_kinds:
                    report.add("slide1.expected_ooxml_kind_counts", "FAIL", "HARD", "Slide 1 contains fewer OOXML objects than required by expected_ooxml_kind", deficient_kinds=deficient_kinds, actual_kind_counts=actual_kind_counts)
                elif unknown_kinds:
                    report.add("slide1.expected_ooxml_kind_counts", "WARN", "HARD", "One or more expected_ooxml_kind values are unknown; those objects were not counted as passing", unknown_expectations=unknown_kinds, expected_kind_counts=expected_kind_counts, actual_kind_counts=actual_kind_counts)
                else:
                    report.add("slide1.expected_ooxml_kind_counts", "PASS", "HARD", "Slide 1 has sufficient OOXML structures for every declared expected_ooxml_kind", expected_kind_counts=expected_kind_counts, actual_kind_counts=actual_kind_counts)

                objects_by_name = {
                    item["name"]: item
                    for item in slides[0]["named_objects"]
                    if isinstance(item.get("name"), str) and item["name"]
                }
                unpaired: list[dict[str, Any]] = []
                type_mismatches: list[dict[str, Any]] = []
                for expectation in all_ooxml_expectations:
                    object_id = expectation.get("object_id") or expectation.get("connection_id")
                    kind = expectation["normalized_ooxml_kind"]
                    if not object_id or kind is None:
                        continue
                    actual_object = objects_by_name.get(object_id)
                    if actual_object is None:
                        unpaired.append(expectation)
                        continue
                    kind_matches = (
                        actual_object["type"] == kind
                        or (kind == "text" and actual_object["type"] == "shape" and actual_object["has_text"])
                        or (kind == "omml" and actual_object["has_omml"])
                        or (kind == "attached_connector" and actual_object["type"] == "connector" and actual_object["start_attached"] and actual_object["end_attached"])
                    )
                    if not kind_matches:
                        type_mismatches.append({"expectation": expectation, "actual": actual_object})
                if type_mismatches:
                    report.add("slide1.named_object_ooxml_kinds", "FAIL", "HARD", "Named Slide 1 objects disagree with their declared expected_ooxml_kind", mismatches=type_mismatches)
                elif unpaired:
                    report.add("slide1.named_object_ooxml_kinds", "WARN", "HARD", "Some plan object IDs could not be paired to cNvPr names; per-object OOXML kind is unproven", unpaired_expectations=unpaired)
                elif unknown_kinds:
                    report.add("slide1.named_object_ooxml_kinds", "WARN", "HARD", "Unknown expected_ooxml_kind values prevent a complete per-object audit", unknown_expectations=unknown_kinds)
                else:
                    report.add("slide1.named_object_ooxml_kinds", "PASS", "HARD", "Every plan object ID was paired to a named OOXML object of the expected kind")

                object_semantic_failures: list[dict[str, Any]] = []
                object_semantic_uncertain: list[dict[str, Any]] = []
                object_geometry_failures: list[dict[str, Any]] = []
                object_geometry_uncertain: list[dict[str, Any]] = []
                manual_math_reviews: list[dict[str, Any]] = []
                bbox_tolerance = max(0.01, full_bleed_tolerance)
                for planned_object in plan.get("objects", []):
                    if not isinstance(planned_object, dict) or not isinstance(planned_object.get("id"), str):
                        continue
                    object_id = planned_object["id"]
                    actual_object = objects_by_name.get(object_id)
                    if actual_object is None:
                        object_semantic_uncertain.append({"object_id": object_id, "reason": "cNvPr name not found"})
                        continue
                    content = planned_object.get("content") if isinstance(planned_object.get("content"), dict) else {}
                    if isinstance(content.get("text"), str) and actual_object.get("text", "") != content["text"]:
                        object_semantic_failures.append({"object_id": object_id, "field": "content.text", "expected": content["text"], "actual": actual_object.get("text", "")})
                    object_reconstruction = planned_object.get("reconstruction") if isinstance(planned_object.get("reconstruction"), dict) else {}
                    object_implementation = planned_object.get("implementation") if isinstance(planned_object.get("implementation"), dict) else {}
                    object_expected_kind = normalize_ooxml_kind(object_implementation.get("ppt_object_kind") or object_reconstruction.get("expected_ooxml_kind"))
                    style = planned_object.get("style") if isinstance(planned_object.get("style"), dict) else {}
                    expected_geometry_raw = style.get("shape_geometry")
                    expected_geometry = normalize_shape_geometry(expected_geometry_raw)
                    actual_geometry = actual_object.get("normalized_shape_geometry")
                    geometry_evidence = {
                        "object_id": object_id,
                        "expected": expected_geometry_raw,
                        "expected_normalized": expected_geometry,
                        "actual": actual_object.get("shape_geometry"),
                        "actual_normalized": actual_geometry,
                        "source": actual_object.get("shape_geometry_source"),
                    }
                    if expected_geometry is None:
                        object_geometry_uncertain.append({**geometry_evidence, "reason": "plan style.shape_geometry is missing or empty"})
                    elif actual_geometry is None:
                        object_geometry_uncertain.append({**geometry_evidence, "reason": "OOXML shape geometry is absent or cannot be resolved"})
                    elif expected_geometry == "custom" or actual_geometry == "custom":
                        object_geometry_uncertain.append({**geometry_evidence, "reason": "custom/freeform geometry cannot be proven equivalent from a preset token"})
                    elif (
                        expected_geometry == "textbox"
                        and actual_geometry == "rect"
                        and actual_object.get("type") == "shape"
                        and actual_object.get("has_text_body") is True
                    ):
                        # Some authoring libraries lower a text box to a rect
                        # preset while retaining the editable p:txBody.  The
                        # direct text body distinguishes that supported lowering
                        # from a non-text rectangle as far as OOXML permits.
                        pass
                    elif expected_geometry != actual_geometry:
                        object_geometry_failures.append(geometry_evidence)
                    if object_expected_kind == "omml" and isinstance(content.get("math_source"), str) and content["math_source"]:
                        expected_math = "".join(unicodedata.normalize("NFKC", content["math_source"]).split())
                        actual_math = "".join(unicodedata.normalize("NFKC", str(actual_object.get("omml_text") or "")).split())
                        manual_math_reviews.append({"object_id": object_id, "math_source": content["math_source"], "omml_text": actual_object.get("omml_text", ""), "normalized_exact_match": bool(actual_math) and actual_math == expected_math})
                    planned_box = planned_object.get("bbox") or planned_object.get("slide_bbox")
                    actual_box = actual_object.get("box")
                    if isinstance(planned_box, dict) and actual_box and width > 0 and height > 0:
                        actual_normalized = {
                            "x": actual_box["x_emu"] / width,
                            "y": actual_box["y_emu"] / height,
                            "width": actual_box["width_emu"] / width,
                            "height": actual_box["height_emu"] / height,
                        }
                        errors = {
                            key: abs(float(planned_box[key]) - actual_normalized[key])
                            for key in ("x", "y", "width", "height")
                            if isinstance(planned_box.get(key), (int, float))
                        }
                        if len(errors) != 4:
                            object_semantic_uncertain.append({"object_id": object_id, "reason": "planned bbox is incomplete"})
                        elif max(errors.values()) > bbox_tolerance:
                            object_semantic_failures.append({"object_id": object_id, "field": "bbox", "expected": planned_box, "actual": actual_normalized, "absolute_errors": errors, "tolerance": bbox_tolerance})
                    elif isinstance(planned_box, dict):
                        object_semantic_uncertain.append({"object_id": object_id, "reason": "actual OOXML transform is unresolved"})
                if object_semantic_failures:
                    report.add("slide1.named_object_content_and_bbox", "FAIL", "HARD", "Named object text or slide bbox differs from the scene plan", failures=object_semantic_failures)
                elif object_semantic_uncertain:
                    report.add("slide1.named_object_content_and_bbox", "WARN", "HARD", "Some named object text/bbox contracts could not be proven", unresolved=object_semantic_uncertain)
                else:
                    report.add("slide1.named_object_content_and_bbox", "PASS", "HARD", "Named object literal content.text and normalized slide bboxes match the scene plan")
                if object_geometry_failures:
                    report.add("slide1.named_object_shape_geometry", "FAIL", "HARD", "Named object DrawingML geometry differs from plan style.shape_geometry", mismatches=object_geometry_failures)
                elif object_geometry_uncertain:
                    report.add("slide1.named_object_shape_geometry", "WARN", "HARD", "Some named object geometry contracts cannot be proven from preset DrawingML geometry", unresolved=object_geometry_uncertain)
                else:
                    report.add("slide1.named_object_shape_geometry", "PASS", "HARD", "Every named object geometry matches normalized plan style.shape_geometry")
                unresolved_math = [item for item in manual_math_reviews if not item["normalized_exact_match"]]
                if unresolved_math:
                    report.add("slide1.math_source_semantics", "WARN", "HARD", "OMML m:t text could not be normalized to an exact math_source match; formula semantics require explicit repair or stronger evidence", formulas=unresolved_math)
                else:
                    report.add("slide1.math_source_semantics", "PASS", "HARD", "Every declared math_source exactly matches normalized object-local OMML m:t text, or no math_source is declared", formulas_checked=len(manual_math_reviews))

                connector_semantic_failures: list[dict[str, Any]] = []
                connector_semantic_uncertain: list[dict[str, Any]] = []
                names_by_numeric_id = {
                    str(item["id"]): item["name"]
                    for item in slides[0]["named_objects"]
                    if item.get("id") is not None and isinstance(item.get("name"), str)
                }

                def arrow_location(connector: dict[str, Any]) -> str:
                    has_start = str(connector.get("head_arrow_type") or "none").lower() not in {"", "none"}
                    has_end = str(connector.get("tail_arrow_type") or "none").lower() not in {"", "none"}
                    return "both" if has_start and has_end else "start" if has_start else "end" if has_end else "none"

                for connection in plan.get("connections", []):
                    if not isinstance(connection, dict) or not isinstance(connection.get("id"), str):
                        continue
                    connection_id = connection["id"]
                    reconstruction = connection.get("reconstruction") if isinstance(connection.get("reconstruction"), dict) else {}
                    implementation = connection.get("implementation") if isinstance(connection.get("implementation"), dict) else {}
                    kind = normalize_ooxml_kind(implementation.get("ppt_object_kind") or reconstruction.get("expected_ooxml_kind"))
                    if kind not in {"connector", "attached_connector"}:
                        continue
                    actual_connector = objects_by_name.get(connection_id)
                    if actual_connector is None or actual_connector["type"] != "connector":
                        connector_semantic_failures.append({"connection_id": connection_id, "field": "named_connector", "reason": "expected named p:cxnSp not found"})
                        continue
                    planned_from = connection.get("from") if isinstance(connection.get("from"), dict) else {}
                    planned_to = connection.get("to") if isinstance(connection.get("to"), dict) else {}
                    actual_from = names_by_numeric_id.get(str(actual_connector.get("start_shape_id")))
                    actual_to = names_by_numeric_id.get(str(actual_connector.get("end_shape_id")))
                    if actual_from != planned_from.get("object_id") or actual_to != planned_to.get("object_id"):
                        connector_semantic_failures.append({
                            "connection_id": connection_id,
                            "field": "attachments",
                            "expected_from": planned_from.get("object_id"), "actual_from": actual_from,
                            "expected_to": planned_to.get("object_id"), "actual_to": actual_to,
                        })
                    style = connection.get("style") if isinstance(connection.get("style"), dict) else {}
                    expected_arrow = style.get("arrow_at")
                    if isinstance(expected_arrow, str) and arrow_location(actual_connector) != expected_arrow.lower():
                        connector_semantic_failures.append({"connection_id": connection_id, "field": "style.arrow_at", "expected": expected_arrow, "actual": arrow_location(actual_connector)})
                    expected_dash = style.get("dash")
                    if isinstance(expected_dash, str):
                        dash_alias = {"none": "solid", "continuous": "solid"}
                        normalized_expected_dash = dash_alias.get(expected_dash.lower(), expected_dash.lower())
                        normalized_actual_dash = dash_alias.get(str(actual_connector.get("dash") or "solid").lower(), str(actual_connector.get("dash") or "solid").lower())
                        if normalized_expected_dash != normalized_actual_dash:
                            connector_semantic_failures.append({"connection_id": connection_id, "field": "style.dash", "expected": normalized_expected_dash, "actual": normalized_actual_dash})
                    expected_width_px = style.get("line_width_px")
                    actual_width_emu = actual_connector.get("line_width_emu")
                    if isinstance(expected_width_px, (int, float)) and not isinstance(expected_width_px, bool):
                        if not isinstance(actual_width_emu, int):
                            connector_semantic_uncertain.append({"connection_id": connection_id, "field": "style.line_width_px", "reason": "a:ln width is implicit"})
                        else:
                            expected_width_emu = float(expected_width_px) * 9525.0
                            width_error = abs(actual_width_emu - expected_width_emu) / expected_width_emu
                            if width_error > 0.25:
                                connector_semantic_failures.append({"connection_id": connection_id, "field": "style.line_width_px", "expected_px": expected_width_px, "actual_emu": actual_width_emu, "relative_error": width_error})
                    expected_color = style.get("line_color")
                    if isinstance(expected_color, str):
                        expected_hex = expected_color.strip().lstrip("#").upper()
                        actual_hex = str(actual_connector.get("line_color_srgb") or "").upper()
                        if actual_hex:
                            if actual_hex != expected_hex:
                                connector_semantic_failures.append({"connection_id": connection_id, "field": "style.line_color", "expected": expected_hex, "actual": actual_hex})
                        elif actual_connector.get("line_color_scheme"):
                            connector_semantic_uncertain.append({"connection_id": connection_id, "field": "style.line_color", "reason": "theme color resolution is not implemented", "scheme": actual_connector.get("line_color_scheme")})
                        else:
                            connector_semantic_uncertain.append({"connection_id": connection_id, "field": "style.line_color", "reason": "line color is implicit or unparsed"})
                if connector_semantic_failures:
                    report.add("slide1.connection_semantics", "FAIL", "HARD", "Named native connectors disagree with endpoint or typed-style contracts", failures=connector_semantic_failures, unresolved=connector_semantic_uncertain)
                elif connector_semantic_uncertain:
                    report.add("slide1.connection_semantics", "WARN", "HARD", "Some native connector style evidence is theme-based or implicit and was not treated as PASS", unresolved=connector_semantic_uncertain)
                else:
                    report.add("slide1.connection_semantics", "PASS", "HARD", "Named native connectors match planned endpoints, arrows, dash, line width, and directly encoded colors")

                required_connectors = sum(
                    expectation["normalized_ooxml_kind"] in {"connector", "attached_connector"}
                    for expectation in expected["connection_ooxml_expectations"]
                )
                required_attached_connectors = sum(
                    expectation["normalized_ooxml_kind"] == "attached_connector"
                    for expectation in expected["connection_ooxml_expectations"]
                )
                if required_connectors == 0:
                    report.add("slide1.connectors_against_plan", "PASS", "INFO", "No connection contract requires p:cxnSp; approved fallback kinds are audited separately", planned_connections=expected["connections"], required_connectors=0, actual=actual["connectors"])
                elif actual["connectors"] < required_connectors:
                    report.add("slide1.connectors_against_plan", "FAIL", "HARD", "Fewer p:cxnSp connectors than connection expected_ooxml_kind contracts require", planned_connections=expected["connections"], required=required_connectors, actual=actual["connectors"])
                else:
                    report.add("slide1.connectors_against_plan", "PASS", "HARD", "Slide 1 has enough p:cxnSp objects for the connection contracts that require them", planned_connections=expected["connections"], required=required_connectors, actual=actual["connectors"])
                if required_attached_connectors:
                    report.add(
                        "slide1.connector_attachment",
                        "PASS" if actual["connectors_attached_both_ends"] >= required_attached_connectors else "FAIL",
                        "HARD",
                        "Enough connectors retain both endpoint attachments" if actual["connectors_attached_both_ends"] >= required_attached_connectors else "Fewer connectors retain both endpoint attachments than explicitly promised",
                        required=required_attached_connectors,
                        actual=actual["connectors_attached_both_ends"],
                    )
                else:
                    report.add("slide1.connector_attachment", "PASS", "INFO", "No connection expected_ooxml_kind explicitly promises both-end attachment", attached_connector_count=actual["connectors_attached_both_ends"])

                if expected["text_objects"] == 0:
                    report.add("slide1.text_against_plan", "PASS", "INFO", "No explicit text/label object types were identified in the scene plan", expected=0, actual=actual["text_shapes"])
                elif actual["text_shapes"] == 0:
                    report.add("slide1.text_against_plan", "FAIL", "HARD", "Scene plan declares text-like objects but Slide 1 has no live DrawingML text", expected=expected["text_objects"], actual=0)
                elif actual["text_shapes"] < expected["text_objects"]:
                    report.add("slide1.text_against_plan", "WARN", "HARD", "Fewer text shapes than text-like plan objects were detected; merged text may be intentional but cannot be inferred", expected=expected["text_objects"], actual=actual["text_shapes"])
                else:
                    report.add("slide1.text_against_plan", "PASS", "HARD", "Slide 1 has live text structures consistent with the plan", expected=expected["text_objects"], actual=actual["text_shapes"])

                if expected["formula_objects"]:
                    if actual["omml_objects"]:
                        report.add("slide1.omml_equations", "PASS", "HARD", "OMML equation objects were found for formula-like plan objects", expected_formula_objects=expected["formula_objects"], omml_objects=actual["omml_objects"])
                    else:
                        report.add("slide1.omml_equations", "WARN", "HARD", "Formula-like objects are planned but no OMML was found; the plan does not prove that OMML was promised", expected_formula_objects=expected["formula_objects"], omml_objects=0)
                else:
                    report.add("slide1.omml_equations", "PASS", "INFO", "No formula-like object types were identified in the scene plan", omml_objects=actual["omml_objects"])

            source_width = source.get("display_width_px") or source.get("width_px")
            source_height = source.get("display_height_px") or source.get("height_px")

            if len(slides) >= 2:
                slide2 = slides[1]
                embedded_used: list[dict[str, Any]] = []
                broken_targets: list[dict[str, Any]] = []
                for relation in slide2["image_relationships"]:
                    if relation["id"] not in slide2["used_image_relationship_ids"]:
                        continue
                    target = relation["resolved_target"]
                    if relation["external"] or not target or target not in names:
                        broken_targets.append(relation)
                    else:
                        embedded_used.append({**relation, "sha256": sha256_bytes(zf.read(target)), "byte_size": zf.getinfo(target).file_size})
                if broken_targets:
                    report.add("slide2.reference_image_relationship", "FAIL", "HARD", "Slide 2 uses an external, missing, or unresolved image relationship", broken_relationships=broken_targets)
                elif len(slide2["pictures"]) == 1 and len(embedded_used) == 1 and len(slide2["image_relationships"]) == 1:
                    report.add("slide2.reference_image_relationship", "PASS", "HARD", "Slide 2 has exactly one picture backed by exactly one embedded image relationship", embedded_image=embedded_used[0])
                else:
                    report.add(
                        "slide2.reference_image_relationship",
                        "FAIL",
                        "HARD",
                        "Slide 2 must contain exactly one picture and one embedded image relationship",
                        picture_count=len(slide2["pictures"]),
                        used_embedded_image_count=len(embedded_used),
                        image_relationship_count=len(slide2["image_relationships"]),
                    )

                if source_sha:
                    matching = [item for item in embedded_used if item["sha256"] == source_sha]
                    report.add(
                        "slide2.reference_source_identity",
                        "PASS" if len(matching) == 1 and len(embedded_used) == 1 else "FAIL",
                        "HARD",
                        "The sole Slide 2 image is byte-identical to the locked source" if len(matching) == 1 and len(embedded_used) == 1 else "The sole-reference requirement or locked-source SHA-256 check failed",
                        expected_sha256=source_sha,
                        embedded_sha256=[item["sha256"] for item in embedded_used],
                    )
                else:
                    report.add("slide2.reference_source_identity", "WARN", "HARD", "Source SHA-256 is unavailable; exact reference identity cannot be proven")

                if len(slide2["pictures"]) == 1:
                    reference_picture = slide2["pictures"][0]
                    if reference_picture["cropped"] or reference_picture["tiled"]:
                        report.add("slide2.reference_no_crop", "FAIL", "HARD", "Slide 2 reference picture has a crop/fill offset or tile fill", crop=reference_picture["crop"], fill_rect=reference_picture["fill_rect"], tiled=reference_picture["tiled"])
                    else:
                        report.add("slide2.reference_no_crop", "PASS", "HARD", "Slide 2 reference picture has no a:srcRect/fillRect crop offsets and is not tiled", crop=reference_picture["crop"], fill_rect=reference_picture["fill_rect"], tiled=False)

                    if reference_picture["hidden"] or reference_picture["fully_transparent"]:
                        report.add("slide2.reference_visible", "FAIL", "HARD", "Slide 2 reference picture is hidden or fully transparent")
                    else:
                        report.add("slide2.reference_visible", "PASS", "HARD", "Slide 2 reference picture is not marked hidden or fully transparent")

                    if (
                        reference_picture["geometry_uncertain"]
                        or reference_picture["box"] is None
                        or not isinstance(source_width, (int, float))
                        or not isinstance(source_height, (int, float))
                        or source_width <= 0
                        or source_height <= 0
                        or width <= 0
                        or height <= 0
                    ):
                        report.add("slide2.reference_contain_center", "WARN", "HARD", "Reference geometry or source dimensions are incomplete; contain/center placement is unproven")
                    else:
                        scale = min(width / source_width, height / source_height)
                        expected_width = source_width * scale
                        expected_height = source_height * scale
                        expected_x = (width - expected_width) / 2
                        expected_y = (height - expected_height) / 2
                        box = reference_picture["box"]
                        normalized_errors = {
                            "x": abs(box["x_emu"] - expected_x) / width,
                            "y": abs(box["y_emu"] - expected_y) / height,
                            "width": abs(box["width_emu"] - expected_width) / width,
                            "height": abs(box["height_emu"] - expected_height) / height,
                        }
                        actual_picture_ratio = box["width_emu"] / box["height_emu"] if box["height_emu"] else 0
                        source_ratio = source_width / source_height
                        ratio_error = abs(actual_picture_ratio / source_ratio - 1) if actual_picture_ratio else float("inf")
                        geometry_matches = max(normalized_errors.values()) <= full_bleed_tolerance and ratio_error <= aspect_tolerance
                        report.add(
                            "slide2.reference_contain_center",
                            "PASS" if geometry_matches else "FAIL",
                            "HARD",
                            "Slide 2 source picture is uncropped and placed with contain/center geometry" if geometry_matches else "Slide 2 source picture geometry differs from contain/center placement or distorts the source aspect ratio",
                            actual_box_emu=box,
                            expected_box_emu={"x_emu": expected_x, "y_emu": expected_y, "width_emu": expected_width, "height_emu": expected_height},
                            normalized_errors=normalized_errors,
                            geometry_tolerance=full_bleed_tolerance,
                            aspect_ratio_error=ratio_error,
                            aspect_tolerance=aspect_tolerance,
                        )
                else:
                    report.add("slide2.reference_no_crop", "FAIL", "HARD", "Crop cannot be validated because Slide 2 does not contain exactly one picture")
                    report.add("slide2.reference_contain_center", "FAIL", "HARD", "Contain/center geometry cannot be validated because Slide 2 does not contain exactly one picture")

                extra_slide2_structures = {
                    "auto_shapes": slide2["shape_counts"]["auto_shapes"],
                    "connectors": slide2["shape_counts"]["connectors"],
                    "groups": slide2["shape_counts"]["groups"],
                    "graphic_frames": slide2["shape_counts"]["graphic_frames"],
                    "nonempty_text_runs": slide2["shape_counts"]["nonempty_text_runs"],
                }
                has_extras = any(extra_slide2_structures.values())
                report.add(
                    "slide2.reference_only_content",
                    "PASS" if not has_extras else "FAIL",
                    "HARD",
                    "Slide 2 contains only the source picture" if not has_extras else "Slide 2 contains additional drawing or text structures beyond the source picture",
                    extra_structures=extra_slide2_structures,
                )

            canvas = plan.get("canvas") if isinstance(plan.get("canvas"), dict) else {}
            canvas_width = canvas.get("width_inches")
            canvas_height = canvas.get("height_inches")
            ratios: dict[str, float] = {}
            if width and height:
                ratios["pptx"] = width / height
            if isinstance(source_width, (int, float)) and isinstance(source_height, (int, float)) and source_width > 0 and source_height > 0:
                ratios["source"] = source_width / source_height
            if isinstance(canvas_width, (int, float)) and isinstance(canvas_height, (int, float)) and canvas_width > 0 and canvas_height > 0:
                ratios["scene_plan_canvas"] = canvas_width / canvas_height
            if len(ratios) == 3:
                mismatch = max(abs(ratios["pptx"] / ratios["source"] - 1), abs(ratios["pptx"] / ratios["scene_plan_canvas"] - 1))
                report.add(
                    "presentation.aspect_ratio",
                    "PASS" if mismatch <= aspect_tolerance else "FAIL",
                    "HARD",
                    "PPTX, source, and scene-plan canvas aspect ratios agree" if mismatch <= aspect_tolerance else "PPTX aspect ratio differs from the source or scene-plan canvas",
                    ratios=ratios,
                    maximum_relative_error=mismatch,
                    tolerance=aspect_tolerance,
                )
            else:
                report.add("presentation.aspect_ratio", "WARN", "HARD", "Source, canvas, or PPTX dimensions are missing; aspect-ratio agreement cannot be proven", ratios=ratios)

            approvals = plan.get("approvals") if isinstance(plan.get("approvals"), dict) else {}
            degradations = approvals.get("degradations") if isinstance(approvals.get("degradations"), list) else []
            unapproved = [item for item in degradations if not isinstance(item, dict) or item.get("approved") is not True]
            report.add(
                "plan.degradations_approved",
                "PASS" if not unapproved else "FAIL",
                "HARD",
                "Every recorded degradation is explicitly approved" if not unapproved else "The scene plan contains unapproved or malformed degradations",
                degradation_count=len(degradations),
                unapproved_count=len(unapproved),
            )

            parse_warnings = [*presentation_warnings]
            for slide in slides:
                parse_warnings.extend(f"Slide {slide['index']}: {warning}" for warning in slide["warnings"])
            report.add(
                "presentation.relationship_parse",
                "PASS" if not parse_warnings else "WARN",
                "HARD",
                "All audited slide/relationship XML paths were resolved without parser uncertainty" if not parse_warnings else "Some slide or relationship paths could not be parsed completely",
                warnings=parse_warnings,
            )

            return report.finish(
                input={
                    "pptx": str(pptx),
                    "pptx_sha256": sha256_bytes(pptx.read_bytes()),
                    "source_manifest": str(source_manifest_path.expanduser().resolve()),
                    "source_manifest_sha256": sha256_bytes(source_manifest_path.expanduser().resolve().read_bytes()),
                    "scene_plan": str(scene_plan_path.expanduser().resolve()),
                    "scene_plan_sha256": sha256_bytes(scene_plan_path.expanduser().resolve().read_bytes()),
                    "source_sha256": source_sha,
                },
                presentation={
                    "slide_count": len(slide_parts),
                    "slide_width_emu": width,
                    "slide_height_emu": height,
                    "slide_width_inches": width / EMU_PER_INCH if width else None,
                    "slide_height_inches": height / EMU_PER_INCH if height else None,
                    "hidden_slides": hidden,
                    "orphan_slide_parts": orphan_slide_parts,
                },
                package={
                    "media_parts": media,
                    "vector_media_parts": vector_media,
                    "macro_parts": macro_parts,
                    "embedding_parts": embedding_parts,
                    "external_relationships": external,
                },
                plan_expected_counts=expected,
                slides=slides,
            )
    except (OSError, zipfile.BadZipFile) as exc:
        raise AuditError(f"Cannot audit PPTX '{pptx}': {exc}") from exc


class JsonArgumentParser(argparse.ArgumentParser):
    def error(self, message: str) -> None:
        print(json.dumps({"status": "FAIL", "hard_failures": [{"id": "cli.invalid_arguments", "message": message}], "warnings": [], "checks": []}, ensure_ascii=False), file=sys.stderr)
        raise SystemExit(2)


def build_parser() -> argparse.ArgumentParser:
    parser = JsonArgumentParser(description="Audit a sci-diagram-pptx package using ZIP/XML evidence only.")
    parser.add_argument("pptx", type=Path, help="PPTX deliverable to audit")
    parser.add_argument("--source-manifest", type=Path, required=True, help="source_preflight JSON report")
    parser.add_argument("--scene-plan", type=Path, required=True, help="validated scene-plan JSON")
    parser.add_argument("--output", type=Path, required=True, help="write the audit JSON here")
    parser.add_argument("--full-bleed-tolerance", type=float, default=0.01, help="fractional geometry tolerance (default: 0.01)")
    parser.add_argument("--aspect-tolerance", type=float, default=0.005, help="maximum relative aspect-ratio error (default: 0.005)")
    return parser


def operational_failure(message: str) -> dict[str, Any]:
    check = {"id": "audit.operational", "status": "FAIL", "severity": "HARD", "message": message}
    return {
        "schema_version": "1.0",
        "kind": "sci-diagram-pptx-pptx-audit",
        "tool": "sci-diagram-pptx/audit_pptx.py",
        "tool_version": TOOL_VERSION,
        "generated_at": utc_now(),
        "status": "FAIL",
        "hard_failures": [{"id": check["id"], "message": message, "severity": "HARD"}],
        "warnings": [],
        "checks": [check],
    }


def main(argv: Iterable[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if not 0 <= args.full_bleed_tolerance <= 0.1:
        build_parser().error("--full-bleed-tolerance must be between 0 and 0.1")
    if not 0 <= args.aspect_tolerance <= 0.1:
        build_parser().error("--aspect-tolerance must be between 0 and 0.1")
    try:
        payload = audit_pptx(args.pptx, args.source_manifest, args.scene_plan, args.full_bleed_tolerance, args.aspect_tolerance)
        write_json(payload, args.output)
        return 1 if payload["hard_failures"] else 0
    except (AuditError, OSError) as exc:
        payload = operational_failure(str(exc))
        try:
            write_json(payload, args.output)
        except OSError:
            print(json.dumps(payload, ensure_ascii=False), file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
