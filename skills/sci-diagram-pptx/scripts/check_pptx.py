#!/usr/bin/env python3
"""Run a focused structural check on an editable scientific-diagram PPTX."""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import os
import posixpath
import re
import sys
import zipfile
from pathlib import Path
from urllib.parse import unquote
from xml.etree import ElementTree as ET


NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
}
R_ID = f"{{{NS['r']}}}id"
R_EMBED = f"{{{NS['r']}}}embed"
R_LINK = f"{{{NS['r']}}}link"
REL_NS = NS["rel"]
FONT_SCRIPTS = ("latin", "ea", "cs")
AUTOFIT_NAMES = {"noAutofit", "normAutofit", "spAutoFit"}
LOCAL_PATH_PATTERN = re.compile(
    r"(?<![A-Za-z0-9._-])(?:/(?:Users|Volumes|home)(?:/|\b)|/(?:private/)?tmp(?:/|\b)|[A-Za-z]:[\\/])"
)
IMPORT_PATTERNS = (
    re.compile(r"(?m)^\s*import\s+(?:[^;\n]*?\s+from\s+)?['\"]([^'\"]+)['\"]"),
    re.compile(r"(?m)^\s*export\s+[^;\n]*?\s+from\s+['\"]([^'\"]+)['\"]"),
    re.compile(r"\bimport\s*\(\s*['\"]([^'\"]+)['\"]\s*\)"),
    re.compile(r"\brequire\s*\(\s*['\"]([^'\"]+)['\"]\s*\)"),
)
SECRET_PATTERNS = (
    re.compile(r"-----BEGIN (?:RSA |EC |OPENSSH )?PRIVATE KEY-----"),
    re.compile(r"\b(?:sk|rk)-[A-Za-z0-9_-]{20,}\b"),
    re.compile(r"\bgh[pousr]_[A-Za-z0-9]{20,}\b"),
    re.compile(r"\bAKIA[0-9A-Z]{16}\b"),
)


class CheckError(RuntimeError):
    pass


def sha256_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def parse_xml(archive: zipfile.ZipFile, part: str) -> ET.Element:
    try:
        return ET.fromstring(archive.read(part))
    except KeyError as exc:
        raise CheckError(f"missing required package part: {part}") from exc
    except ET.ParseError as exc:
        raise CheckError(f"malformed XML in {part}: {exc}") from exc


def relationship_part(owner: str) -> str:
    folder, name = posixpath.split(owner)
    return posixpath.join(folder, "_rels", name + ".rels")


def resolve_target(owner: str, target: str) -> str | None:
    target = unquote(target).replace("\\", "/")
    resolved = (
        posixpath.normpath(target.lstrip("/"))
        if target.startswith("/")
        else posixpath.normpath(posixpath.join(posixpath.dirname(owner), target))
    )
    if resolved == ".." or resolved.startswith("../"):
        return None
    return resolved


def relationships(archive: zipfile.ZipFile, owner: str) -> dict[str, dict[str, object]]:
    part = relationship_part(owner)
    if part not in archive.namelist():
        return {}
    root = parse_xml(archive, part)
    result: dict[str, dict[str, object]] = {}
    for node in root.findall(f"{{{REL_NS}}}Relationship"):
        rid = node.get("Id")
        target = node.get("Target")
        kind = node.get("Type") or ""
        if not rid or not target:
            continue
        external = (node.get("TargetMode") or "").lower() == "external"
        result[rid] = {
            "type": kind,
            "target": target,
            "external": external,
            "resolved": None if external else resolve_target(owner, target),
        }
    return result


def relationship_owner(part: str) -> str | None:
    """Return the package part that owns an OPC relationship part."""
    if part == "_rels/.rels":
        return ""
    marker = "/_rels/"
    if marker not in part or not part.endswith(".rels"):
        return None
    folder, name = part.rsplit(marker, 1)
    if not name or name == ".rels":
        return None
    return posixpath.join(folder, name[:-5])


def validate_relationship_targets(archive: zipfile.ZipFile) -> None:
    """Reject malformed, orphaned, escaping, or dangling internal relationships."""
    names = set(archive.namelist())
    for rel_part in sorted(name for name in names if name.endswith(".rels")):
        owner = relationship_owner(rel_part)
        if owner is None:
            raise CheckError(f"invalid relationship part location: {rel_part}")
        if owner and owner not in names:
            raise CheckError(f"relationship part has no owning package part: {rel_part}")
        root = parse_xml(archive, rel_part)
        if root.tag != f"{{{REL_NS}}}Relationships":
            raise CheckError(f"invalid relationship root element in {rel_part}")
        identifiers: set[str] = set()
        for node in root.findall(f"{{{REL_NS}}}Relationship"):
            identifier = node.get("Id")
            target = node.get("Target")
            kind = node.get("Type")
            if not identifier or not target or not kind:
                raise CheckError(f"incomplete relationship in {rel_part}")
            if identifier in identifiers:
                raise CheckError(f"duplicate relationship Id {identifier!r} in {rel_part}")
            identifiers.add(identifier)
            if (node.get("TargetMode") or "").lower() == "external":
                continue
            resolved = resolve_target(owner, target)
            if resolved is None:
                raise CheckError(
                    f"internal relationship escapes the package: {rel_part} -> {target}"
                )
            if resolved not in names:
                raise CheckError(
                    f"internal relationship target is missing: {rel_part} -> {resolved}"
                )


def content_type_map(archive: zipfile.ZipFile) -> tuple[dict[str, str], dict[str, str]]:
    root = parse_xml(archive, "[Content_Types].xml")
    if root.tag != f"{{{NS['ct']}}}Types":
        raise CheckError("invalid [Content_Types].xml root element")
    defaults: dict[str, str] = {}
    overrides: dict[str, str] = {}
    for node in root.findall("ct:Default", NS):
        extension = (node.get("Extension") or "").lower()
        content_type = node.get("ContentType") or ""
        if not extension or not content_type or extension in defaults:
            raise CheckError("invalid or duplicate Default in [Content_Types].xml")
        defaults[extension] = content_type
    for node in root.findall("ct:Override", NS):
        part_name = (node.get("PartName") or "").lstrip("/")
        content_type = node.get("ContentType") or ""
        if (
            not part_name
            or not content_type
            or resolve_target("", part_name) != part_name
            or part_name in overrides
        ):
            raise CheckError("invalid or duplicate Override in [Content_Types].xml")
        overrides[part_name] = content_type
    return defaults, overrides


def package_content_type(
    part: str,
    defaults: dict[str, str],
    overrides: dict[str, str],
) -> str | None:
    if part in overrides:
        return overrides[part]
    extension = posixpath.basename(part).rsplit(".", 1)[-1].lower()
    return defaults.get(extension)


def validate_content_types(archive: zipfile.ZipFile) -> None:
    names = set(archive.namelist())
    defaults, overrides = content_type_map(archive)
    for part in sorted(
        name for name in names
        if name != "[Content_Types].xml" and not name.endswith("/")
    ):
        if package_content_type(part, defaults, overrides) is None:
            raise CheckError(f"package part has no content type declaration: {part}")

    presentation_type = package_content_type("ppt/presentation.xml", defaults, overrides)
    if presentation_type != (
        "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
    ):
        raise CheckError("ppt/presentation.xml has an invalid or missing content type")
    slide_type = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
    for part in sorted(
        name for name in names
        if name.startswith("ppt/slides/") and name.endswith(".xml")
    ):
        if package_content_type(part, defaults, overrides) != slide_type:
            raise CheckError(f"slide part has an invalid or missing content type: {part}")


def validate_root_office_document(archive: zipfile.ZipFile) -> None:
    root_rels = relationships(archive, "")
    office_documents = [
        relation for relation in root_rels.values()
        if str(relation.get("type", "")).endswith("/officeDocument")
    ]
    if len(office_documents) != 1:
        raise CheckError("package root must contain exactly one officeDocument relationship")
    relation = office_documents[0]
    if relation.get("external") or relation.get("resolved") != "ppt/presentation.xml":
        raise CheckError("package root officeDocument relationship does not resolve to ppt/presentation.xml")


def validate_slide_blips(
    archive: zipfile.ZipFile,
    slide_part: str,
    slide: ET.Element,
) -> None:
    slide_rels = relationships(archive, slide_part)
    defaults, overrides = content_type_map(archive)
    for blip in slide.findall(".//a:blip", NS):
        references = [
            (attribute, blip.get(attribute))
            for attribute in (R_EMBED, R_LINK)
            if blip.get(attribute)
        ]
        if not references:
            raise CheckError(f"image blip has no relationship in {slide_part}")
        for _, identifier in references:
            relation = slide_rels.get(identifier or "")
            if relation is None or not str(relation.get("type", "")).endswith("/image"):
                raise CheckError(
                    f"image blip has an unresolved or non-image relationship in {slide_part}: {identifier}"
                )
            if not relation.get("external"):
                target = relation.get("resolved")
                if not isinstance(target, str) or target not in archive.namelist():
                    raise CheckError(
                        f"image blip target is missing in {slide_part}: {identifier}"
                    )
                content_type = package_content_type(target, defaults, overrides) or ""
                if not content_type.lower().startswith("image/"):
                    raise CheckError(
                        f"image blip target has a non-image content type in {slide_part}: {target}"
                    )


def validate_build_source(path: Path) -> dict[str, object]:
    try:
        source_bytes = path.read_bytes()
    except OSError as exc:
        raise CheckError(f"cannot read build source: {exc}") from exc
    try:
        text = source_bytes.decode("utf-8")
    except UnicodeDecodeError as exc:
        raise CheckError("build source is not valid UTF-8") from exc
    if not text.strip():
        raise CheckError("build source is empty")
    if LOCAL_PATH_PATTERN.search(text):
        raise CheckError("build source contains a machine-local absolute path")
    if any(pattern.search(text) for pattern in SECRET_PATTERNS):
        raise CheckError("build source contains high-confidence secret material")

    imports = sorted({
        match.group(1)
        for pattern in IMPORT_PATTERNS
        for match in pattern.finditer(text)
    })
    unsupported = [
        specifier for specifier in imports
        if specifier != "@oai/artifact-tool" and not specifier.startswith("node:")
    ]
    if unsupported:
        raise CheckError(
            "build source imports an unshipped or unsupported helper: "
            + ", ".join(unsupported)
        )
    return check_item(
        "build_source.portable", "PASS",
        "Build source is non-empty UTF-8 and contains no machine-local path, high-confidence secret, or unshipped helper import",
        path=str(path), bytes=len(source_bytes), imports=imports,
    )


def integer(node: ET.Element | None, attribute: str) -> int | None:
    if node is None:
        return None
    try:
        return int(node.get(attribute, ""))
    except ValueError:
        return None


def object_box(node: ET.Element) -> tuple[int, int, int, int] | None:
    transform = node.find("./p:spPr/a:xfrm", NS)
    if transform is None:
        transform = node.find("./p:xfrm", NS)
    offset = transform.find("a:off", NS) if transform is not None else None
    extent = transform.find("a:ext", NS) if transform is not None else None
    values = (
        integer(offset, "x"), integer(offset, "y"),
        integer(extent, "cx"), integer(extent, "cy"),
    )
    if any(value is None for value in values):
        return None
    x, y, width, height = values
    return int(x), int(y), int(width), int(height)


def rotated_box(
    box: tuple[float, float, float, float],
    angle_degrees: float,
    center: tuple[float, float] | None = None,
) -> tuple[float, float, float, float]:
    if not angle_degrees % 360:
        return box
    x, y, width, height = box
    center_x, center_y = center or (x + width / 2, y + height / 2)
    radians = math.radians(angle_degrees)
    cosine, sine = math.cos(radians), math.sin(radians)
    corners = []
    for point_x, point_y in (
        (x, y), (x + width, y), (x + width, y + height), (x, y + height)
    ):
        delta_x, delta_y = point_x - center_x, point_y - center_y
        corners.append((
            center_x + delta_x * cosine - delta_y * sine,
            center_y + delta_x * sine + delta_y * cosine,
        ))
    left = min(point[0] for point in corners)
    top = min(point[1] for point in corners)
    right = max(point[0] for point in corners)
    bottom = max(point[1] for point in corners)
    return left, top, right - left, bottom - top


def image_box(
    node: ET.Element,
    parents: dict[ET.Element, ET.Element],
) -> tuple[tuple[int, int, int, int], float] | None:
    """Resolve an image-bearing object's box through standard nested group transforms."""
    raw_box = object_box(node)
    if raw_box is None:
        return None
    box: tuple[float, float, float, float] = tuple(float(value) for value in raw_box)
    transformed_area = abs(box[2] * box[3])

    own_transform = node.find("./p:spPr/a:xfrm", NS)
    if own_transform is None:
        own_transform = node.find("./p:xfrm", NS)
    own_rotation = integer(own_transform, "rot") or 0
    if own_rotation:
        box = rotated_box(box, own_rotation / 60000)

    ancestor = parents.get(node)
    while ancestor is not None:
        if ancestor.tag == f"{{{NS['p']}}}grpSp":
            transform = ancestor.find("./p:grpSpPr/a:xfrm", NS)
            offset = transform.find("a:off", NS) if transform is not None else None
            extent = transform.find("a:ext", NS) if transform is not None else None
            child_offset = transform.find("a:chOff", NS) if transform is not None else None
            child_extent = transform.find("a:chExt", NS) if transform is not None else None
            values = (
                integer(offset, "x"), integer(offset, "y"),
                integer(extent, "cx"), integer(extent, "cy"),
                integer(child_offset, "x"), integer(child_offset, "y"),
                integer(child_extent, "cx"), integer(child_extent, "cy"),
            )
            if any(value is None for value in values):
                return None
            group_x, group_y, group_width, group_height = values[:4]
            child_x, child_y, child_width, child_height = values[4:]
            if not child_width or not child_height:
                return None
            transformed_area *= abs(
                (group_width / child_width) * (group_height / child_height)
            )
            left = (box[0] - child_x) / child_width
            top = (box[1] - child_y) / child_height
            right = (box[0] + box[2] - child_x) / child_width
            bottom = (box[1] + box[3] - child_y) / child_height
            if (transform.get("flipH") or "").lower() in {"1", "true"}:
                left, right = 1 - right, 1 - left
            if (transform.get("flipV") or "").lower() in {"1", "true"}:
                top, bottom = 1 - bottom, 1 - top
            box = (
                group_x + left * group_width,
                group_y + top * group_height,
                (right - left) * group_width,
                (bottom - top) * group_height,
            )
            group_rotation = integer(transform, "rot") or 0
            if group_rotation:
                box = rotated_box(
                    box,
                    group_rotation / 60000,
                    (group_x + group_width / 2, group_y + group_height / 2),
                )
        ancestor = parents.get(ancestor)
    return tuple(int(round(value)) for value in box), transformed_area


def clip_box_to_slide(
    box: tuple[int, int, int, int],
    slide_width: int,
    slide_height: int,
) -> tuple[int, int, int, int] | None:
    x, y, width, height = box
    left = max(0, x)
    top = max(0, y)
    right = min(slide_width, x + width)
    bottom = min(slide_height, y + height)
    if right <= left or bottom <= top:
        return None
    return left, top, right - left, bottom - top


def union_area(rectangles: list[tuple[int, int, int, int]]) -> int:
    valid = [(x, y, x + w, y + h) for x, y, w, h in rectangles if w > 0 and h > 0]
    xs = sorted({value for rectangle in valid for value in (rectangle[0], rectangle[2])})
    area = 0
    for left, right in zip(xs, xs[1:]):
        if right <= left:
            continue
        intervals = sorted(
            (top, bottom)
            for x1, top, x2, bottom in valid
            if x1 < right and x2 > left
        )
        covered = 0
        if intervals:
            start, end = intervals[0]
            for next_start, next_end in intervals[1:]:
                if next_start <= end:
                    end = max(end, next_end)
                else:
                    covered += end - start
                    start, end = next_start, next_end
            covered += end - start
        area += (right - left) * covered
    return area


def check_item(identifier: str, status: str, message: str, **evidence: object) -> dict[str, object]:
    item: dict[str, object] = {"id": identifier, "status": status, "message": message}
    if evidence:
        item["evidence"] = evidence
    return item


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def text_preview(value: str, limit: int = 100) -> str:
    compact = " ".join(value.split())
    return compact if len(compact) <= limit else compact[: limit - 1] + "\u2026"


def shape_text(shape: ET.Element) -> str:
    return " ".join(
        (node.text or "").strip()
        for node in shape.findall(".//a:t", NS)
        if (node.text or "").strip()
    )


def shape_evidence(shape: ET.Element | None) -> dict[str, object]:
    if shape is None:
        return {}
    properties = shape.find("./p:nvSpPr/p:cNvPr", NS)
    evidence: dict[str, object] = {}
    if properties is not None:
        if properties.get("id"):
            evidence["shape_id"] = properties.get("id")
        if properties.get("name"):
            evidence["name"] = properties.get("name")
    text = shape_text(shape)
    if text:
        evidence["text"] = text_preview(text)
    return evidence


def ancestor_shape(node: ET.Element, parents: dict[ET.Element, ET.Element]) -> ET.Element | None:
    current: ET.Element | None = node
    while current is not None:
        if current.tag == f"{{{NS['p']}}}sp":
            return current
        current = parents.get(current)
    return None


def font_declarations(properties: ET.Element | None) -> dict[str, str | None]:
    if properties is None:
        return {}
    result: dict[str, str | None] = {}
    for script in FONT_SCRIPTS:
        node = properties.find(f"a:{script}", NS)
        if node is not None:
            result[script] = node.get("typeface")
    return result


def effective_run_fonts(
    run: ET.Element,
    parents: dict[ET.Element, ET.Element],
) -> dict[str, str | None]:
    direct = font_declarations(run.find("a:rPr", NS))
    if direct:
        return direct

    paragraph = parents.get(run)
    if paragraph is not None and paragraph.tag == f"{{{NS['a']}}}p":
        paragraph_default = font_declarations(paragraph.find("./a:pPr/a:defRPr", NS))
        if paragraph_default:
            return paragraph_default

        text_body = parents.get(paragraph)
        if text_body is not None:
            paragraph_properties = paragraph.find("a:pPr", NS)
            level = integer(paragraph_properties, "lvl") or 0
            list_default = font_declarations(text_body.find(
                f"./a:lstStyle/a:lvl{level + 1}pPr/a:defRPr", NS
            ))
            if list_default:
                return list_default
    return {}


def has_unicode_script_character(value: str) -> list[str]:
    characters = {
        character for character in value
        if character in "\u00b2\u00b3\u00b9" or "\u2070" <= character <= "\u209f"
    }
    return sorted(characters, key=ord)


def slide1_compatibility_checks(
    slide: ET.Element,
) -> tuple[list[dict[str, object]], list[dict[str, object]], list[dict[str, object]]]:
    """Return hard failures, warnings, and passes for cross-platform-risk structures."""
    failures: list[dict[str, object]] = []
    warnings: list[dict[str, object]] = []
    passes: list[dict[str, object]] = []
    parents = {child: node for node in slide.iter() for child in node}

    custom_geometry: list[dict[str, object]] = []
    missing_text_rectangles: list[dict[str, object]] = []
    for shape in slide.findall(".//p:sp", NS):
        geometry = shape.find("./p:spPr/a:custGeom", NS)
        if geometry is None:
            continue
        text = shape_text(shape)
        text_rectangle = geometry.find("a:rect", NS)
        has_text_rectangle = (
            text_rectangle is not None
            and all(text_rectangle.get(edge) for edge in ("l", "t", "r", "b"))
        )
        record = shape_evidence(shape)
        record.update({
            "text_bearing": bool(text),
            "has_text_rectangle": has_text_rectangle,
        })
        custom_geometry.append(record)
        if text and not has_text_rectangle:
            missing_text_rectangles.append(record)

    if custom_geometry:
        warnings.append(check_item(
            "slide1.custom_geometry_inventory", "WARN",
            "Slide 1 contains custom geometry that merits a native PowerPoint compatibility check",
            custom_geometry_count=len(custom_geometry),
            text_bearing_count=sum(bool(item["text_bearing"]) for item in custom_geometry),
            objects=custom_geometry[:25],
            truncated=max(0, len(custom_geometry) - 25),
        ))
    else:
        passes.append(check_item(
            "slide1.custom_geometry_inventory", "PASS",
            "Slide 1 contains no custom geometry",
            custom_geometry_count=0,
        ))

    if missing_text_rectangles:
        failures.append(check_item(
            "slide1.custom_geometry_text_rect", "FAIL",
            "Text-bearing custom geometry lacks an explicit DrawingML text rectangle",
            objects=missing_text_rectangles[:25],
            truncated=max(0, len(missing_text_rectangles) - 25),
        ))
    else:
        passes.append(check_item(
            "slide1.custom_geometry_text_rect", "PASS",
            "Every text-bearing custom geometry has an explicit text rectangle",
            text_bearing_custom_geometry=sum(
                bool(item["text_bearing"]) for item in custom_geometry
            ),
        ))

    typeface_counts: dict[str, dict[str, int]] = {script: {} for script in FONT_SCRIPTS}
    theme_references: list[dict[str, object]] = []
    empty_typefaces: list[dict[str, object]] = []
    for script in FONT_SCRIPTS:
        for font in slide.findall(f".//a:{script}", NS):
            typeface = font.get("typeface")
            label = typeface if typeface else "(empty)"
            typeface_counts[script][label] = typeface_counts[script].get(label, 0) + 1
    for script, values in typeface_counts.items():
        for typeface, count in sorted(values.items()):
            record = {"script": script, "typeface": typeface, "count": count}
            if typeface.startswith("+"):
                theme_references.append(record)
            elif typeface == "(empty)":
                empty_typefaces.append(record)

    text_runs = []
    for tag in ("r", "fld"):
        text_runs.extend(slide.findall(f".//a:{tag}", NS))
    missing_font_runs: list[dict[str, object]] = []
    partial_font_runs = 0
    visible_text_run_count = 0
    for run in text_runs:
        text = "".join(node.text or "" for node in run.findall("a:t", NS)).strip()
        if not text:
            continue
        visible_text_run_count += 1
        declarations = effective_run_fonts(run, parents)
        if not declarations:
            record = shape_evidence(ancestor_shape(run, parents))
            record["run_text"] = text_preview(text)
            missing_font_runs.append(record)
        elif len(declarations) < len(FONT_SCRIPTS):
            partial_font_runs += 1

    font_evidence = {
        "text_run_count": visible_text_run_count,
        "typefaces": {
            script: [
                {"typeface": typeface, "count": count}
                for typeface, count in sorted(values.items())
            ]
            for script, values in typeface_counts.items()
        },
        "theme_references": theme_references,
        "empty_typefaces": empty_typefaces,
        "runs_without_explicit_typeface": len(missing_font_runs),
        "runs_with_partial_script_declarations": partial_font_runs,
        "missing_examples": missing_font_runs[:15],
        "truncated": max(0, len(missing_font_runs) - 15),
    }
    if theme_references or empty_typefaces or missing_font_runs or partial_font_runs:
        warnings.append(check_item(
            "slide1.font_inventory", "WARN",
            "Some Slide 1 text uses a theme placeholder, empty typeface, or inherited font",
            **font_evidence,
        ))
    else:
        passes.append(check_item(
            "slide1.font_inventory", "PASS",
            "Slide 1 text has explicit font declarations; no platform font whitelist was applied",
            **font_evidence,
        ))

    text_body_count = 0
    wrap_none_count = 0
    norm_autofit_count = 0
    shape_autofit_count = 0
    no_autofit_count = 0
    invalid_autofit: list[dict[str, object]] = []
    layout_risks: list[dict[str, object]] = []
    for body_properties in slide.findall(".//a:bodyPr", NS):
        text_body = parents.get(body_properties)
        if text_body is None:
            continue
        visible_text = "".join(
            node.text or "" for node in text_body.findall(".//a:t", NS)
        ).strip()
        if not visible_text:
            continue
        text_body_count += 1
        autofit = [
            local_name(child.tag) for child in body_properties
            if local_name(child.tag) in AUTOFIT_NAMES
        ]
        wrap_none = (body_properties.get("wrap") or "").lower() == "none"
        wrap_none_count += int(wrap_none)
        norm_autofit_count += autofit.count("normAutofit")
        shape_autofit_count += autofit.count("spAutoFit")
        no_autofit_count += autofit.count("noAutofit")
        record = shape_evidence(ancestor_shape(body_properties, parents))
        record.update({"wrap": body_properties.get("wrap"), "autofit": autofit})
        if len(autofit) > 1:
            invalid_autofit.append(record)
        if wrap_none or "normAutofit" in autofit or "spAutoFit" in autofit:
            layout_risks.append(record)

    layout_evidence = {
        "text_body_count": text_body_count,
        "wrap_none_count": wrap_none_count,
        "no_autofit_count": no_autofit_count,
        "normal_autofit_count": norm_autofit_count,
        "shape_autofit_count": shape_autofit_count,
    }
    if invalid_autofit:
        failures.append(check_item(
            "slide1.text_autofit_exclusive", "FAIL",
            "A text body contains multiple mutually exclusive AutoFit children",
            **layout_evidence,
            objects=invalid_autofit[:25],
            truncated=max(0, len(invalid_autofit) - 25),
        ))
    else:
        passes.append(check_item(
            "slide1.text_autofit_exclusive", "PASS",
            "No text body contains multiple AutoFit children",
            **layout_evidence,
        ))

    if layout_risks:
        warnings.append(check_item(
            "slide1.text_layout_inventory", "WARN",
            "Slide 1 uses wrap=none or an AutoFit mode that may reflow across PowerPoint environments",
            **layout_evidence,
            objects=layout_risks[:25],
            truncated=max(0, len(layout_risks) - 25),
        ))
    else:
        passes.append(check_item(
            "slide1.text_layout_inventory", "PASS",
            "No wrap=none, normal AutoFit, or shape AutoFit text layout was found",
            **layout_evidence,
        ))

    math_paragraphs = list(slide.findall(".//m:oMathPara", NS))
    math_objects = list(slide.findall(".//m:oMath", NS))
    invalid_math_paragraphs = [
        paragraph for paragraph in math_paragraphs
        if paragraph.find(".//m:oMath", NS) is None
    ]
    math_records: list[dict[str, object]] = []
    empty_math: list[dict[str, object]] = []
    for math in math_objects:
        value = "".join(node.text or "" for node in math.findall(".//m:t", NS)).strip()
        record = {"kind": local_name(math.tag), "text": text_preview(value)}
        math_records.append(record)
        if not value:
            empty_math.append(record)
    invalid_math = [
        {
            "kind": local_name(paragraph.tag),
            "text": text_preview("".join(
                node.text or "" for node in paragraph.findall(".//m:t", NS)
            )),
            "reason": "oMathPara contains no oMath child",
        }
        for paragraph in invalid_math_paragraphs
    ]
    math_evidence = {
        "math_object_count": len(math_objects),
        "math_paragraph_count": len(math_paragraphs),
        "objects": math_records[:25],
        "truncated": max(0, len(math_records) - 25),
    }
    if empty_math or invalid_math:
        failures.append(check_item(
            "slide1.office_math", "FAIL",
            "Slide 1 contains an empty or incomplete Office Math structure",
            **math_evidence,
            empty_objects=empty_math[:25],
            invalid_objects=invalid_math[:25],
        ))
    else:
        passes.append(check_item(
            "slide1.office_math", "PASS",
            "Office Math objects are non-empty" if math_objects else "Slide 1 contains no Office Math objects",
            **math_evidence,
        ))

    unicode_script_runs: list[dict[str, object]] = []
    for text_node in slide.findall(".//a:t", NS):
        value = text_node.text or ""
        characters = has_unicode_script_character(value)
        if not characters:
            continue
        record = shape_evidence(ancestor_shape(text_node, parents))
        record.update({"run_text": text_preview(value), "characters": characters})
        unicode_script_runs.append(record)
    native_baseline_count = 0
    for tag in ("rPr", "defRPr", "endParaRPr"):
        for properties in slide.findall(f".//a:{tag}", NS):
            baseline = integer(properties, "baseline")
            if baseline:
                native_baseline_count += 1
    script_evidence = {
        "unicode_script_run_count": len(unicode_script_runs),
        "native_baseline_property_count": native_baseline_count,
        "objects": unicode_script_runs[:25],
        "truncated": max(0, len(unicode_script_runs) - 25),
    }
    if unicode_script_runs:
        warnings.append(check_item(
            "slide1.script_notation", "WARN",
            "Unicode superscript or subscript characters may trigger font fallback",
            **script_evidence,
        ))
    else:
        passes.append(check_item(
            "slide1.script_notation", "PASS",
            "No Unicode superscript or subscript characters were found",
            **script_evidence,
        ))

    return failures, warnings, passes


def slide_order(archive: zipfile.ZipFile) -> tuple[list[str], tuple[int, int]]:
    presentation_part = "ppt/presentation.xml"
    root = parse_xml(archive, presentation_part)
    rels = relationships(archive, presentation_part)
    slide_parts: list[str] = []
    for slide_id in root.findall("./p:sldIdLst/p:sldId", NS):
        relation = rels.get(slide_id.get(R_ID, ""))
        if not relation or not str(relation.get("type", "")).endswith("/slide"):
            raise CheckError("presentation contains an unresolved slide relationship")
        target = relation.get("resolved")
        if not isinstance(target, str) or target not in archive.namelist():
            raise CheckError("presentation references a missing slide part")
        slide_parts.append(target)
    size = root.find("./p:sldSz", NS)
    width, height = integer(size, "cx"), integer(size, "cy")
    if not slide_parts or not width or not height:
        raise CheckError("presentation has no usable slide or slide size")
    return slide_parts, (width, height)


def image_bearers(
    slide: ET.Element,
) -> tuple[list[tuple[int, int, int, int]], list[float], int, int, int]:
    parent = {child: node for node in slide.iter() for child in node}
    rectangles: list[tuple[int, int, int, int]] = []
    transformed_areas: list[float] = []
    grouped_count = 0
    unresolved_grouped = 0
    count = 0
    candidates = slide.findall(".//p:pic", NS) + slide.findall(".//p:sp", NS) + slide.findall(".//p:graphicFrame", NS)
    for node in candidates:
        if node.find(".//a:blip", NS) is None:
            continue
        count += 1
        ancestor = parent.get(node)
        grouped = False
        while ancestor is not None:
            if ancestor.tag == f"{{{NS['p']}}}grpSp":
                grouped = True
                break
            ancestor = parent.get(ancestor)
        if grouped:
            grouped_count += 1
        geometry = image_box(node, parent)
        if geometry:
            box, transformed_area = geometry
            rectangles.append(box)
            transformed_areas.append(transformed_area)
        elif grouped:
            unresolved_grouped += 1
    if slide.find("./p:cSld/p:bg/p:bgPr/a:blipFill", NS) is not None:
        count += 1
    return rectangles, transformed_areas, count, grouped_count, unresolved_grouped


def source_reference_check(
    archive: zipfile.ZipFile,
    slide_part: str,
    source: Path,
) -> tuple[list[dict[str, object]], list[dict[str, object]]]:
    failures: list[dict[str, object]] = []
    warnings: list[dict[str, object]] = []
    root = parse_xml(archive, slide_part)
    pictures = root.findall(".//p:pic", NS)
    all_blips = root.findall(".//a:blip", NS)
    if len(pictures) != 1 or len(all_blips) != 1:
        failures.append(check_item(
            "slide2.single_source_picture", "FAIL",
            "Reference slide must contain exactly one image",
            pictures=len(pictures), image_references=len(all_blips),
        ))
        return failures, warnings

    crop_nodes = pictures[0].findall(".//a:srcRect", NS) + pictures[0].findall(".//a:fillRect", NS)
    crop_values = []
    for crop in crop_nodes:
        crop_values.extend(integer(crop, edge) or 0 for edge in ("l", "t", "r", "b"))
    if any(crop_values):
        failures.append(check_item(
            "slide2.source_not_cropped", "FAIL",
            "Reference image contains an additional OOXML crop",
        ))

    relation_id = all_blips[0].get(R_EMBED)
    relation = relationships(archive, slide_part).get(relation_id or "")
    target = relation.get("resolved") if relation else None
    if not isinstance(target, str) or target not in archive.namelist():
        failures.append(check_item(
            "slide2.source_bytes", "FAIL", "Reference image relationship is unresolved"
        ))
    elif sha256_bytes(archive.read(target)) != sha256_file(source):
        failures.append(check_item(
            "slide2.source_bytes", "FAIL",
            "Reference slide does not embed the supplied source bytes",
        ))
    return failures, warnings


def inspect_package(
    pptx: Path,
    source: Path | None = None,
    *,
    require_single_slide: bool = False,
    build_source: Path | None = None,
) -> dict[str, object]:
    hard_failures: list[dict[str, object]] = []
    warnings: list[dict[str, object]] = []
    checks: list[dict[str, object]] = []
    pptx = pptx.expanduser().resolve()
    source = source.expanduser().resolve() if source else None
    build_source = build_source.expanduser().resolve() if build_source else None

    if build_source is not None:
        try:
            checks.append(validate_build_source(build_source))
        except CheckError as exc:
            hard_failures.append(check_item(
                "build_source.portable", "FAIL", str(exc), path=str(build_source)
            ))

    try:
        if not pptx.is_file():
            raise CheckError(f"PPTX does not exist: {pptx}")
        if source and not source.is_file():
            raise CheckError(f"source image does not exist: {source}")
        if not zipfile.is_zipfile(pptx):
            raise CheckError("file is not a readable ZIP/PPTX package")

        with zipfile.ZipFile(pptx) as archive:
            corrupt = archive.testzip()
            if corrupt:
                raise CheckError(f"corrupt ZIP member: {corrupt}")
            names = set(archive.namelist())
            if "[Content_Types].xml" not in names:
                raise CheckError("missing [Content_Types].xml")
            validate_content_types(archive)
            validate_relationship_targets(archive)
            validate_root_office_document(archive)
            slide_parts, (slide_width, slide_height) = slide_order(archive)
            for slide_part in slide_parts:
                slide = parse_xml(archive, slide_part)
                validate_slide_blips(archive, slide_part, slide)
            checks.append(check_item(
                "package.readable", "PASS",
                "PPTX content types, internal relationships, and slide order are readable",
                slide_count=len(slide_parts), slide_size_emu=[slide_width, slide_height],
            ))
            if require_single_slide:
                if len(slide_parts) == 1:
                    checks.append(check_item(
                        "delivery.single_slide", "PASS",
                        "The delivery contains exactly one editable slide",
                        slide_count=1,
                    ))
                else:
                    hard_failures.append(check_item(
                        "delivery.single_slide", "FAIL",
                        "The delivery must contain exactly one editable slide; keep the source image as an external companion file",
                        slide_count=len(slide_parts),
                    ))

            forbidden = sorted(
                name for name in names
                if name.endswith("vbaProject.bin")
                or "/embeddings/" in name
                or "/oleObjects/" in name
                or "/activeX/" in name
            )
            external_media = []
            external_other = []
            for rel_part in sorted(name for name in names if name.endswith(".rels")):
                rel_root = parse_xml(archive, rel_part)
                for rel in rel_root.findall(f"{{{REL_NS}}}Relationship"):
                    if (rel.get("TargetMode") or "").lower() != "external":
                        continue
                    kind = (rel.get("Type") or "").lower()
                    item = {"part": rel_part, "type": rel.get("Type"), "target": rel.get("Target")}
                    if any(word in kind for word in ("image", "audio", "video", "media", "oleobject", "package")):
                        external_media.append(item)
                    else:
                        external_other.append(item)
            if forbidden or external_media:
                hard_failures.append(check_item(
                    "package.no_unsafe_embeds", "FAIL",
                    "Package contains a macro, OLE/embedded object, ActiveX part, or external media",
                    package_parts=forbidden, external_media=external_media,
                ))
            else:
                checks.append(check_item(
                    "package.no_unsafe_embeds", "PASS",
                    "No macro, OLE/embedded object, ActiveX part, or external media was found",
                ))
            if external_other:
                warnings.append(check_item(
                    "package.external_relationships", "WARN",
                    "Package contains external relationships such as hyperlinks",
                    relationships=external_other,
                ))

            slide1 = parse_xml(archive, slide_parts[0])
            native_shapes = [
                node for node in slide1.findall(".//p:sp", NS)
                if node.find(".//a:blip", NS) is None
            ]
            connectors = slide1.findall(".//p:cxnSp", NS)
            text_runs = [node.text or "" for node in slide1.findall(".//a:t", NS) if (node.text or "").strip()]
            (
                rectangles,
                transformed_areas,
                raster_count,
                grouped_raster_count,
                unresolved_grouped_raster,
            ) = image_bearers(slide1)
            if not native_shapes and not connectors:
                hard_failures.append(check_item(
                    "slide1.native_objects", "FAIL",
                    "Slide 1 contains no native shape or connector",
                    native_shapes=0, connectors=0, text_runs=len(text_runs),
                ))
            else:
                checks.append(check_item(
                    "slide1.native_objects", "PASS",
                    "Slide 1 contains editable native objects",
                    native_shapes=len(native_shapes), connectors=len(connectors), text_runs=len(text_runs),
                ))

            slide_area = slide_width * slide_height
            clipped_rectangles: list[tuple[int, int, int, int]] = []
            visible_areas: list[float] = []
            for box, transformed_area in zip(rectangles, transformed_areas):
                clipped = clip_box_to_slide(box, slide_width, slide_height)
                if clipped is None:
                    continue
                clipped_rectangles.append(clipped)
                visible_areas.append(min(transformed_area, clipped[2] * clipped[3]))
            max_ratio = max(
                (visible_area / slide_area for visible_area in visible_areas),
                default=0.0,
            )
            union_ratio = (
                min(union_area(clipped_rectangles), sum(visible_areas)) / slide_area
                if clipped_rectangles else 0.0
            )
            has_image_background = slide1.find("./p:cSld/p:bg/p:bgPr/a:blipFill", NS) is not None
            raster_simulation = has_image_background or max_ratio >= 0.85 or (raster_count >= 2 and union_ratio >= 0.80)
            if raster_simulation:
                hard_failures.append(check_item(
                    "slide1.not_flattened", "FAIL",
                    "Slide 1 appears to use a whole-page image or image mosaic",
                    raster_objects=raster_count,
                    grouped_raster_objects=grouped_raster_count,
                    largest_coverage=round(max_ratio, 4),
                    union_coverage=round(union_ratio, 4),
                    image_background=has_image_background,
                ))
            else:
                checks.append(check_item(
                    "slide1.not_flattened", "PASS",
                    "No whole-page image or large image mosaic was detected",
                    raster_objects=raster_count,
                    grouped_raster_objects=grouped_raster_count,
                    largest_coverage=round(max_ratio, 4),
                    union_coverage=round(union_ratio, 4),
                ))
            if unresolved_grouped_raster:
                warnings.append(check_item(
                    "slide1.grouped_raster_geometry", "WARN",
                    "A grouped raster object has an incomplete transform and needs visual confirmation",
                    unresolved_objects=unresolved_grouped_raster,
                ))

            compatibility_failures, compatibility_warnings, compatibility_passes = (
                slide1_compatibility_checks(slide1)
            )
            hard_failures.extend(compatibility_failures)
            warnings.extend(compatibility_warnings)
            checks.extend(compatibility_passes)

            if require_single_slide and source:
                source_hash = sha256_file(source)
                matching_media = sorted(
                    part for part in names
                    if part.startswith("ppt/media/")
                    and not part.endswith("/")
                    and sha256_bytes(archive.read(part)) == source_hash
                )
                if matching_media:
                    hard_failures.append(check_item(
                        "source.not_embedded", "FAIL",
                        "The source image must remain external and must not be embedded in the PPTX",
                        matching_media=matching_media,
                        sha256=source_hash,
                    ))
                else:
                    checks.append(check_item(
                        "source.external_companion", "PASS",
                        "The supplied source image is external and no identical media bytes are embedded in the PPTX",
                        path=str(source),
                        bytes=source.stat().st_size,
                        sha256=source_hash,
                    ))
            elif source and len(slide_parts) >= 2:
                failures, source_warnings = source_reference_check(archive, slide_parts[1], source)
                hard_failures.extend(failures)
                warnings.extend(source_warnings)
                if not failures:
                    checks.append(check_item(
                        "slide2.source_reference", "PASS",
                        "Slide 2 embeds the supplied source bytes without an OOXML crop",
                    ))
            elif source:
                warnings.append(check_item(
                    "slide2.source_reference", "WARN",
                    "One-slide deliverable: supplied source is not embedded as a reference slide",
                ))
            elif len(slide_parts) >= 2:
                warnings.append(check_item(
                    "slide2.source_reference", "WARN",
                    "No --source was supplied, so the reference slide identity was not checked",
                ))

    except (OSError, zipfile.BadZipFile, CheckError) as exc:
        hard_failures.append(check_item("package.readable", "FAIL", str(exc)))

    return {
        "kind": "sci-diagram-pptx-check",
        "status": "FAIL" if hard_failures else "PASS",
        "pptx": str(pptx),
        "source": str(source) if source else None,
        "build_source": str(build_source) if build_source else None,
        "hard_failures": hard_failures,
        "warnings": warnings,
        "checks": checks,
    }


def write_report(report: dict[str, object], output: Path | None) -> None:
    rendered = json.dumps(report, ensure_ascii=False, indent=2) + "\n"
    if output is None:
        sys.stdout.write(rendered)
        return
    target = output.expanduser().resolve()
    target.parent.mkdir(parents=True, exist_ok=True)
    temporary = target.with_name(target.name + ".tmp")
    temporary.write_text(rendered, encoding="utf-8")
    os.replace(temporary, target)
    print(f"{report['status']}: {target}")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Run focused structural checks on a scientific-diagram PPTX.")
    parser.add_argument("pptx", type=Path, help="PPTX to inspect")
    parser.add_argument(
        "--source",
        type=Path,
        help=(
            "source image; treated as an external companion with --require-single-slide, "
            "otherwise used for the optional legacy Slide 2 identity check"
        ),
    )
    parser.add_argument(
        "--build-source",
        type=Path,
        help="optional executed build.mjs to check for basic portability hazards",
    )
    parser.add_argument(
        "--require-single-slide",
        action="store_true",
        help="require exactly one editable slide and keep --source external to the PPTX",
    )
    parser.add_argument("--output", type=Path, help="optional JSON report path")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    report = inspect_package(
        args.pptx,
        args.source,
        require_single_slide=args.require_single_slide,
        build_source=args.build_source,
    )
    write_report(report, args.output)
    return 0 if report["status"] == "PASS" else 1


if __name__ == "__main__":
    raise SystemExit(main())
