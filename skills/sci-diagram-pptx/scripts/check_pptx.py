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
    r"(?<![A-Za-z0-9._-])(?:/(?:Users|Volumes|home|root|srv|app|workspace|opt|etc|mnt|data|var)(?:/|\b)|/(?:private/)?tmp(?:/|\b)|[A-Za-z]:[\\/])"
)
MODULE_DECLARATION_PATTERN = re.compile(r"(?m)(?:^[ \t]*|;[ \t]*)(import|export)\b")
JS_GAP_PATTERN = r"(?:\s|/\*[\s\S]*?\*/|//[^\n]*(?:\n|$))*"
DYNAMIC_IMPORT_PATTERN = re.compile(rf"\bimport{JS_GAP_PATTERN}\(")
REQUIRE_PATTERN = re.compile(rf"\brequire{JS_GAP_PATTERN}\(")
CREATE_REQUIRE_PATTERN = re.compile(r"\bcreateRequire\b")
SECRET_PATTERNS = (
    re.compile(r"-----BEGIN (?:RSA |EC |OPENSSH )?PRIVATE KEY-----"),
    re.compile(r"\b(?:sk|rk)-[A-Za-z0-9_-]{20,}\b"),
    re.compile(r"\bgh[pousr]_[A-Za-z0-9]{20,}\b"),
    re.compile(r"\bAKIA[0-9A-Z]{16}\b"),
)
BUILD_SOURCE_PACKAGES = {"@oai/artifact-tool", "pptxgenjs"}
SAFE_NODE_MODULES = {"node:fs", "node:fs/promises", "node:path", "node:url"}

ZIP_MAX_MEMBERS = 10_000
ZIP_MAX_ENTRY_UNCOMPRESSED = 256 * 1024 * 1024
ZIP_MAX_TOTAL_UNCOMPRESSED = 1024 * 1024 * 1024
ZIP_MAX_XML_UNCOMPRESSED = 32 * 1024 * 1024
ZIP_MAX_COMPRESSION_RATIO = 200.0
ZIP_RATIO_MIN_UNCOMPRESSED = 1024 * 1024

UNSAFE_RELATIONSHIP_KINDS = {
    "activexcontrol",
    "activexcontrolbinary",
    "control",
    "controlprop",
    "controlproperties",
    "embeddedcontrolpersistence",
    "embeddedobject",
    "embeddedpackage",
    "oleobject",
    "package",
    "vbaproject",
}
UNSAFE_CONTENT_TYPE_FRAGMENTS = (
    "activex",
    "controlproperties",
    "oleobject",
    "vbaproject",
)


class CheckError(RuntimeError):
    def __init__(
        self,
        message: str,
        *,
        identifier: str = "package.readable",
        **evidence: object,
    ) -> None:
        super().__init__(message)
        self.identifier = identifier
        self.evidence = evidence


def sha256_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def is_xml_part(name: str) -> bool:
    lowered = name.lower()
    return lowered.endswith((".xml", ".rels")) or lowered == "[content_types].xml"


def validate_zip_budget(archive: zipfile.ZipFile) -> dict[str, object]:
    """Reject ambiguous or unusually expensive ZIP packages before reading members."""
    infos = archive.infolist()
    seen: set[str] = set()
    duplicates: list[str] = []
    total_uncompressed = 0
    largest_entry = 0
    largest_xml = 0
    highest_ratio = 0.0

    for info in infos:
        if info.filename in seen:
            duplicates.append(info.filename)
        seen.add(info.filename)
        if info.flag_bits & 0x1:
            raise CheckError(
                "encrypted ZIP members are not supported",
                identifier="package.encryption",
                member=info.filename,
            )
        if info.is_dir():
            continue
        total_uncompressed += info.file_size
        largest_entry = max(largest_entry, info.file_size)
        if info.file_size > ZIP_MAX_ENTRY_UNCOMPRESSED:
            raise CheckError(
                "ZIP member exceeds the uncompressed size budget",
                identifier="package.zip_budget",
                member=info.filename,
                uncompressed_bytes=info.file_size,
                limit_bytes=ZIP_MAX_ENTRY_UNCOMPRESSED,
                budget="single_member",
            )
        if total_uncompressed > ZIP_MAX_TOTAL_UNCOMPRESSED:
            raise CheckError(
                "ZIP package exceeds the total uncompressed size budget",
                identifier="package.zip_budget",
                uncompressed_bytes=total_uncompressed,
                limit_bytes=ZIP_MAX_TOTAL_UNCOMPRESSED,
                budget="total",
            )
        if is_xml_part(info.filename):
            largest_xml = max(largest_xml, info.file_size)
            if info.file_size > ZIP_MAX_XML_UNCOMPRESSED:
                raise CheckError(
                    "XML member exceeds the XML size budget",
                    identifier="package.zip_budget",
                    member=info.filename,
                    uncompressed_bytes=info.file_size,
                    limit_bytes=ZIP_MAX_XML_UNCOMPRESSED,
                    budget="xml_member",
                )
        if info.file_size >= ZIP_RATIO_MIN_UNCOMPRESSED:
            ratio = (
                float("inf") if info.compress_size == 0
                else info.file_size / info.compress_size
            )
            highest_ratio = max(highest_ratio, ratio)
            if ratio > ZIP_MAX_COMPRESSION_RATIO:
                raise CheckError(
                    "ZIP member has an extreme compression ratio",
                    identifier="package.zip_budget",
                    member=info.filename,
                    uncompressed_bytes=info.file_size,
                    compressed_bytes=info.compress_size,
                    compression_ratio=round(ratio, 2),
                    limit=ZIP_MAX_COMPRESSION_RATIO,
                    budget="compression_ratio",
                )

    if duplicates:
        unique_duplicates = sorted(set(duplicates))
        raise CheckError(
            "ZIP package contains duplicate member names",
            identifier="package.unique_members",
            members=unique_duplicates[:25],
            duplicate_count=len(unique_duplicates),
            truncated=max(0, len(unique_duplicates) - 25),
        )
    if len(infos) > ZIP_MAX_MEMBERS:
        raise CheckError(
            "ZIP package exceeds the member-count budget",
            identifier="package.zip_budget",
            member_count=len(infos),
            limit=ZIP_MAX_MEMBERS,
            budget="member_count",
        )
    return {
        "member_count": len(infos),
        "total_uncompressed_bytes": total_uncompressed,
        "largest_member_bytes": largest_entry,
        "largest_xml_member_bytes": largest_xml,
        "highest_checked_compression_ratio": round(highest_ratio, 2),
        "limits": {
            "members": ZIP_MAX_MEMBERS,
            "single_member_bytes": ZIP_MAX_ENTRY_UNCOMPRESSED,
            "total_uncompressed_bytes": ZIP_MAX_TOTAL_UNCOMPRESSED,
            "xml_member_bytes": ZIP_MAX_XML_UNCOMPRESSED,
            "compression_ratio": ZIP_MAX_COMPRESSION_RATIO,
            "ratio_minimum_uncompressed_bytes": ZIP_RATIO_MIN_UNCOMPRESSED,
        },
    }


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


def relationship_kind(value: str) -> str:
    return re.split(r"[/#]", value.rstrip("/"))[-1].lower()


def unsafe_embed_inventory(archive: zipfile.ZipFile) -> dict[str, object]:
    """Find unsafe active content by path, relationship, content type, and XML marker."""
    names = set(archive.namelist())
    defaults, overrides = content_type_map(archive)
    package_parts = sorted(
        name for name in names
        if not name.endswith("/")
        and (
            name.lower().endswith("vbaproject.bin")
            or "/embeddings/" in name.lower()
            or "/oleobjects/" in name.lower()
            or "/activex/" in name.lower()
        )
    )
    unsafe_internal_relationships: list[dict[str, object]] = []
    external_media: list[dict[str, object]] = []
    external_other: list[dict[str, object]] = []
    for rel_part in sorted(name for name in names if name.endswith(".rels")):
        rel_root = parse_xml(archive, rel_part)
        for relation in rel_root.findall(f"{{{REL_NS}}}Relationship"):
            kind = relation.get("Type") or ""
            record = {
                "part": rel_part,
                "id": relation.get("Id"),
                "type": kind,
                "target": relation.get("Target"),
            }
            if (relation.get("TargetMode") or "").lower() == "external":
                lowered = kind.lower()
                if any(word in lowered for word in (
                    "image", "audio", "video", "media", "oleobject", "package",
                    "control", "activex",
                )):
                    external_media.append(record)
                else:
                    external_other.append(record)
            elif relationship_kind(kind) in UNSAFE_RELATIONSHIP_KINDS:
                unsafe_internal_relationships.append(record)

    unsafe_content_types: list[dict[str, str]] = []
    for part in sorted(
        name for name in names
        if name != "[Content_Types].xml" and not name.endswith("/")
    ):
        content_type = package_content_type(part, defaults, overrides) or ""
        lowered = content_type.lower()
        if any(fragment in lowered for fragment in UNSAFE_CONTENT_TYPE_FRAGMENTS):
            unsafe_content_types.append({"part": part, "content_type": content_type})

    unsafe_xml_markers: list[dict[str, str]] = []
    for part in sorted(name for name in names if is_xml_part(name)):
        if part.endswith(".rels") or part == "[Content_Types].xml":
            continue
        root = parse_xml(archive, part)
        for element in root.iter():
            if element.tag.startswith("{"):
                namespace, name = element.tag[1:].split("}", 1)
            else:
                namespace, name = "", element.tag
            lowered_name = name.lower()
            lowered_namespace = namespace.lower()
            marked = (
                "activex" in lowered_namespace
                or lowered_name in {"oleobj", "oleobject"}
                or (
                    lowered_name in {"control", "controls", "ocx"}
                    and (
                        namespace == NS["p"]
                        or "microsoft.com/office" in lowered_namespace
                        or "schemas-microsoft-com:office" in lowered_namespace
                    )
                )
            )
            if marked:
                unsafe_xml_markers.append({
                    "part": part,
                    "tag": element.tag,
                })
                break
    return {
        "package_parts": package_parts,
        "unsafe_internal_relationships": unsafe_internal_relationships,
        "unsafe_content_types": unsafe_content_types,
        "unsafe_xml_markers": unsafe_xml_markers,
        "external_media": external_media,
        "external_other": external_other,
    }


def image_package_inventory(
    archive: zipfile.ZipFile,
) -> dict[str, list[dict[str, object]]]:
    """Index every image/* package part and any internal relationships to it."""
    names = set(archive.namelist())
    defaults, overrides = content_type_map(archive)
    result: dict[str, list[dict[str, object]]] = {
        part: []
        for part in sorted(names)
        if not part.endswith("/")
        and (package_content_type(part, defaults, overrides) or "").lower().startswith("image/")
    }
    for rel_part in sorted(name for name in names if name.endswith(".rels")):
        owner = relationship_owner(rel_part)
        if owner is None:
            continue
        root = parse_xml(archive, rel_part)
        for relation in root.findall(f"{{{REL_NS}}}Relationship"):
            if (relation.get("TargetMode") or "").lower() == "external":
                continue
            resolved = resolve_target(owner, relation.get("Target") or "")
            if not resolved or resolved not in result:
                continue
            content_type = package_content_type(resolved, defaults, overrides) or ""
            result[resolved].append({
                "owner": owner,
                "relationship_part": rel_part,
                "relationship_id": relation.get("Id"),
                "relationship_type": relation.get("Type"),
                "content_type": content_type,
            })
    return result


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


def static_module_specifiers(text: str) -> list[str]:
    """Extract static ESM specifiers, including conventional multiline imports."""
    declarations = list(MODULE_DECLARATION_PATTERN.finditer(text))
    imports: set[str] = set()
    for index, declaration in enumerate(declarations):
        kind = declaration.group(1)
        start = declaration.end()
        if kind == "import" and text[start:].lstrip().startswith("."):
            # `import.meta` is an expression, not a module declaration.
            continue
        next_start = (
            declarations[index + 1].start()
            if index + 1 < len(declarations)
            else len(text)
        )
        semicolon = text.find(";", start, next_start)
        end = semicolon + 1 if semicolon >= 0 else next_start
        statement = text[start:end]
        match = re.search(r"\bfrom\s*['\"]([^'\"]+)['\"]", statement, re.DOTALL)
        if match is None and kind == "import":
            match = re.match(r"\s*['\"]([^'\"]+)['\"]", statement, re.DOTALL)
        if match is not None:
            imports.add(match.group(1))
        elif kind == "import":
            raise CheckError("build source contains an unparseable import declaration")
    return sorted(imports)


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

    if DYNAMIC_IMPORT_PATTERN.search(text):
        raise CheckError("build source contains a computed or dynamic import")
    if REQUIRE_PATTERN.search(text):
        raise CheckError("build source contains require(), which is not portable here")
    if CREATE_REQUIRE_PATTERN.search(text):
        raise CheckError("build source contains createRequire, which can bypass import checks")

    imports = static_module_specifiers(text)
    unsupported = [
        specifier for specifier in imports
        if specifier not in BUILD_SOURCE_PACKAGES and specifier not in SAFE_NODE_MODULES
    ]
    if unsupported:
        raise CheckError(
            "build source imports an unshipped or unsupported helper: "
            + ", ".join(unsupported)
        )
    authoring_packages = sorted(BUILD_SOURCE_PACKAGES.intersection(imports))
    if len(authoring_packages) > 1:
        raise CheckError(
            "build source imports more than one authoring runtime: "
            + ", ".join(authoring_packages)
        )
    if not authoring_packages:
        raise CheckError("build source must import exactly one authoring runtime")
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
    slide_number: int = 1,
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
            "Slide 1 contains custom geometry; inspect the render for displaced or clipped labels",
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
            "Slide 1 uses wrap=none or an AutoFit mode that may reflow across presentation renderers",
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

    if slide_number != 1:
        for item in failures + warnings + passes:
            identifier = str(item.get("id", ""))
            if identifier.startswith("slide1."):
                item["id"] = f"slide{slide_number}." + identifier.removeprefix("slide1.")
            item["message"] = str(item.get("message", "")).replace(
                "Slide 1", f"Slide {slide_number}"
            )
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
    root: ET.Element,
    *,
    include_shapes: bool = True,
    include_background: bool = True,
) -> tuple[list[tuple[int, int, int, int]], list[float], int, int, int, int]:
    parent = {child: node for node in root.iter() for child in node}
    rectangles: list[tuple[int, int, int, int]] = []
    transformed_areas: list[float] = []
    grouped_count = 0
    unresolved_grouped = 0
    count = 0
    hidden_count = 0
    candidates = (
        root.findall(".//p:pic", NS)
        + root.findall(".//p:sp", NS)
        + root.findall(".//p:graphicFrame", NS)
        if include_shapes else []
    )
    for node in candidates:
        if node.find(".//a:blip", NS) is None:
            continue
        if image_object_hidden(node, parent):
            hidden_count += 1
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
    if include_background and root.find("./p:cSld/p:bg/p:bgPr/a:blipFill", NS) is not None:
        count += 1
    return (
        rectangles,
        transformed_areas,
        count,
        grouped_count,
        unresolved_grouped,
        hidden_count,
    )


def xml_boolean(value: str | None) -> bool:
    return (value or "").strip().lower() in {"1", "true", "on", "yes"}


def xml_boolean_default(value: str | None, default: bool = True) -> bool:
    if value is None:
        return default
    return xml_boolean(value)


def related_internal_part(
    archive: zipfile.ZipFile,
    owner: str,
    relationship_suffix: str,
) -> str | None:
    matches = [
        relation for relation in relationships(archive, owner).values()
        if not relation.get("external")
        and str(relation.get("type", "")).lower().endswith(relationship_suffix.lower())
    ]
    if len(matches) > 1:
        raise CheckError(
            f"{owner} contains multiple {relationship_suffix.rsplit('/', 1)[-1]} relationships"
        )
    if not matches:
        return None
    target = matches[0].get("resolved")
    return target if isinstance(target, str) else None


def visible_image_layers(
    archive: zipfile.ZipFile,
    slide_part: str,
    slide: ET.Element,
) -> list[dict[str, object]]:
    """Resolve the selected slide's visible slide/layout/master image layers."""
    layers: list[dict[str, object]] = [
        {"kind": "slide", "part": slide_part, "root": slide, "include_shapes": True}
    ]
    layout_part = related_internal_part(archive, slide_part, "/slideLayout")
    layout: ET.Element | None = None
    if layout_part is not None:
        layout = parse_xml(archive, layout_part)
        layers.append({
            "kind": "layout",
            "part": layout_part,
            "root": layout,
            "include_shapes": True,
        })
    master_part = (
        related_internal_part(archive, layout_part, "/slideMaster")
        if layout_part is not None else None
    )
    if master_part is not None:
        master = parse_xml(archive, master_part)
        show_master_shapes = (
            xml_boolean_default(slide.get("showMasterSp"))
            and xml_boolean_default(layout.get("showMasterSp") if layout is not None else None)
        )
        layers.append({
            "kind": "master",
            "part": master_part,
            "root": master,
            "include_shapes": show_master_shapes,
        })

    background_part: str | None = None
    for layer in layers:
        root = layer["root"]
        if (
            isinstance(root, ET.Element)
            and root.find("./p:cSld/p:bg/p:bgPr/a:blipFill", NS) is not None
        ):
            background_part = str(layer["part"])
            break
    for layer in layers:
        layer["include_background"] = layer["part"] == background_part
    return layers


def non_visual_properties(node: ET.Element) -> ET.Element | None:
    for path in (
        "./p:nvPicPr/p:cNvPr",
        "./p:nvSpPr/p:cNvPr",
        "./p:nvCxnSpPr/p:cNvPr",
        "./p:nvGraphicFramePr/p:cNvPr",
        "./p:nvGrpSpPr/p:cNvPr",
    ):
        properties = node.find(path, NS)
        if properties is not None:
            return properties
    return None


def object_marked_hidden(
    node: ET.Element,
    parents: dict[ET.Element, ET.Element],
) -> bool:
    current: ET.Element | None = node
    while current is not None:
        properties = non_visual_properties(current)
        if properties is not None and xml_boolean(properties.get("hidden")):
            return True
        current = parents.get(current)
    return False


def image_object_hidden(
    node: ET.Element,
    parents: dict[ET.Element, ET.Element],
) -> bool:
    if object_marked_hidden(node, parents):
        return True

    for effect_name, attribute in (
        ("alpha", "val"),
        ("alphaMod", "amt"),
        ("alphaModFix", "amt"),
    ):
        for effect in node.findall(f".//a:{effect_name}", NS):
            if integer(effect, attribute) == 0:
                return True
    return False


def segment_intersects_slide(
    start: tuple[float, float],
    end: tuple[float, float],
    slide_width: int,
    slide_height: int,
) -> bool:
    """Return whether a closed line segment intersects the slide rectangle."""
    x1, y1 = start
    x2, y2 = end
    delta_x, delta_y = x2 - x1, y2 - y1
    lower, upper = 0.0, 1.0
    for direction, distance in (
        (-delta_x, x1),
        (delta_x, slide_width - x1),
        (-delta_y, y1),
        (delta_y, slide_height - y1),
    ):
        if direction == 0:
            if distance < 0:
                return False
            continue
        ratio = distance / direction
        if direction < 0:
            lower = max(lower, ratio)
        else:
            upper = min(upper, ratio)
        if lower > upper:
            return False
    return True


def native_object_inventory(
    slide: ET.Element,
    slide_width: int,
    slide_height: int,
) -> dict[str, object]:
    """Count only visible, resolved native objects that intersect the slide."""
    parents = {child: node for node in slide.iter() for child in node}
    shapes = [
        node for node in slide.findall(".//p:sp", NS)
        if node.find(".//a:blip", NS) is None
    ]
    connectors = list(slide.findall(".//p:cxnSp", NS))
    excluded = {
        "hidden": 0,
        "unreadable_geometry": 0,
        "zero_dimensions": 0,
        "outside_slide": 0,
    }
    visible_shapes = 0
    visible_connectors = 0

    for shape in shapes:
        if object_marked_hidden(shape, parents):
            excluded["hidden"] += 1
            continue
        raw_box = object_box(shape)
        if raw_box is None:
            excluded["unreadable_geometry"] += 1
            continue
        if raw_box[2] <= 0 or raw_box[3] <= 0:
            excluded["zero_dimensions"] += 1
            continue
        geometry = image_box(shape, parents)
        if geometry is None:
            excluded["unreadable_geometry"] += 1
            continue
        if clip_box_to_slide(geometry[0], slide_width, slide_height) is None:
            excluded["outside_slide"] += 1
            continue
        visible_shapes += 1

    for connector in connectors:
        if object_marked_hidden(connector, parents):
            excluded["hidden"] += 1
            continue
        raw_box = object_box(connector)
        if raw_box is None:
            excluded["unreadable_geometry"] += 1
            continue
        if raw_box[2] < 0 or raw_box[3] < 0 or (raw_box[2] == 0 and raw_box[3] == 0):
            excluded["zero_dimensions"] += 1
            continue
        geometry = image_box(connector, parents)
        if geometry is None:
            excluded["unreadable_geometry"] += 1
            continue
        x, y, width, height = geometry[0]
        transform = connector.find("./p:spPr/a:xfrm", NS)
        flipped = (
            xml_boolean(transform.get("flipH") if transform is not None else None)
            != xml_boolean(transform.get("flipV") if transform is not None else None)
        )
        start = (x + width, y) if flipped else (x, y)
        end = (x, y + height) if flipped else (x + width, y + height)
        if not segment_intersects_slide(start, end, slide_width, slide_height):
            excluded["outside_slide"] += 1
            continue
        visible_connectors += 1

    return {
        "native_shapes": visible_shapes,
        "connectors": visible_connectors,
        "candidate_native_shapes": len(shapes),
        "candidate_connectors": len(connectors),
        "excluded_objects": excluded,
    }


def crop_rectangle(node: ET.Element) -> tuple[dict[str, int | None], bool, bool]:
    values: dict[str, int | None] = {}
    for edge in ("l", "t", "r", "b"):
        raw = node.get(edge)
        if raw is None:
            values[edge] = 0
            continue
        try:
            values[edge] = int(raw)
        except ValueError:
            values[edge] = None
    parsed = all(value is not None for value in values.values())
    numeric = [int(value) for value in values.values() if value is not None]
    valid = (
        parsed
        and all(0 <= value <= 100_000 for value in numeric)
        and int(values["l"] or 0) + int(values["r"] or 0) < 100_000
        and int(values["t"] or 0) + int(values["b"] or 0) < 100_000
    )
    nonzero = parsed and any(value != 0 for value in numeric)
    return values, valid, nonzero


def source_raster_inset_inventory(
    archive: zipfile.ZipFile,
    slide_part: str,
    slide: ET.Element,
    source_hash: str,
    slide_width: int,
    slide_height: int,
    slide_number: int = 1,
) -> tuple[dict[str, object], list[str]]:
    """Inventory exact source bytes used as visible OOXML-cropped raster insets."""
    image_parts = image_package_inventory(archive)
    matching_media = sorted(
        part for part in image_parts
        if sha256_bytes(archive.read(part)) == source_hash
    )
    matching_media_set = set(matching_media)
    source_relationships = [
        {"image_part": part, **record}
        for part in matching_media
        for record in image_parts[part]
    ]
    non_selected_relationships = [
        record for record in source_relationships
        if record.get("owner") != slide_part
    ]
    slide_rels = relationships(archive, slide_part)
    parents = {child: node for node in slide.iter() for child in node}
    image_object_tags = {
        f"{{{NS['p']}}}pic",
        f"{{{NS['p']}}}sp",
        f"{{{NS['p']}}}graphicFrame",
    }
    image_objects = [node for node in slide.iter() if node.tag in image_object_tags]
    object_indexes = {node: index for index, node in enumerate(image_objects, start=1)}
    records: list[dict[str, object]] = []
    visible_rectangles: list[tuple[int, int, int, int]] = []
    visible_coverages: list[float] = []
    referenced_media: set[str] = set()
    source_object_keys: set[tuple[str, int]] = set()
    slide_area = slide_width * slide_height

    for blip in slide.findall(".//a:blip", NS):
        relation_id = blip.get(R_EMBED) or blip.get(R_LINK)
        relation = slide_rels.get(relation_id or "")
        media_part = relation.get("resolved") if relation else None
        if not isinstance(media_part, str) or media_part not in matching_media_set:
            continue
        referenced_media.add(media_part)

        owner: ET.Element | None = parents.get(blip)
        while owner is not None and owner.tag not in image_object_tags:
            owner = parents.get(owner)
        object_index = object_indexes.get(owner) if owner is not None else None
        object_type = local_name(owner.tag) if owner is not None else "unsupported"
        object_key = (object_type, object_index or -len(records) - 1)
        source_object_keys.add(object_key)
        properties = non_visual_properties(owner) if owner is not None else None
        hidden = owner is None or image_object_hidden(owner, parents)
        reasons: list[str] = []

        crop_nodes = owner.findall(".//a:srcRect", NS) if owner is not None else []
        crop_records: list[dict[str, int | None]] = []
        crop_valid = False
        crop_nonzero = False
        if len(crop_nodes) == 1:
            crop_values, crop_valid, crop_nonzero = crop_rectangle(crop_nodes[0])
            crop_records.append(crop_values)
            if not crop_valid:
                reasons.append("invalid_src_rect_crop")
            elif not crop_nonzero:
                reasons.append("missing_src_rect_crop")
        elif not crop_nodes:
            reasons.append("missing_src_rect_crop")
        else:
            for crop_node in crop_nodes:
                crop_records.append(crop_rectangle(crop_node)[0])
            reasons.append("ambiguous_src_rect_crop")

        if hidden:
            reasons.append("hidden_source_inset")

        box: tuple[int, int, int, int] | None = None
        clipped: tuple[int, int, int, int] | None = None
        visible_area = 0.0
        coverage = 0.0
        raw_box = object_box(owner) if owner is not None else None
        if raw_box is None:
            reasons.append("unresolved_geometry")
        elif raw_box[2] <= 0 or raw_box[3] <= 0:
            reasons.append("zero_dimensions")
        else:
            geometry = image_box(owner, parents) if owner is not None else None
            if geometry is None:
                reasons.append("unresolved_geometry")
            else:
                box, transformed_area = geometry
                clipped = clip_box_to_slide(box, slide_width, slide_height)
                if clipped is None:
                    reasons.append("outside_slide")
                else:
                    visible_area = min(transformed_area, clipped[2] * clipped[3])
                    if visible_area <= 0:
                        reasons.append("zero_visible_area")
                    else:
                        coverage = visible_area / slide_area
                        if not hidden:
                            visible_rectangles.append(clipped)
                            visible_coverages.append(coverage)

        record: dict[str, object] = {
            "usage_index": len(records) + 1,
            "object_index": object_index,
            "object_type": object_type,
            "object_id": properties.get("id") if properties is not None else None,
            "name": properties.get("name") if properties is not None else None,
            "relationship_id": relation_id,
            "media_part": media_part,
            "src_rects": crop_records,
            "src_rect_valid": crop_valid,
            "src_rect_nonzero": crop_nonzero,
            "hidden": hidden,
            "box_emu": list(box) if box is not None else None,
            "visible_box_emu": list(clipped) if clipped is not None else None,
            "visible_area_emu2": int(round(visible_area)),
            "visible_coverage": round(coverage, 4),
            "qualifies_individually": not reasons,
            "violations": reasons,
        }
        records.append(record)

    unused_media = sorted(matching_media_set - referenced_media)
    union_coverage = (
        union_area(visible_rectangles) / slide_area if visible_rectangles else 0.0
    )
    violations = sorted({
        str(reason)
        for record in records
        for reason in record["violations"]  # type: ignore[union-attr]
    })
    if unused_media:
        violations.append("unreferenced_source_media")
    if non_selected_relationships:
        violations.append("source_used_outside_inspected_slide")
    violations = sorted(set(violations))

    evidence: dict[str, object] = {
        "source_sha256": source_hash,
        "matching_media_parts": matching_media,
        "matching_image_parts": matching_media,
        "unused_media_parts": unused_media,
        "source_image_relationships": source_relationships,
        "non_selected_source_relationships": non_selected_relationships,
        "source_media_part_count": len(matching_media),
        "source_usage_count": len(records),
        "source_object_count": len(source_object_keys),
        "qualifying_object_count": sum(
            bool(record["qualifies_individually"]) for record in records
        ),
        "largest_visible_coverage": round(max(visible_coverages, default=0.0), 4),
        "union_visible_coverage": round(union_coverage, 4),
        "policy": {
            "requires_nonzero_src_rect": True,
            "requires_visible_positive_dimensions": True,
            "whole_page_or_mosaic_check_id": f"slide{slide_number}.not_flattened",
        },
        "objects": records[:25],
        "truncated": max(0, len(records) - 25),
        "violations": violations,
    }
    return evidence, violations


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
    inspect_slide: int = 1,
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
        if isinstance(inspect_slide, bool) or not isinstance(inspect_slide, int) or inspect_slide < 1:
            raise CheckError(
                "inspect_slide must be a positive 1-based slide number",
                identifier="inspection.slide_selection",
                requested_slide=inspect_slide,
            )
        if not pptx.is_file():
            raise CheckError(f"PPTX does not exist: {pptx}")
        if source and not source.is_file():
            raise CheckError(f"source image does not exist: {source}")
        if not zipfile.is_zipfile(pptx):
            raise CheckError("file is not a readable ZIP/PPTX package")

        with zipfile.ZipFile(pptx) as archive:
            zip_evidence = validate_zip_budget(archive)
            checks.append(check_item(
                "package.unique_members", "PASS",
                "ZIP package member names are unique",
                member_count=zip_evidence["member_count"],
            ))
            checks.append(check_item(
                "package.zip_budget", "PASS",
                "ZIP package is within bounded member, size, XML, and compression budgets",
                **zip_evidence,
            ))
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
            if inspect_slide > len(slide_parts):
                raise CheckError(
                    "requested slide is outside the presentation slide range",
                    identifier="inspection.slide_selection",
                    requested_slide=inspect_slide,
                    slide_count=len(slide_parts),
                )
            for slide_part in slide_parts:
                slide = parse_xml(archive, slide_part)
                validate_slide_blips(archive, slide_part, slide)
            checks.append(check_item(
                "package.readable", "PASS",
                "PPTX content types, internal relationships, and slide order are readable",
                slide_count=len(slide_parts), slide_size_emu=[slide_width, slide_height],
            ))
            selected_slide_part = slide_parts[inspect_slide - 1]
            selected_slide = parse_xml(archive, selected_slide_part)
            slide_prefix = f"slide{inspect_slide}"
            slide_label = f"Slide {inspect_slide}"
            checks.append(check_item(
                "inspection.slide_selection", "PASS",
                f"{slide_label} was selected for deep structural checks",
                requested_slide=inspect_slide,
                selected_part=selected_slide_part,
                slide_count=len(slide_parts),
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

            unsafe = unsafe_embed_inventory(archive)
            unsafe_evidence = {
                key: unsafe[key]
                for key in (
                    "package_parts",
                    "unsafe_internal_relationships",
                    "unsafe_content_types",
                    "unsafe_xml_markers",
                    "external_media",
                )
            }
            if any(unsafe_evidence.values()):
                hard_failures.append(check_item(
                    "package.no_unsafe_embeds", "FAIL",
                    "Package contains a macro, OLE/package object, ActiveX/control marker, or external media",
                    **unsafe_evidence,
                ))
            else:
                checks.append(check_item(
                    "package.no_unsafe_embeds", "PASS",
                    "No macro, OLE/package object, ActiveX/control marker, or external media was found",
                ))
            if unsafe["external_other"]:
                warnings.append(check_item(
                    "package.external_relationships", "WARN",
                    "Package contains external relationships such as hyperlinks",
                    relationships=unsafe["external_other"],
                ))

            native_inventory = native_object_inventory(
                selected_slide, slide_width, slide_height
            )
            native_shapes = int(native_inventory["native_shapes"])
            connectors = int(native_inventory["connectors"])
            text_runs = [
                node.text or "" for node in selected_slide.findall(".//a:t", NS)
                if (node.text or "").strip()
            ]
            if native_shapes == 0 and connectors == 0:
                hard_failures.append(check_item(
                    f"{slide_prefix}.native_objects", "FAIL",
                    f"{slide_label} contains no visible native shape or connector with readable on-slide geometry",
                    **native_inventory,
                    text_runs=len(text_runs),
                ))
            else:
                checks.append(check_item(
                    f"{slide_prefix}.native_objects", "PASS",
                    f"{slide_label} contains visible editable native objects with readable on-slide geometry",
                    **native_inventory,
                    text_runs=len(text_runs),
                ))

            rectangles: list[tuple[int, int, int, int]] = []
            transformed_areas: list[float] = []
            raster_count = 0
            grouped_raster_count = 0
            unresolved_grouped_raster = 0
            hidden_raster_count = 0
            has_image_background = False
            layer_evidence: list[dict[str, object]] = []
            for layer in visible_image_layers(archive, selected_slide_part, selected_slide):
                layer_root = layer["root"]
                layer_part = str(layer["part"])
                if not isinstance(layer_root, ET.Element):
                    continue
                if bool(layer["include_shapes"]) or bool(layer["include_background"]):
                    validate_slide_blips(archive, layer_part, layer_root)
                (
                    layer_rectangles,
                    layer_areas,
                    layer_rasters,
                    layer_grouped,
                    layer_unresolved,
                    layer_hidden,
                ) = image_bearers(
                    layer_root,
                    include_shapes=bool(layer["include_shapes"]),
                    include_background=bool(layer["include_background"]),
                )
                rectangles.extend(layer_rectangles)
                transformed_areas.extend(layer_areas)
                raster_count += layer_rasters
                grouped_raster_count += layer_grouped
                unresolved_grouped_raster += layer_unresolved
                hidden_raster_count += layer_hidden
                layer_background = bool(layer["include_background"]) and (
                    layer_root.find("./p:cSld/p:bg/p:bgPr/a:blipFill", NS) is not None
                )
                has_image_background = has_image_background or layer_background
                layer_evidence.append({
                    "kind": layer["kind"],
                    "part": layer_part,
                    "shapes_visible": bool(layer["include_shapes"]),
                    "image_background_visible": layer_background,
                    "raster_objects": layer_rasters,
                    "grouped_raster_objects": layer_grouped,
                    "unresolved_grouped_raster_objects": layer_unresolved,
                    "hidden_raster_objects": layer_hidden,
                })

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
            raster_simulation = has_image_background or max_ratio >= 0.85 or (raster_count >= 2 and union_ratio >= 0.80)
            if raster_simulation:
                hard_failures.append(check_item(
                    f"{slide_prefix}.not_flattened", "FAIL",
                    f"{slide_label} appears to use a whole-page image or image mosaic across its visible slide/layout/master layers",
                    raster_objects=raster_count,
                    grouped_raster_objects=grouped_raster_count,
                    largest_coverage=round(max_ratio, 4),
                    union_coverage=round(union_ratio, 4),
                    image_background=has_image_background,
                    layers=layer_evidence,
                ))
            else:
                checks.append(check_item(
                    f"{slide_prefix}.not_flattened", "PASS",
                    f"No whole-page image or large image mosaic was detected on {slide_label} or its visible inherited layers",
                    raster_objects=raster_count,
                    grouped_raster_objects=grouped_raster_count,
                    largest_coverage=round(max_ratio, 4),
                    union_coverage=round(union_ratio, 4),
                    hidden_raster_objects=hidden_raster_count,
                    layers=layer_evidence,
                ))
            if unresolved_grouped_raster:
                warnings.append(check_item(
                    f"{slide_prefix}.grouped_raster_geometry", "WARN",
                    f"A visible grouped raster object on {slide_label} or an inherited layer has an incomplete transform and needs visual confirmation",
                    unresolved_objects=unresolved_grouped_raster,
                    layers=layer_evidence,
                ))

            compatibility_failures, compatibility_warnings, compatibility_passes = (
                slide1_compatibility_checks(selected_slide, inspect_slide)
            )
            hard_failures.extend(compatibility_failures)
            warnings.extend(compatibility_warnings)
            checks.extend(compatibility_passes)

            if require_single_slide and source:
                source_hash = sha256_file(source)
                inset_evidence, inset_violations = source_raster_inset_inventory(
                    archive,
                    selected_slide_part,
                    selected_slide,
                    source_hash,
                    slide_width,
                    slide_height,
                    inspect_slide,
                )
                if inset_violations:
                    warnings.append(check_item(
                        "source.raster_inset_inventory", "WARN",
                        "Exact source bytes are embedded, but one or more uses do not satisfy the local raster-inset policy",
                        **inset_evidence,
                    ))
                    hard_failures.append(check_item(
                        "source.not_embedded", "FAIL",
                        "Source bytes may be embedded only as explicitly cropped, visible raster insets; whole-page and mosaic use is checked separately",
                        inventory_check_id="source.raster_inset_inventory",
                        matching_media=inset_evidence["matching_media_parts"],
                        source_usage_count=inset_evidence["source_usage_count"],
                        violations=inset_violations,
                        sha256=source_hash,
                    ))
                else:
                    source_usage_count = int(inset_evidence["source_usage_count"])
                    checks.append(check_item(
                        "source.raster_inset_inventory", "PASS",
                        (
                            "Every embedded copy of the source is a nonzero-cropped, visible raster inset"
                            if source_usage_count
                            else "No exact source bytes are embedded as raster media"
                        ),
                        **inset_evidence,
                    ))
                    checks.append(check_item(
                        "source.external_companion", "PASS",
                        (
                            "The supplied source remains an external companion; embedded copies are limited to qualifying cropped insets"
                            if source_usage_count
                            else "The supplied source image is external and no identical media bytes are embedded in the PPTX"
                        ),
                        path=str(source),
                        bytes=source.stat().st_size,
                        sha256=source_hash,
                        qualifying_source_insets=source_usage_count,
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

    except CheckError as exc:
        hard_failures.append(check_item(
            exc.identifier, "FAIL", str(exc), **exc.evidence
        ))
    except (OSError, RuntimeError, NotImplementedError, EOFError, zipfile.BadZipFile) as exc:
        hard_failures.append(check_item("package.readable", "FAIL", str(exc)))

    return {
        "kind": "sci-diagram-pptx-check",
        "status": "FAIL" if hard_failures else "PASS",
        "pptx": str(pptx),
        "source": str(source) if source else None,
        "build_source": str(build_source) if build_source else None,
        "inspect_slide": inspect_slide,
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


def positive_slide_number(value: str) -> int:
    try:
        number = int(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError("slide must be a positive integer") from exc
    if number < 1:
        raise argparse.ArgumentTypeError("slide must be a positive integer")
    return number


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
    parser.add_argument(
        "--slide",
        type=positive_slide_number,
        default=1,
        metavar="N",
        help="1-based slide number to deep-inspect (default: 1)",
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
        inspect_slide=args.slide,
    )
    write_report(report, args.output)
    return 0 if report["status"] == "PASS" else 1


if __name__ == "__main__":
    raise SystemExit(main())
