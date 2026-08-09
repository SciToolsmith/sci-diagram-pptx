#!/usr/bin/env python3
"""Run a focused structural check on an editable scientific-diagram PPTX."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import posixpath
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
}
R_ID = f"{{{NS['r']}}}id"
R_EMBED = f"{{{NS['r']}}}embed"
REL_NS = NS["rel"]
FONT_SCRIPTS = ("latin", "ea", "cs")
AUTOFIT_NAMES = {"noAutofit", "normAutofit", "spAutoFit"}


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


def image_bearers(slide: ET.Element) -> tuple[list[tuple[int, int, int, int]], int, bool]:
    parent = {child: node for node in slide.iter() for child in node}
    rectangles: list[tuple[int, int, int, int]] = []
    uncertain_grouped = False
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
            uncertain_grouped = True
            continue
        box = object_box(node)
        if box:
            rectangles.append(box)
    if slide.find("./p:cSld/p:bg/p:bgPr/a:blipFill", NS) is not None:
        count += 1
    return rectangles, count, uncertain_grouped


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
) -> dict[str, object]:
    hard_failures: list[dict[str, object]] = []
    warnings: list[dict[str, object]] = []
    checks: list[dict[str, object]] = []
    pptx = pptx.expanduser().resolve()
    source = source.expanduser().resolve() if source else None

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
            slide_parts, (slide_width, slide_height) = slide_order(archive)
            checks.append(check_item(
                "package.readable", "PASS", "PPTX package and slide order are readable",
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
            rectangles, raster_count, grouped_raster = image_bearers(slide1)
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
            max_ratio = max((width * height / slide_area for _, _, width, height in rectangles), default=0.0)
            union_ratio = union_area(rectangles) / slide_area if rectangles else 0.0
            has_image_background = slide1.find("./p:cSld/p:bg/p:bgPr/a:blipFill", NS) is not None
            raster_simulation = has_image_background or max_ratio >= 0.85 or (raster_count >= 2 and union_ratio >= 0.80)
            if raster_simulation:
                hard_failures.append(check_item(
                    "slide1.not_flattened", "FAIL",
                    "Slide 1 appears to use a whole-page image or image mosaic",
                    raster_objects=raster_count,
                    largest_coverage=round(max_ratio, 4),
                    union_coverage=round(union_ratio, 4),
                    image_background=has_image_background,
                ))
            else:
                checks.append(check_item(
                    "slide1.not_flattened", "PASS",
                    "No whole-page image or large image mosaic was detected",
                    raster_objects=raster_count,
                    largest_coverage=round(max_ratio, 4),
                    union_coverage=round(union_ratio, 4),
                ))
            if grouped_raster:
                warnings.append(check_item(
                    "slide1.grouped_raster_geometry", "WARN",
                    "A raster object inside a group needs visual confirmation",
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
    )
    write_report(report, args.output)
    return 0 if report["status"] == "PASS" else 1


if __name__ == "__main__":
    raise SystemExit(main())
