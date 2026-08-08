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
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}
R_ID = f"{{{NS['r']}}}id"
R_EMBED = f"{{{NS['r']}}}embed"
REL_NS = NS["rel"]


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


def inspect_package(pptx: Path, source: Path | None = None) -> dict[str, object]:
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

            if source and len(slide_parts) >= 2:
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

    checks.extend(item for item in hard_failures if item not in checks)
    checks.extend(item for item in warnings if item not in checks)
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
    parser.add_argument("--source", type=Path, help="optional source image for Slide 2 identity check")
    parser.add_argument("--output", type=Path, help="optional JSON report path")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    report = inspect_package(args.pptx, args.source)
    write_report(report, args.output)
    return 0 if report["status"] == "PASS" else 1


if __name__ == "__main__":
    raise SystemExit(main())
