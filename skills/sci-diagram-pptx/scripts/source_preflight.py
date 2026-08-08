#!/usr/bin/env python3
"""Inspect a source image before rebuilding it as editable slide content.

The script intentionally uses only the Python standard library.  It reads image
metadata; it never decodes pixels, renders an image, or creates a presentation.
Its JSON output is also the source manifest consumed by validate_scene_plan.py.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import mimetypes
import os
import struct
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable


MANIFEST_VERSION = "1.0"
ORIENTATIONS: dict[int, dict[str, Any]] = {
    1: {"name": "normal", "rotation_degrees_clockwise": 0, "mirrored": False, "swaps_dimensions": False},
    2: {"name": "mirror-horizontal", "rotation_degrees_clockwise": 0, "mirrored": True, "swaps_dimensions": False},
    3: {"name": "rotate-180", "rotation_degrees_clockwise": 180, "mirrored": False, "swaps_dimensions": False},
    4: {"name": "mirror-vertical", "rotation_degrees_clockwise": 180, "mirrored": True, "swaps_dimensions": False},
    5: {"name": "transpose", "rotation_degrees_clockwise": 90, "mirrored": True, "swaps_dimensions": True},
    6: {"name": "rotate-90-cw", "rotation_degrees_clockwise": 90, "mirrored": False, "swaps_dimensions": True},
    7: {"name": "transverse", "rotation_degrees_clockwise": 270, "mirrored": True, "swaps_dimensions": True},
    8: {"name": "rotate-270-cw", "rotation_degrees_clockwise": 270, "mirrored": False, "swaps_dimensions": True},
}

TIFF_TYPES: dict[int, tuple[str, int]] = {
    1: ("B", 1),   # BYTE
    2: ("s", 1),   # ASCII
    3: ("H", 2),   # SHORT
    4: ("I", 4),   # LONG
    5: ("II", 8),  # RATIONAL
    7: ("s", 1),   # UNDEFINED
    9: ("i", 4),   # SLONG
    10: ("ii", 8), # SRATIONAL
}

EXIF_TAG_NAMES = {
    0x010F: "make",
    0x0110: "model",
    0x0112: "orientation",
    0x0131: "software",
    0x0132: "datetime",
    0x013B: "artist",
    0x8298: "copyright",
    0x9003: "datetime_original",
    0x9004: "datetime_digitized",
    0xA002: "pixel_width",
    0xA003: "pixel_height",
    0xA434: "lens_model",
}


class ImageParseError(ValueError):
    """Raised when the source is unsupported or structurally invalid."""


def _u16(data: bytes, offset: int, endian: str = ">") -> int:
    if offset < 0 or offset + 2 > len(data):
        raise ImageParseError("Unexpected end of file while reading a 16-bit value")
    return struct.unpack_from(endian + "H", data, offset)[0]


def _u32(data: bytes, offset: int, endian: str = ">") -> int:
    if offset < 0 or offset + 4 > len(data):
        raise ImageParseError("Unexpected end of file while reading a 32-bit value")
    return struct.unpack_from(endian + "I", data, offset)[0]


def _first(value: Any) -> Any:
    if isinstance(value, list) and value:
        return value[0]
    return value


def _decode_tiff_value(raw: bytes, field_type: int, count: int, endian: str) -> Any:
    if count < 0 or count > 1_000_000:
        raise ImageParseError("Unreasonable TIFF field element count")
    if field_type == 2:
        return raw[:count].split(b"\0", 1)[0].decode("utf-8", "replace")
    if field_type == 7:
        return {"byte_count": min(count, len(raw))}
    if field_type == 1:
        values = list(raw[:count])
    elif field_type in (3, 4, 9):
        fmt = {3: "H", 4: "I", 9: "i"}[field_type]
        size = struct.calcsize(fmt)
        values = [struct.unpack_from(endian + fmt, raw, i * size)[0] for i in range(count)]
    elif field_type in (5, 10):
        fmt = "II" if field_type == 5 else "ii"
        values = []
        for i in range(count):
            numerator, denominator = struct.unpack_from(endian + fmt, raw, i * 8)
            values.append(None if denominator == 0 else numerator / denominator)
    else:
        return None
    return values[0] if len(values) == 1 else values


def _parse_ifd(tiff: bytes, offset: int, endian: str) -> tuple[dict[int, Any], int | None]:
    if offset <= 0 or offset + 2 > len(tiff):
        raise ImageParseError("Invalid TIFF IFD offset")
    count = _u16(tiff, offset, endian)
    if count > 4096:
        raise ImageParseError("Unreasonable TIFF IFD entry count")
    end = offset + 2 + count * 12
    if end + 4 > len(tiff):
        raise ImageParseError("Truncated TIFF IFD")
    tags: dict[int, Any] = {}
    for index in range(count):
        entry = offset + 2 + index * 12
        tag = _u16(tiff, entry, endian)
        field_type = _u16(tiff, entry + 2, endian)
        element_count = _u32(tiff, entry + 4, endian)
        type_info = TIFF_TYPES.get(field_type)
        if type_info is None:
            continue
        byte_count = type_info[1] * element_count
        if byte_count > 64 * 1024 * 1024:
            raise ImageParseError("Unreasonable TIFF field size")
        if byte_count <= 4:
            raw = tiff[entry + 8 : entry + 8 + byte_count]
        else:
            value_offset = _u32(tiff, entry + 8, endian)
            if value_offset + byte_count > len(tiff):
                continue
            raw = tiff[value_offset : value_offset + byte_count]
        try:
            tags[tag] = _decode_tiff_value(raw, field_type, element_count, endian)
        except (struct.error, ImageParseError):
            continue
    next_ifd = _u32(tiff, end, endian)
    return tags, next_ifd or None


def _parse_exif(payload: bytes) -> tuple[dict[str, Any], list[str]]:
    """Return a privacy-conscious EXIF summary and non-fatal parse warnings."""
    warnings: list[str] = []
    if payload.startswith(b"Exif\x00\x00"):
        tiff = payload[6:]
    else:
        tiff = payload
    summary: dict[str, Any] = {
        "present": True,
        "parsed": False,
        "byte_order": None,
        "ifd0_tag_count": 0,
        "exif_tag_count": 0,
        "gps_present": False,
        "tags": {},
    }
    try:
        if len(tiff) < 8 or tiff[:2] not in (b"II", b"MM"):
            raise ImageParseError("EXIF payload is not a TIFF stream")
        endian = "<" if tiff[:2] == b"II" else ">"
        summary["byte_order"] = "little" if endian == "<" else "big"
        if _u16(tiff, 2, endian) != 42:
            raise ImageParseError("Invalid TIFF magic in EXIF payload")
        ifd0, _ = _parse_ifd(tiff, _u32(tiff, 4, endian), endian)
        summary["ifd0_tag_count"] = len(ifd0)
        exif_ifd: dict[int, Any] = {}
        exif_offset = _first(ifd0.get(0x8769))
        if isinstance(exif_offset, int) and exif_offset:
            exif_ifd, _ = _parse_ifd(tiff, exif_offset, endian)
        summary["exif_tag_count"] = len(exif_ifd)
        summary["gps_present"] = isinstance(_first(ifd0.get(0x8825)), int)
        selected: dict[str, Any] = {}
        for tag, name in EXIF_TAG_NAMES.items():
            value = ifd0.get(tag, exif_ifd.get(tag))
            if value is not None and value != "":
                selected[name] = value
        summary["tags"] = selected
        summary["parsed"] = True
    except (ImageParseError, struct.error) as exc:
        warnings.append(str(exc))
    return summary, warnings


def _empty_exif() -> dict[str, Any]:
    return {
        "present": False,
        "parsed": False,
        "byte_order": None,
        "ifd0_tag_count": 0,
        "exif_tag_count": 0,
        "gps_present": False,
        "tags": {},
    }


def _parse_png(data: bytes) -> dict[str, Any]:
    if len(data) < 33 or data[:8] != b"\x89PNG\r\n\x1a\n":
        raise ImageParseError("Invalid or truncated PNG signature/header")
    if data[12:16] != b"IHDR":
        raise ImageParseError("PNG does not start with an IHDR chunk")
    width, height = struct.unpack_from(">II", data, 16)
    bit_depth, color_type = data[24], data[25]
    exif = _empty_exif()
    parse_warnings: list[str] = []
    has_transparency_chunk = False
    offset = 8
    while offset + 12 <= len(data):
        length = _u32(data, offset)
        chunk_type = data[offset + 4 : offset + 8]
        chunk_start = offset + 8
        chunk_end = chunk_start + length
        if chunk_end + 4 > len(data):
            parse_warnings.append("PNG contains a truncated chunk")
            break
        chunk = data[chunk_start:chunk_end]
        if chunk_type == b"tRNS":
            has_transparency_chunk = True
        elif chunk_type == b"eXIf":
            exif, exif_warnings = _parse_exif(chunk)
            parse_warnings.extend(exif_warnings)
        offset = chunk_end + 4
        if chunk_type == b"IEND":
            break
    color_models = {0: "grayscale", 2: "rgb", 3: "indexed", 4: "grayscale-alpha", 6: "rgba"}
    alpha_present = color_type in (4, 6) or has_transparency_chunk
    return {
        "format": "PNG",
        "mime_type": "image/png",
        "width_px": width,
        "height_px": height,
        "bit_depth": bit_depth,
        "color_model": color_models.get(color_type, f"png-color-type-{color_type}"),
        "alpha": {
            "present": alpha_present,
            "channel_available": color_type in (4, 6),
            "detection": "color-type-or-tRNS",
        },
        "exif": exif,
        "parse_warnings": parse_warnings,
    }


JPEG_SOF_MARKERS = {0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF}


def _parse_jpeg(data: bytes) -> dict[str, Any]:
    if len(data) < 4 or data[:2] != b"\xff\xd8":
        raise ImageParseError("Invalid or truncated JPEG signature")
    width = height = bit_depth = components = None
    exif = _empty_exif()
    parse_warnings: list[str] = []
    offset = 2
    while offset < len(data):
        while offset < len(data) and data[offset] != 0xFF:
            offset += 1
        while offset < len(data) and data[offset] == 0xFF:
            offset += 1
        if offset >= len(data):
            break
        marker = data[offset]
        offset += 1
        if marker in (0x00, 0x01, 0xD8) or 0xD0 <= marker <= 0xD7:
            continue
        if marker in (0xD9, 0xDA):
            break
        if offset + 2 > len(data):
            parse_warnings.append("JPEG contains a truncated marker length")
            break
        segment_length = _u16(data, offset)
        if segment_length < 2 or offset + segment_length > len(data):
            parse_warnings.append("JPEG contains an invalid or truncated segment")
            break
        payload = data[offset + 2 : offset + segment_length]
        if marker == 0xE1 and payload.startswith(b"Exif\x00\x00"):
            exif, exif_warnings = _parse_exif(payload)
            parse_warnings.extend(exif_warnings)
        elif marker in JPEG_SOF_MARKERS and len(payload) >= 6:
            bit_depth = payload[0]
            height = _u16(payload, 1)
            width = _u16(payload, 3)
            components = payload[5]
        offset += segment_length
    if not width or not height:
        raise ImageParseError("JPEG dimensions were not found before scan data")
    color_model = {1: "grayscale", 3: "ycbcr-or-rgb", 4: "cmyk-or-ycck"}.get(components, "unknown")
    return {
        "format": "JPEG",
        "mime_type": "image/jpeg",
        "width_px": width,
        "height_px": height,
        "bit_depth": bit_depth,
        "color_model": color_model,
        "alpha": {"present": False, "channel_available": False, "detection": "format-does-not-support-alpha"},
        "exif": exif,
        "parse_warnings": parse_warnings,
    }


def _parse_gif(data: bytes) -> dict[str, Any]:
    if len(data) < 13 or data[:6] not in (b"GIF87a", b"GIF89a"):
        raise ImageParseError("Invalid or truncated GIF header")
    width, height = struct.unpack_from("<HH", data, 6)
    transparent = False
    search_at = 13
    while True:
        gce = data.find(b"\x21\xf9\x04", search_at)
        if gce < 0 or gce + 8 > len(data):
            break
        transparent = transparent or bool(data[gce + 3] & 0x01)
        search_at = gce + 8
    return {
        "format": "GIF",
        "mime_type": "image/gif",
        "width_px": width,
        "height_px": height,
        "bit_depth": (data[10] & 0x07) + 1,
        "color_model": "indexed",
        "alpha": {"present": transparent, "channel_available": False, "detection": "graphics-control-transparency"},
        "exif": _empty_exif(),
        "parse_warnings": [],
    }


def _parse_bmp(data: bytes) -> dict[str, Any]:
    if len(data) < 26 or data[:2] != b"BM":
        raise ImageParseError("Invalid or truncated BMP header")
    dib_size = _u32(data, 14, "<")
    if dib_size == 12:
        width, height, bits = _u16(data, 18, "<"), _u16(data, 20, "<"), _u16(data, 24, "<")
        top_down = False
        compression = 0
    elif dib_size >= 40 and len(data) >= 54:
        width = struct.unpack_from("<i", data, 18)[0]
        signed_height = struct.unpack_from("<i", data, 22)[0]
        top_down = signed_height < 0
        height = abs(signed_height)
        bits = _u16(data, 28, "<")
        compression = _u32(data, 30, "<")
        width = abs(width)
    else:
        raise ImageParseError(f"Unsupported BMP DIB header size: {dib_size}")
    alpha_value: bool | None = False
    detection = "bits-per-pixel"
    if bits == 32:
        # BI_ALPHABITFIELDS (6) or a non-zero V4/V5 alpha mask is definitive.
        alpha_mask = _u32(data, 14 + 52, "<") if dib_size >= 56 and len(data) >= 14 + 56 else 0
        if compression == 6 or alpha_mask:
            alpha_value = True
            detection = "alpha-bitfield"
        else:
            alpha_value = None
            detection = "32-bit-bmp-alpha-usage-ambiguous"
    return {
        "format": "BMP",
        "mime_type": "image/bmp",
        "width_px": width,
        "height_px": height,
        "bit_depth": bits,
        "color_model": "rgba-or-bgra" if bits == 32 else "rgb-or-indexed",
        "top_down": top_down,
        "alpha": {"present": alpha_value, "channel_available": bits == 32, "detection": detection},
        "exif": _empty_exif(),
        "parse_warnings": ["BMP has a 32-bit channel, but actual alpha usage cannot be proven from the header"] if alpha_value is None else [],
    }


def _parse_tiff(data: bytes) -> dict[str, Any]:
    if len(data) < 8 or data[:2] not in (b"II", b"MM"):
        raise ImageParseError("Invalid or truncated TIFF header")
    endian = "<" if data[:2] == b"II" else ">"
    if _u16(data, 2, endian) != 42:
        raise ImageParseError("Unsupported TIFF variant (classic TIFF required)")
    tags, _ = _parse_ifd(data, _u32(data, 4, endian), endian)
    width, height = _first(tags.get(256)), _first(tags.get(257))
    if not isinstance(width, int) or not isinstance(height, int):
        raise ImageParseError("TIFF width/height tags are missing")
    bits = tags.get(258)
    samples = _first(tags.get(277))
    extras = tags.get(338)
    if extras is not None:
        alpha_present: bool | None = True
        detection = "extra-samples-tag"
    elif isinstance(samples, int) and samples <= 3:
        alpha_present = False
        detection = "samples-per-pixel"
    else:
        alpha_present = None
        detection = "alpha-usage-ambiguous"
    exif, exif_warnings = _parse_exif(data)
    return {
        "format": "TIFF",
        "mime_type": "image/tiff",
        "width_px": width,
        "height_px": height,
        "bit_depth": bits,
        "color_model": "tiff-photometric",
        "alpha": {"present": alpha_present, "channel_available": alpha_present is not False, "detection": detection},
        "exif": exif,
        "parse_warnings": exif_warnings,
    }


def _parse_webp(data: bytes) -> dict[str, Any]:
    if len(data) < 20 or data[:4] != b"RIFF" or data[8:12] != b"WEBP":
        raise ImageParseError("Invalid or truncated WebP header")
    width = height = None
    alpha_present: bool | None = False
    alpha_available = False
    detection = "container-features"
    exif = _empty_exif()
    parse_warnings: list[str] = []
    offset = 12
    while offset + 8 <= len(data):
        fourcc = data[offset : offset + 4]
        size = _u32(data, offset + 4, "<")
        start, end = offset + 8, offset + 8 + size
        if end > len(data):
            parse_warnings.append("WebP contains a truncated chunk")
            break
        chunk = data[start:end]
        if fourcc == b"VP8X" and len(chunk) >= 10:
            flags = chunk[0]
            alpha_available = bool(flags & 0x10)
            alpha_present = alpha_available
            width = int.from_bytes(chunk[4:7], "little") + 1
            height = int.from_bytes(chunk[7:10], "little") + 1
        elif fourcc == b"VP8 " and len(chunk) >= 10 and chunk[3:6] == b"\x9d\x01\x2a":
            width = int.from_bytes(chunk[6:8], "little") & 0x3FFF
            height = int.from_bytes(chunk[8:10], "little") & 0x3FFF
        elif fourcc == b"VP8L" and len(chunk) >= 5 and chunk[0] == 0x2F:
            bits = int.from_bytes(chunk[1:5], "little")
            width = (bits & 0x3FFF) + 1
            height = ((bits >> 14) & 0x3FFF) + 1
            alpha_available = True
            alpha_present = bool((bits >> 28) & 1)
            detection = "vp8l-alpha-used-bit"
        elif fourcc == b"ALPH":
            alpha_available = alpha_present = True
            detection = "alpha-chunk"
        elif fourcc == b"EXIF":
            exif, exif_warnings = _parse_exif(chunk)
            parse_warnings.extend(exif_warnings)
        offset = end + (size & 1)
    if not width or not height:
        raise ImageParseError("WebP canvas dimensions were not found")
    return {
        "format": "WEBP",
        "mime_type": "image/webp",
        "width_px": width,
        "height_px": height,
        "bit_depth": 8,
        "color_model": "rgb-or-yuv",
        "alpha": {"present": alpha_present, "channel_available": alpha_available, "detection": detection},
        "exif": exif,
        "parse_warnings": parse_warnings,
    }


def inspect_image(data: bytes) -> dict[str, Any]:
    if data.startswith(b"\x89PNG\r\n\x1a\n"):
        return _parse_png(data)
    if data.startswith(b"\xff\xd8"):
        return _parse_jpeg(data)
    if data.startswith((b"GIF87a", b"GIF89a")):
        return _parse_gif(data)
    if data.startswith(b"BM"):
        return _parse_bmp(data)
    if data.startswith((b"II*\x00", b"MM\x00*")):
        return _parse_tiff(data)
    if data.startswith(b"RIFF") and len(data) >= 12 and data[8:12] == b"WEBP":
        return _parse_webp(data)
    raise ImageParseError("Unsupported image format; supported formats: PNG, JPEG, GIF, WebP, BMP, TIFF")


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _issue(code: str, message: str, **details: Any) -> dict[str, Any]:
    issue = {"code": code, "message": message}
    issue.update(details)
    return issue


def build_manifest(path: Path, expected_sha256: str | None, max_pixels: int) -> dict[str, Any]:
    errors: list[dict[str, Any]] = []
    warnings: list[dict[str, Any]] = []
    checks: list[dict[str, Any]] = []
    resolved = path.expanduser().resolve()
    source: dict[str, Any] = {
        "path": str(resolved),
        "file_name": resolved.name,
        "byte_size": None,
        "sha256": None,
        "format": None,
        "mime_type": None,
        "width_px": None,
        "height_px": None,
        "display_width_px": None,
        "display_height_px": None,
        "aspect_ratio": None,
        "display_aspect_ratio": None,
        "pixel_count": None,
        "orientation": None,
        "alpha": None,
        "exif": _empty_exif(),
        "image_details": {},
    }
    if not resolved.exists():
        errors.append(_issue("source.not_found", f"Source does not exist: {resolved}"))
        checks.append({"id": "source.exists", "status": "fail"})
        return _manifest(source, checks, errors, warnings)
    checks.append({"id": "source.exists", "status": "pass"})
    if not resolved.is_file():
        errors.append(_issue("source.not_file", f"Source is not a regular file: {resolved}"))
        checks.append({"id": "source.regular_file", "status": "fail"})
        return _manifest(source, checks, errors, warnings)
    checks.append({"id": "source.regular_file", "status": "pass"})
    try:
        stat = resolved.stat()
        source["byte_size"] = stat.st_size
        if stat.st_size == 0:
            errors.append(_issue("source.empty", "Source file is empty"))
            checks.append({"id": "source.nonempty", "status": "fail"})
            return _manifest(source, checks, errors, warnings)
        checks.append({"id": "source.nonempty", "status": "pass"})
        actual_sha256 = sha256_file(resolved)
        source["sha256"] = actual_sha256
        checks.append({"id": "source.sha256", "status": "pass", "value": actual_sha256})
        if expected_sha256:
            match = actual_sha256 == expected_sha256.lower()
            checks.append({"id": "source.expected_sha256", "status": "pass" if match else "fail"})
            if not match:
                errors.append(_issue("source.hash_mismatch", "Source SHA-256 does not match --expect-sha256", expected=expected_sha256.lower(), actual=actual_sha256))
        data = resolved.read_bytes()
        image = inspect_image(data)
    except (OSError, ImageParseError, struct.error) as exc:
        errors.append(_issue("source.inspect_failed", str(exc)))
        checks.append({"id": "image.parse", "status": "fail"})
        return _manifest(source, checks, errors, warnings)
    source["format"] = image.pop("format")
    source["mime_type"] = image.pop("mime_type", mimetypes.guess_type(resolved.name)[0])
    source["width_px"] = int(image.pop("width_px"))
    source["height_px"] = int(image.pop("height_px"))
    source["pixel_count"] = source["width_px"] * source["height_px"]
    source["aspect_ratio"] = source["width_px"] / source["height_px"]
    source["alpha"] = image.pop("alpha")
    source["exif"] = image.pop("exif")
    parse_warnings = image.pop("parse_warnings", [])
    source["image_details"] = image
    orientation_raw = _first(source["exif"].get("tags", {}).get("orientation"))
    orientation_value = orientation_raw if isinstance(orientation_raw, int) and orientation_raw in ORIENTATIONS else 1
    orientation = {"exif_value": orientation_raw, **ORIENTATIONS[orientation_value]}
    source["orientation"] = orientation
    if orientation["swaps_dimensions"]:
        source["display_width_px"], source["display_height_px"] = source["height_px"], source["width_px"]
    else:
        source["display_width_px"], source["display_height_px"] = source["width_px"], source["height_px"]
    source["display_aspect_ratio"] = source["display_width_px"] / source["display_height_px"]
    checks.append({"id": "image.parse", "status": "pass", "format": source["format"]})
    checks.append({"id": "image.dimensions", "status": "pass", "width_px": source["width_px"], "height_px": source["height_px"], "aspect_ratio": source["aspect_ratio"]})
    checks.append({"id": "image.alpha", "status": "pass" if source["alpha"]["present"] is not None else "warn", "present": source["alpha"]["present"]})
    checks.append({"id": "image.exif", "status": "pass" if source["exif"]["present"] and source["exif"]["parsed"] else "not_present" if not source["exif"]["present"] else "warn"})
    checks.append({"id": "image.orientation", "status": "pass", "effective": orientation["name"]})
    if orientation_raw not in (None, *ORIENTATIONS.keys()):
        warnings.append(_issue("image.orientation_invalid", f"Unknown EXIF orientation {orientation_raw}; treating it as normal"))
    elif orientation_value != 1:
        warnings.append(_issue("image.orientation_transform_required", f"Apply EXIF orientation '{orientation['name']}' before using source coordinates", exif_value=orientation_value))
    if source["pixel_count"] > max_pixels:
        warnings.append(_issue("image.pixel_count_large", f"Image has {source['pixel_count']:,} pixels, above the configured {max_pixels:,} threshold", pixel_count=source["pixel_count"], threshold=max_pixels))
    if source["alpha"]["present"] is None:
        warnings.append(_issue("image.alpha_ambiguous", "The container exposes alpha-capable data, but metadata cannot prove whether alpha is used"))
    for message in parse_warnings:
        warnings.append(_issue("image.metadata_warning", message))
    return _manifest(source, checks, errors, warnings)


def _manifest(source: dict[str, Any], checks: list[dict[str, Any]], errors: list[dict[str, Any]], warnings: list[dict[str, Any]]) -> dict[str, Any]:
    return {
        "manifest_version": MANIFEST_VERSION,
        "kind": "sci-diagram-pptx-source-manifest",
        "generated_at": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        "ok": not errors,
        "source": source,
        "checks": checks,
        "errors": errors,
        "warnings": warnings,
    }


class JsonArgumentParser(argparse.ArgumentParser):
    def error(self, message: str) -> None:
        payload = {"ok": False, "errors": [_issue("cli.invalid_arguments", message)]}
        print(json.dumps(payload, ensure_ascii=False), file=sys.stderr)
        raise SystemExit(2)


def _sha256_arg(value: str) -> str:
    lowered = value.lower()
    if len(lowered) != 64 or any(ch not in "0123456789abcdef" for ch in lowered):
        raise argparse.ArgumentTypeError("must be exactly 64 hexadecimal characters")
    return lowered


def _write_json(payload: dict[str, Any], output: Path | None, compact: bool) -> None:
    rendered = json.dumps(payload, ensure_ascii=False, sort_keys=False, indent=None if compact else 2) + "\n"
    if output is None:
        sys.stdout.write(rendered)
        return
    output = output.expanduser().resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    temporary = output.with_name(output.name + ".tmp")
    temporary.write_text(rendered, encoding="utf-8")
    os.replace(temporary, output)


def build_parser() -> argparse.ArgumentParser:
    parser = JsonArgumentParser(
        description="Read-only preflight for a source image; emits a JSON source manifest.",
        epilog="Supported formats: PNG, JPEG, GIF, WebP, BMP, and classic TIFF.",
    )
    parser.add_argument("source", type=Path, help="source image to inspect")
    parser.add_argument("--expect-sha256", type=_sha256_arg, help="fail if the source SHA-256 differs")
    parser.add_argument("--max-pixels", type=int, default=100_000_000, help="warn above this pixel count (default: 100000000)")
    parser.add_argument("--output", type=Path, help="write JSON to this file instead of stdout")
    parser.add_argument("--compact", action="store_true", help="emit compact JSON")
    parser.add_argument("--strict", action="store_true", help="return non-zero when warnings are present")
    return parser


def main(argv: Iterable[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if args.max_pixels <= 0:
        build_parser().error("--max-pixels must be greater than zero")
    try:
        payload = build_manifest(args.source, args.expect_sha256, args.max_pixels)
        if args.strict and payload["warnings"]:
            payload["ok"] = False
            payload["errors"].append(_issue("strict.warnings_present", "Warnings are errors in --strict mode", warning_count=len(payload["warnings"])))
        _write_json(payload, args.output, args.compact)
        return 0 if payload["ok"] else 1
    except OSError as exc:
        payload = {"ok": False, "errors": [_issue("io.write_failed", str(exc))]}
        print(json.dumps(payload, ensure_ascii=False), file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
