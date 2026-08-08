#!/usr/bin/env python3
"""Materialize an explicitly selected schematic panel without rescaling.

The parent image is decoded with Pillow, EXIF orientation is applied, and the
integer bounding box is interpreted in displayed-pixel coordinates.  The crop
is saved as PNG without resizing, color conversion, enhancement, or retouching.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable

try:
    from PIL import Image, ImageOps
except ImportError:  # pragma: no cover - exercised only on a broken runtime
    Image = None  # type: ignore[assignment]
    ImageOps = None  # type: ignore[assignment]


TOOL_NAME = "sci-diagram-pptx/panel_crop.py"
TOOL_VERSION = "1.0"
MANIFEST_KIND = "sci-diagram-pptx-panel-crop-manifest"
ISO_DATETIME = re.compile(r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(?::\d{2}(?:\.\d+)?)?(?:Z|[+-]\d{2}:\d{2})$")


class CropError(ValueError):
    """A user-correctable panel crop failure."""


def _sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _visible_pixel_sha256(image: Any) -> str:
    rgba = image.convert("RGBA")
    digest = hashlib.sha256()
    digest.update(f"RGBA:{rgba.width}x{rgba.height}:".encode("ascii"))
    digest.update(rgba.tobytes())
    return digest.hexdigest()


def _issue(code: str, message: str, **evidence: Any) -> dict[str, Any]:
    item: dict[str, Any] = {"id": code, "code": code, "message": message}
    if evidence:
        item["evidence"] = evidence
    return item


def _failure(code: str, message: str, **evidence: Any) -> dict[str, Any]:
    issue = _issue(code, message, **evidence)
    return {
        "kind": MANIFEST_KIND,
        "tool": TOOL_NAME,
        "tool_version": TOOL_VERSION,
        "status": "FAIL",
        "ok": False,
        "hard_failures": [issue],
        "errors": [issue],
        "warnings": [],
        "checks": [
            {
                "id": code,
                "status": "FAIL",
                "message": message,
                "evidence": evidence,
            }
        ],
    }


def _validate_selected_at(value: str) -> None:
    if not ISO_DATETIME.fullmatch(value):
        raise CropError("--selected-at must be an ISO-8601 timestamp with a timezone")
    try:
        parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError as exc:
        raise CropError("--selected-at is not a valid calendar timestamp") from exc
    if parsed.tzinfo is None:
        raise CropError("--selected-at must include Z or an explicit UTC offset")


def _stage_png(image: Any, destination: Path) -> Path:
    destination.parent.mkdir(parents=True, exist_ok=True)
    descriptor, raw_path = tempfile.mkstemp(prefix=f".{destination.name}.", suffix=".tmp", dir=destination.parent)
    os.close(descriptor)
    staged = Path(raw_path)
    try:
        # Explicit format avoids extension inference.  No conversion, resize,
        # enhancement, palette rewrite, or metadata-driven retouch is applied.
        image.save(staged, format="PNG", optimize=False, compress_level=6)
        os.chmod(staged, 0o644)
        with staged.open("rb") as handle:
            os.fsync(handle.fileno())
        return staged
    except Exception as exc:
        staged.unlink(missing_ok=True)
        raise CropError(f"Could not save crop as an exact PNG without mode conversion: {exc}") from exc


def _stage_json(payload: dict[str, Any], destination: Path) -> Path:
    destination.parent.mkdir(parents=True, exist_ok=True)
    descriptor, raw_path = tempfile.mkstemp(prefix=f".{destination.name}.", suffix=".tmp", dir=destination.parent)
    staged = Path(raw_path)
    rendered = json.dumps(payload, ensure_ascii=False, indent=2) + "\n"
    try:
        with os.fdopen(descriptor, "w", encoding="utf-8") as handle:
            handle.write(rendered)
            handle.flush()
            os.fsync(handle.fileno())
        os.chmod(staged, 0o644)
        return staged
    except Exception:
        staged.unlink(missing_ok=True)
        raise


def _publish_noreplace(staged: Path, destination: Path) -> tuple[int, int]:
    """Atomically publish one staged file and refuse an existing destination."""
    try:
        os.link(staged, destination)
    except FileExistsError as exc:
        raise CropError(f"Output already exists; refusing to overwrite: {destination}") from exc
    except OSError as exc:
        raise CropError(f"Could not atomically publish {destination}: {exc}") from exc
    identity = destination.stat()
    staged.unlink()
    return identity.st_dev, identity.st_ino


def _rollback_if_same(path: Path, identity: tuple[int, int] | None) -> None:
    if identity is None:
        return
    try:
        stat = path.stat()
        if (stat.st_dev, stat.st_ino) == identity:
            path.unlink()
    except FileNotFoundError:
        pass


def build_manifest(
    parent_path: Path,
    bbox: tuple[int, int, int, int],
    label: str,
    selected_by: str,
    selected_at: str,
    evidence: str,
    output_image: Path,
) -> tuple[dict[str, Any], Any]:
    if Image is None or ImageOps is None:
        raise CropError("Pillow is required; install the 'Pillow' package")
    parent_path = parent_path.expanduser().resolve()
    if not parent_path.exists() or not parent_path.is_file():
        raise CropError(f"Parent image does not exist or is not a regular file: {parent_path}")
    if not label.strip() or not selected_by.strip() or not evidence.strip():
        raise CropError("--label, --selected-by, and --evidence must be non-empty")
    _validate_selected_at(selected_at)
    x, y, width, height = bbox
    if x < 0 or y < 0 or width <= 0 or height <= 0:
        raise CropError("--bbox requires X/Y >= 0 and W/H > 0")

    parent_sha256 = _sha256_file(parent_path)
    try:
        with Image.open(parent_path) as opened:
            opened.load()
            raw_width, raw_height = opened.size
            raw_mode = opened.mode
            raw_format = opened.format or "UNKNOWN"
            try:
                exif_orientation = int(opened.getexif().get(274, 1))
            except (TypeError, ValueError):
                exif_orientation = 1
            displayed = ImageOps.exif_transpose(opened)
            displayed.load()
    except Exception as exc:
        raise CropError(f"Could not decode parent image: {exc}") from exc

    display_width, display_height = displayed.size
    if x + width > display_width or y + height > display_height:
        raise CropError(
            f"Panel bbox ({x}, {y}, {width}, {height}) exceeds displayed parent size "
            f"{display_width}x{display_height}"
        )
    crop = displayed.crop((x, y, x + width, y + height))
    if crop.size != (width, height):
        raise CropError("Pillow returned an unexpected crop size")

    checks = [
        {
            "id": "parent.identity-and-orientation",
            "status": "PASS",
            "message": "Parent image was identity-hashed and EXIF-oriented before cropping",
            "evidence": {
                "sha256": parent_sha256,
                "raw_size": [raw_width, raw_height],
                "display_size": [display_width, display_height],
                "exif_orientation": exif_orientation,
            },
        },
        {
            "id": "selection.user-explicit",
            "status": "PASS",
            "message": "Panel crop carries explicit user-selection provenance",
            "evidence": {"selected_by": selected_by, "selected_at": selected_at, "label": label},
        },
        {
            "id": "bbox.within-displayed-parent",
            "status": "PASS",
            "message": "Integer panel bbox is fully contained in displayed parent pixels",
            "evidence": {"bbox": {"x": x, "y": y, "width": width, "height": height}},
        },
    ]
    manifest: dict[str, Any] = {
        "kind": MANIFEST_KIND,
        "tool": TOOL_NAME,
        "tool_version": TOOL_VERSION,
        "generated_at": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        "status": "PASS",
        "ok": True,
        "hard_failures": [],
        "errors": [],
        "warnings": [],
        "checks": checks,
        "parent": {
            "path": str(parent_path),
            "sha256": parent_sha256,
            "format": raw_format,
            "mode": raw_mode,
            "raw_width_px": raw_width,
            "raw_height_px": raw_height,
            "display_width_px": display_width,
            "display_height_px": display_height,
            "exif_orientation": exif_orientation,
            "orientation_applied": exif_orientation not in (0, 1),
        },
        "panel": {
            "label": label,
            "bbox": {"x": x, "y": y, "width": width, "height": height},
        },
        "crop": {
            "path": str(output_image.expanduser().resolve()),
            "sha256": None,
            "format": "PNG",
            "mode": crop.mode,
            "width_px": width,
            "height_px": height,
            "visible_pixel_sha256": _visible_pixel_sha256(crop),
        },
        "selection": {
            "selection_source": "user-explicit",
            "selected_by": selected_by,
            "selected_at": selected_at,
            "evidence": evidence,
        },
    }
    return manifest, crop


class JsonArgumentParser(argparse.ArgumentParser):
    def error(self, message: str) -> None:
        print(json.dumps(_failure("cli.invalid_arguments", message), ensure_ascii=False), file=sys.stderr)
        raise SystemExit(2)


def build_parser() -> argparse.ArgumentParser:
    parser = JsonArgumentParser(description="Create an exact EXIF-oriented panel crop and a provenance manifest.")
    parser.add_argument("parent", type=Path, help="parent image")
    parser.add_argument("--bbox", nargs=4, required=True, type=int, metavar=("X", "Y", "W", "H"), help="integer bbox in displayed parent pixels")
    parser.add_argument("--label", required=True, help="user-selected panel label or description")
    parser.add_argument("--selected-by", required=True, help="identity or role that made the explicit selection")
    parser.add_argument("--selected-at", required=True, help="timezone-aware ISO-8601 selection timestamp")
    parser.add_argument("--evidence", required=True, help="reference to the explicit user selection evidence")
    parser.add_argument("--output-image", required=True, type=Path, help="new PNG crop path; must not already exist")
    parser.add_argument("--output-manifest", required=True, type=Path, help="new JSON manifest path; must not already exist")
    return parser


def main(argv: Iterable[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    output_image = args.output_image.expanduser().resolve()
    output_manifest = args.output_manifest.expanduser().resolve()
    staged_image: Path | None = None
    staged_manifest: Path | None = None
    image_identity: tuple[int, int] | None = None
    try:
        if output_image == output_manifest:
            raise CropError("--output-image and --output-manifest must be different paths")
        for output in (output_image, output_manifest):
            if output.exists() or output.is_symlink():
                raise CropError(f"Output already exists; refusing to overwrite: {output}")

        manifest, crop = build_manifest(
            args.parent,
            tuple(args.bbox),
            args.label,
            args.selected_by,
            args.selected_at,
            args.evidence,
            output_image,
        )
        staged_image = _stage_png(crop, output_image)
        with Image.open(staged_image) as verification:
            verification.load()
            actual_pixel_hash = _visible_pixel_sha256(verification)
            actual_size = verification.size
        expected_pixel_hash = manifest["crop"]["visible_pixel_sha256"]
        if actual_size != (manifest["crop"]["width_px"], manifest["crop"]["height_px"]):
            raise CropError("Staged PNG dimensions differ from the requested crop")
        if actual_pixel_hash != expected_pixel_hash:
            raise CropError("Staged PNG visible pixels differ from the exact in-memory crop")
        manifest["crop"]["sha256"] = _sha256_file(staged_image)
        manifest["checks"].append(
            {
                "id": "crop.pixel-exact",
                "status": "PASS",
                "message": "Saved PNG has the exact requested dimensions and visible pixels",
                "evidence": {
                    "visible_pixel_sha256": actual_pixel_hash,
                    "file_sha256": manifest["crop"]["sha256"],
                    "size": list(actual_size),
                },
            }
        )
        staged_manifest = _stage_json(manifest, output_manifest)

        image_identity = _publish_noreplace(staged_image, output_image)
        staged_image = None
        try:
            _publish_noreplace(staged_manifest, output_manifest)
            staged_manifest = None
        except Exception:
            _rollback_if_same(output_image, image_identity)
            image_identity = None
            raise
        print(json.dumps(manifest, ensure_ascii=False, indent=2))
        return 0
    except (CropError, OSError, ValueError) as exc:
        payload = _failure("panel_crop.failed", str(exc))
        print(json.dumps(payload, ensure_ascii=False), file=sys.stderr)
        return 1
    finally:
        if staged_image is not None:
            staged_image.unlink(missing_ok=True)
        if staged_manifest is not None:
            staged_manifest.unlink(missing_ok=True)


if __name__ == "__main__":
    raise SystemExit(main())
