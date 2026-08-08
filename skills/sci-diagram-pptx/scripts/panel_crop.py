#!/usr/bin/env python3
"""Crop one explicitly selected diagram panel without resizing or retouching."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import sys
import tempfile
from pathlib import Path

try:
    from PIL import Image, ImageOps
except ImportError:  # pragma: no cover
    Image = None  # type: ignore[assignment]
    ImageOps = None  # type: ignore[assignment]


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Crop a user-selected panel in displayed-pixel coordinates."
    )
    parser.add_argument("source", type=Path, help="parent image")
    parser.add_argument(
        "--bbox",
        nargs=4,
        required=True,
        type=int,
        metavar=("X", "Y", "W", "H"),
        help="0-based bbox after EXIF orientation",
    )
    parser.add_argument("--output", required=True, type=Path, help="new PNG path")
    parser.add_argument(
        "--output-json",
        type=Path,
        help="optional compact crop record",
    )
    return parser


def fail(message: str) -> int:
    print(f"panel_crop: {message}", file=sys.stderr)
    return 1


def save_new_png(image: object, output: Path) -> None:
    if output.exists():
        raise ValueError(f"refusing to overwrite existing output: {output}")
    output.parent.mkdir(parents=True, exist_ok=True)
    descriptor, temporary_name = tempfile.mkstemp(
        prefix=f".{output.name}.", suffix=".tmp", dir=output.parent
    )
    os.close(descriptor)
    temporary = Path(temporary_name)
    try:
        image.save(temporary, format="PNG", optimize=False)  # type: ignore[attr-defined]
        os.link(temporary, output)
    finally:
        temporary.unlink(missing_ok=True)


def write_new_json(payload: dict[str, object], output: Path) -> None:
    if output.exists():
        raise ValueError(f"refusing to overwrite existing output: {output}")
    output.parent.mkdir(parents=True, exist_ok=True)
    rendered = json.dumps(payload, ensure_ascii=False, indent=2) + "\n"
    descriptor, temporary_name = tempfile.mkstemp(
        prefix=f".{output.name}.", suffix=".tmp", dir=output.parent
    )
    temporary = Path(temporary_name)
    try:
        with os.fdopen(descriptor, "w", encoding="utf-8") as handle:
            handle.write(rendered)
        os.link(temporary, output)
    finally:
        temporary.unlink(missing_ok=True)


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if Image is None or ImageOps is None:
        return fail("Pillow is required")

    source = args.source.expanduser().resolve()
    output = args.output.expanduser().resolve()
    record_path = args.output_json.expanduser().resolve() if args.output_json else None
    if not source.is_file():
        return fail(f"source image does not exist: {source}")
    if output.exists() or (record_path and record_path.exists()):
        return fail("an output path already exists; choose a new path")

    x, y, width, height = args.bbox
    if x < 0 or y < 0 or width <= 0 or height <= 0:
        return fail("bbox requires X/Y >= 0 and W/H > 0")

    try:
        with Image.open(source) as opened:
            displayed = ImageOps.exif_transpose(opened)
            displayed.load()
        parent_width, parent_height = displayed.size
        if x + width > parent_width or y + height > parent_height:
            return fail(
                f"bbox exceeds displayed image size {parent_width}x{parent_height}"
            )
        crop = displayed.crop((x, y, x + width, y + height))
        save_new_png(crop, output)
        if record_path:
            write_new_json(
                {
                    "source_sha256": sha256_file(source),
                    "source_size_px": [parent_width, parent_height],
                    "bbox": [x, y, width, height],
                    "output_sha256": sha256_file(output),
                    "output_size_px": [width, height],
                },
                record_path,
            )
    except Exception as exc:
        output.unlink(missing_ok=True)
        if record_path:
            record_path.unlink(missing_ok=True)
        return fail(str(exc))

    print(output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
