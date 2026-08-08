#!/usr/bin/env python3
"""Render a PPTX into a fresh directory and emit hash-bound slide evidence."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import subprocess
import sys
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable
from xml.etree import ElementTree as ET

try:
    from PIL import Image, UnidentifiedImageError
except ImportError as exc:  # pragma: no cover
    raise SystemExit("Pillow is required: python -m pip install Pillow") from exc


TOOL_VERSION = "1.0"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def write_json(payload: dict[str, Any], path: Path) -> None:
    target = path.expanduser().resolve()
    target.parent.mkdir(parents=True, exist_ok=True)
    temporary = target.with_name(target.name + ".tmp")
    temporary.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    os.replace(temporary, target)


def pptx_slide_count(path: Path) -> int | None:
    try:
        with zipfile.ZipFile(path) as zf:
            root = ET.fromstring(zf.read("ppt/presentation.xml"))
        return len(root.findall(f"./{{{P_NS}}}sldIdLst/{{{P_NS}}}sldId"))
    except (OSError, KeyError, zipfile.BadZipFile, ET.ParseError):
        return None


def natural_key(path: Path) -> list[Any]:
    return [int(piece) if piece.isdigit() else piece.lower() for piece in re.split(r"(\d+)", path.name)]


def build_report(pptx_path: Path, render_script_path: Path, render_dir_path: Path, timeout: int, helper_python_path: Path | None = None) -> dict[str, Any]:
    pptx = pptx_path.expanduser().resolve()
    render_script = render_script_path.expanduser().resolve()
    render_dir = render_dir_path.expanduser().resolve()
    helper_python = (helper_python_path or Path(sys.executable)).expanduser().resolve()
    checks: list[dict[str, Any]] = []

    def add(check_id: str, status: str, message: str, **evidence: Any) -> None:
        item: dict[str, Any] = {"id": check_id, "status": status, "severity": "HARD", "message": message}
        if evidence:
            item["evidence"] = evidence
        checks.append(item)

    inputs_ok = pptx.is_file() and render_script.is_file() and render_script.suffix.lower() == ".py" and helper_python.is_file() and os.access(helper_python, os.X_OK)
    add("render.inputs_readable", "PASS" if inputs_ok else "FAIL", "PPTX, render helper, and helper Python are readable/executable" if inputs_ok else "PPTX, render helper, or helper Python is missing/invalid", pptx=str(pptx), render_script=str(render_script), helper_python=str(helper_python))
    render_dir_fresh = not render_dir.exists() or (render_dir.is_dir() and not any(render_dir.iterdir()))
    add("render.fresh_output_directory", "PASS" if render_dir_fresh else "FAIL", "Render directory is new or empty" if render_dir_fresh else "Render directory is not empty; stale render provenance cannot be excluded", render_dir=str(render_dir))
    if inputs_ok and render_dir_fresh:
        render_dir.mkdir(parents=True, exist_ok=True)

    command = [str(helper_python), str(render_script), str(pptx), "--output_dir", str(render_dir)]
    return_code: int | None = None
    stdout = stderr = ""
    timed_out = False
    if inputs_ok and render_dir_fresh:
        try:
            process = subprocess.run(command, capture_output=True, text=True, timeout=timeout, check=False)
            return_code, stdout, stderr = process.returncode, process.stdout, process.stderr
        except subprocess.TimeoutExpired as exc:
            timed_out = True
            stdout = exc.stdout.decode("utf-8", "replace") if isinstance(exc.stdout, bytes) else (exc.stdout or "")
            stderr = exc.stderr.decode("utf-8", "replace") if isinstance(exc.stderr, bytes) else (exc.stderr or "")
    helper_ok = return_code == 0 and not timed_out
    add("render.helper_exit", "PASS" if helper_ok else "FAIL", "Render helper exited successfully" if helper_ok else "Render helper failed or timed out", return_code=return_code, timed_out=timed_out, timeout_seconds=timeout)
    explicit_success = helper_ok and "Slides rendered to " in stdout and not stderr.strip()
    add("render.explicit_result", "PASS" if explicit_success else "FAIL", "Render helper explicitly reported a clean render" if explicit_success else "Render helper output was not an unambiguous clean success")

    expected_slides = pptx_slide_count(pptx) if pptx.is_file() else None
    png_paths = sorted(render_dir.glob("*.png"), key=natural_key) if render_dir.is_dir() else []
    render_records: list[dict[str, Any]] = []
    invalid_pngs: list[str] = []
    for index, path in enumerate(png_paths, start=1):
        try:
            with Image.open(path) as image:
                image.verify()
            with Image.open(path) as image:
                width, height = image.size
        except (OSError, UnidentifiedImageError):
            invalid_pngs.append(str(path))
            continue
        render_records.append({"slide_index": index, "path": str(path.resolve()), "sha256": sha256_file(path), "width_px": width, "height_px": height})
    files_ok = expected_slides is not None and expected_slides >= 1 and len(render_records) == expected_slides and not invalid_pngs
    add(
        "render.slide_files",
        "PASS" if files_ok else "FAIL",
        "Fresh render contains one valid PNG per PPTX slide" if files_ok else "Rendered PNG files are missing, extra, invalid, or do not match PPTX slide count",
        expected_slide_count=expected_slides, png_count=len(png_paths), valid_png_count=len(render_records), invalid_pngs=invalid_pngs,
    )

    hard_failures = [{"id": item["id"], "message": item["message"], "severity": "HARD"} for item in checks if item["status"] == "FAIL"]
    python_version = None
    if helper_python.is_file() and os.access(helper_python, os.X_OK):
        try:
            version_process = subprocess.run([str(helper_python), "--version"], capture_output=True, text=True, timeout=10, check=False)
            python_version = (version_process.stdout or version_process.stderr).strip() or None
        except (OSError, subprocess.TimeoutExpired):
            pass
    return {
        "schema_version": "1.0",
        "kind": "sci-diagram-pptx-render-report",
        "tool": "sci-diagram-pptx/render_evidence.py",
        "tool_version": TOOL_VERSION,
        "generated_at": utc_now(),
        "status": "FAIL" if hard_failures else "PASS",
        "hard_failures": hard_failures,
        "warnings": [],
        "checks": checks,
        "pptx": str(pptx),
        "pptx_sha256": sha256_file(pptx) if pptx.is_file() else None,
        "render_script": str(render_script),
        "render_script_sha256": sha256_file(render_script) if render_script.is_file() else None,
        "helper_python": str(helper_python),
        "helper_python_sha256": sha256_file(helper_python) if helper_python.is_file() else None,
        "helper_python_version": python_version,
        "render_dir": str(render_dir),
        "command": command,
        "return_code": return_code,
        "stdout": stdout[-100_000:],
        "stderr": stderr[-100_000:],
        "renders": render_records,
    }


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Render a PPTX in a fresh directory and bind slide PNG hashes to the PPTX.")
    parser.add_argument("pptx", type=Path)
    parser.add_argument("--render-script", type=Path, required=True)
    parser.add_argument("--render-dir", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--helper-python", type=Path, default=Path(sys.executable), help="Python interpreter used to run render_slides.py")
    parser.add_argument("--timeout", type=int, default=300)
    return parser


def main(argv: Iterable[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if args.timeout <= 0:
        build_parser().error("--timeout must be positive")
    try:
        report = build_report(args.pptx, args.render_script, args.render_dir, args.timeout, args.helper_python)
        write_json(report, args.output)
        return 0 if report["status"] == "PASS" else 1
    except OSError as exc:
        payload = {
            "schema_version": "1.0", "kind": "sci-diagram-pptx-render-report",
            "tool": "sci-diagram-pptx/render_evidence.py", "tool_version": TOOL_VERSION,
            "generated_at": utc_now(), "status": "FAIL",
            "hard_failures": [{"id": "render.operational", "message": str(exc), "severity": "HARD"}],
            "warnings": [], "checks": [{"id": "render.operational", "status": "FAIL", "severity": "HARD", "message": str(exc)}],
        }
        try:
            write_json(payload, args.output)
        except OSError:
            print(json.dumps(payload, ensure_ascii=False), file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
