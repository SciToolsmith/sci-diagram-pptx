#!/usr/bin/env python3
"""Render a PPTX with an isolated LibreOffice profile and pdftoppm."""

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import tempfile
from pathlib import Path


def find_command(explicit: str | None, candidates: tuple[str, ...]) -> str:
    if explicit:
        resolved = shutil.which(explicit) if not Path(explicit).is_file() else explicit
        if resolved:
            return str(resolved)
        raise RuntimeError(f"command not found: {explicit}")
    for candidate in candidates:
        resolved = shutil.which(candidate)
        if resolved:
            return resolved
    raise RuntimeError(f"required command not found: {' or '.join(candidates)}")


def run_checked(command: list[str], timeout: int) -> subprocess.CompletedProcess[str]:
    result = subprocess.run(
        command,
        check=False,
        capture_output=True,
        text=True,
        timeout=timeout,
    )
    if result.returncode != 0:
        detail = "\n".join(part.strip() for part in (result.stdout, result.stderr) if part.strip())
        raise RuntimeError(f"command failed ({result.returncode}): {' '.join(command)}\n{detail}")
    return result


def render(
    source: Path,
    output_dir: Path,
    dpi: int,
    keep_pdf: bool,
    timeout: int,
    soffice_arg: str | None,
    pdftoppm_arg: str | None,
) -> dict[str, object]:
    source = source.resolve()
    if not source.is_file() or source.suffix.lower() != ".pptx":
        raise RuntimeError(f"input must be an existing .pptx file: {source}")
    if dpi < 72 or dpi > 600:
        raise RuntimeError("dpi must be between 72 and 600")

    output_dir = output_dir.resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    existing_outputs = sorted(output_dir.glob("slide-*.png"))
    if existing_outputs or (output_dir / "rendered.pdf").exists():
        raise RuntimeError(
            "output directory already contains render outputs; use a fresh task directory"
        )
    soffice = find_command(soffice_arg, ("libreoffice", "soffice"))
    pdftoppm = find_command(pdftoppm_arg, ("pdftoppm",))

    with tempfile.TemporaryDirectory(prefix=".sci-lo-profile-", dir=output_dir) as profile_raw:
        with tempfile.TemporaryDirectory(prefix=".sci-render-", dir=output_dir) as work_raw:
            profile = Path(profile_raw).resolve()
            work = Path(work_raw).resolve()
            profile_uri = profile.as_uri()

            run_checked(
                [
                    soffice,
                    f"-env:UserInstallation={profile_uri}",
                    "--headless",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    str(work),
                    str(source),
                ],
                timeout,
            )

            pdf_candidates = sorted(work.glob("*.pdf"))
            if len(pdf_candidates) != 1:
                raise RuntimeError(f"LibreOffice produced {len(pdf_candidates)} PDF files")
            generated_pdf = pdf_candidates[0]

            prefix = output_dir / "slide"
            run_checked(
                [pdftoppm, "-png", "-r", str(dpi), str(generated_pdf), str(prefix)],
                timeout,
            )
            slides = sorted(output_dir.glob("slide-*.png"))
            if not slides:
                raise RuntimeError("pdftoppm produced no slide images")

            kept_pdf: Path | None = None
            if keep_pdf:
                kept_pdf = output_dir / "rendered.pdf"
                shutil.copy2(generated_pdf, kept_pdf)

    return {
        "input": str(source),
        "outputDir": str(output_dir),
        "slides": [str(path) for path in slides],
        "pdf": str(kept_pdf) if kept_pdf else None,
        "dpi": dpi,
        "isolatedLibreOfficeProfile": True,
    }


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pptx", type=Path)
    parser.add_argument("--output-dir", type=Path, required=True)
    parser.add_argument("--dpi", type=int, default=160)
    parser.add_argument("--keep-pdf", action="store_true")
    parser.add_argument("--timeout", type=int, default=120)
    parser.add_argument("--soffice")
    parser.add_argument("--pdftoppm")
    return parser


def main() -> int:
    args = build_parser().parse_args()
    try:
        report = render(
            source=args.pptx,
            output_dir=args.output_dir,
            dpi=args.dpi,
            keep_pdf=args.keep_pdf,
            timeout=args.timeout,
            soffice_arg=args.soffice,
            pdftoppm_arg=args.pdftoppm,
        )
    except (RuntimeError, subprocess.TimeoutExpired) as error:
        print(json.dumps({"ready": False, "error": str(error)}, ensure_ascii=False))
        return 1
    print(json.dumps({"ready": True, **report}, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
