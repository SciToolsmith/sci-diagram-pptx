#!/usr/bin/env python3
"""Align a slide render to its source and emit QA-only visual evidence.

Outputs are deliberately limited to ``report.json`` and ``heatmap.png``.  The
aligned image is kept in memory and must not be used as presentation content.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import os
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable

try:
    from PIL import Image, ImageChops, ImageFilter, ImageOps, ImageStat, UnidentifiedImageError
except ImportError as exc:  # pragma: no cover - exercised only without dependency
    raise SystemExit("Pillow is required: python -m pip install Pillow") from exc


TOOL_VERSION = "1.0"
RESAMPLING = getattr(Image, "Resampling", Image)


class CompareError(RuntimeError):
    """An operational error that prevents trustworthy image comparison."""


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")


def write_json(payload: dict[str, Any], output: Path) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    temporary = output.with_name(output.name + ".tmp")
    temporary.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    os.replace(temporary, output)


def flatten_to_rgb(image: Image.Image) -> Image.Image:
    rgba = image.convert("RGBA")
    white = Image.new("RGBA", rgba.size, (255, 255, 255, 255))
    return Image.alpha_composite(white, rgba).convert("RGB")


def open_image(path: Path) -> tuple[Image.Image, dict[str, Any]]:
    resolved = path.expanduser().resolve()
    if not resolved.exists() or not resolved.is_file():
        raise CompareError(f"Image does not exist or is not a regular file: {resolved}")
    try:
        digest = hashlib.sha256()
        with resolved.open("rb") as source_handle:
            for block in iter(lambda: source_handle.read(1024 * 1024), b""):
                digest.update(block)
        with Image.open(resolved) as opened:
            metadata = {
                "path": str(resolved),
                "sha256": digest.hexdigest(),
                "format": opened.format,
                "mode": opened.mode,
                "width_px": opened.width,
                "height_px": opened.height,
                "frame_count": int(getattr(opened, "n_frames", 1)),
                "has_icc_profile": bool(opened.info.get("icc_profile")),
                "has_alpha": "A" in opened.getbands() or "transparency" in opened.info,
                "exif_orientation_applied": True,
            }
            oriented = ImageOps.exif_transpose(opened)
            metadata["display_width_px"] = oriented.width
            metadata["display_height_px"] = oriented.height
            return flatten_to_rgb(oriented.copy()), metadata
    except (OSError, UnidentifiedImageError) as exc:
        raise CompareError(f"Cannot decode image '{resolved}': {exc}") from exc


def overlap_score(reference: Image.Image, candidate: Image.Image, dx: int, dy: int) -> tuple[float, float]:
    """Mean absolute luma error over overlap for candidate shifted by dx/dy."""
    width, height = reference.size
    ref_x = max(0, dx)
    ref_y = max(0, dy)
    cand_x = max(0, -dx)
    cand_y = max(0, -dy)
    overlap_width = min(width - ref_x, width - cand_x)
    overlap_height = min(height - ref_y, height - cand_y)
    if overlap_width <= 0 or overlap_height <= 0:
        return float("inf"), 0.0
    ref_crop = reference.crop((ref_x, ref_y, ref_x + overlap_width, ref_y + overlap_height))
    cand_crop = candidate.crop((cand_x, cand_y, cand_x + overlap_width, cand_y + overlap_height))
    mean_error = ImageStat.Stat(ImageChops.difference(ref_crop, cand_crop)).mean[0]
    overlap_ratio = overlap_width * overlap_height / (width * height)
    # Prevent a translation from winning merely by discarding difficult borders.
    score = mean_error + (1.0 - overlap_ratio) * 8.0
    return float(score), overlap_ratio


def best_translation(reference_rgb: Image.Image, candidate_rgb: Image.Image, max_shift: int) -> dict[str, Any]:
    if max_shift <= 0:
        return {"dx_px": 0, "dy_px": 0, "search_score": None, "overlap_ratio": 1.0, "at_search_boundary": False, "structure_stddev": None}
    width, height = reference_rgb.size
    scale = min(1.0, 256.0 / max(width, height))
    small_size = (max(8, round(width * scale)), max(8, round(height * scale)))
    reference = reference_rgb.convert("L").resize(small_size, RESAMPLING.BILINEAR)
    candidate = candidate_rgb.convert("L").resize(small_size, RESAMPLING.BILINEAR)
    small_shift = max(1, int(math.ceil(max_shift * scale)))

    best: tuple[float, int, int, float] | None = None
    for dy in range(-small_shift, small_shift + 1):
        for dx in range(-small_shift, small_shift + 1):
            score, overlap = overlap_score(reference, candidate, dx, dy)
            current = (score, abs(dx) + abs(dy), dx, dy, overlap)
            if best is None or current[:2] < (best[0], abs(best[1]) + abs(best[2])):
                best = (score, dx, dy, overlap)
    assert best is not None
    coarse_dx = round(best[1] / scale)
    coarse_dy = round(best[2] / scale)

    full_reference = reference_rgb.convert("L")
    full_candidate = candidate_rgb.convert("L")
    refined: tuple[float, int, int, float] | None = None
    radius = max(2, int(math.ceil(1.5 / max(scale, 1e-9))))
    for dy in range(max(-max_shift, coarse_dy - radius), min(max_shift, coarse_dy + radius) + 1):
        for dx in range(max(-max_shift, coarse_dx - radius), min(max_shift, coarse_dx + radius) + 1):
            score, overlap = overlap_score(full_reference, full_candidate, dx, dy)
            current = (score, abs(dx) + abs(dy), dx, dy, overlap)
            if refined is None or current[:2] < (refined[0], abs(refined[1]) + abs(refined[2])):
                refined = (score, dx, dy, overlap)
    assert refined is not None
    stddev = float(ImageStat.Stat(reference).stddev[0])
    return {
        "dx_px": refined[1],
        "dy_px": refined[2],
        "search_score": refined[0] / 255.0,
        "overlap_ratio": refined[3],
        "at_search_boundary": abs(refined[1]) == max_shift or abs(refined[2]) == max_shift,
        "structure_stddev": stddev,
    }


def shifted_canvas(image: Image.Image, dx: int, dy: int, size: tuple[int, int]) -> Image.Image:
    canvas = Image.new("RGB", size, "white")
    canvas.paste(image, (dx, dy))
    return canvas


def max_channel_difference(difference_rgb: Image.Image) -> Image.Image:
    red, green, blue = difference_rgb.split()
    return ImageChops.lighter(ImageChops.lighter(red, green), blue)


def image_metrics(reference: Image.Image, aligned: Image.Image, pixel_threshold: int) -> tuple[dict[str, Any], Image.Image]:
    difference = ImageChops.difference(reference, aligned)
    pixels = reference.width * reference.height
    channel_count = pixels * 3
    total_abs = 0
    total_sq = 0
    maximum = 0
    histogram = difference.histogram()
    for channel in range(3):
        channel_histogram = histogram[channel * 256 : (channel + 1) * 256]
        total_abs += sum(value * count for value, count in enumerate(channel_histogram))
        total_sq += sum(value * value * count for value, count in enumerate(channel_histogram))
        for value in range(255, -1, -1):
            if channel_histogram[value]:
                maximum = max(maximum, value)
                break
    mae_255 = total_abs / channel_count
    rmse_255 = math.sqrt(total_sq / channel_count)
    maximum_channel = max_channel_difference(difference)
    max_histogram = maximum_channel.histogram()
    mismatch_pixels = sum(max_histogram[pixel_threshold + 1 :])
    exact_pixels = max_histogram[0]

    ref_edges = reference.convert("L").filter(ImageFilter.FIND_EDGES)
    aligned_edges = aligned.convert("L").filter(ImageFilter.FIND_EDGES)
    edge_diff = ImageChops.difference(ref_edges, aligned_edges)
    edge_mae = ImageStat.Stat(edge_diff).mean[0] / 255.0
    psnr = None if rmse_255 == 0 else 20.0 * math.log10(255.0 / rmse_255)
    bbox = maximum_channel.point(lambda value: 255 if value > pixel_threshold else 0).getbbox()
    return {
        "mae": mae_255 / 255.0,
        "mae_255": mae_255,
        "rmse": rmse_255 / 255.0,
        "rmse_255": rmse_255,
        "psnr_db": psnr,
        "maximum_channel_error": maximum,
        "pixel_threshold": pixel_threshold,
        "mismatch_pixel_ratio": mismatch_pixels / pixels,
        "mismatch_pixel_count": mismatch_pixels,
        "exact_pixel_ratio": exact_pixels / pixels,
        "edge_mae": edge_mae,
        "difference_bbox_px": list(bbox) if bbox else None,
    }, maximum_channel


def count_mask(mask: Image.Image) -> int:
    histogram = mask.convert("L").histogram()
    return sum(histogram[128:])


def interior_edges(image: Image.Image, threshold: int = 24) -> Image.Image:
    edges = image.convert("L").filter(ImageFilter.FIND_EDGES)
    width, height = edges.size
    border = min(2, width // 2, height // 2)
    if border:
        edges.paste(0, (0, 0, width, border))
        edges.paste(0, (0, height - border, width, height))
        edges.paste(0, (0, 0, border, height))
        edges.paste(0, (width - border, 0, width, height))
    return edges.point(lambda value: 255 if value > threshold else 0, mode="L")


def estimated_background(image: Image.Image) -> tuple[int, int, int]:
    width, height = image.size
    patch_width = max(1, min(32, width // 10))
    patch_height = max(1, min(32, height // 10))
    boxes = (
        (0, 0, patch_width, patch_height),
        (width - patch_width, 0, width, patch_height),
        (0, height - patch_height, patch_width, height),
        (width - patch_width, height - patch_height, width, height),
    )
    channel_medians = [ImageStat.Stat(image.crop(box)).median for box in boxes]
    result = []
    for channel in range(3):
        values = sorted(sample[channel] for sample in channel_medians)
        result.append(round((values[1] + values[2]) / 2))
    return tuple(result)  # type: ignore[return-value]


def structural_recall_metrics(reference: Image.Image, aligned: Image.Image, pixel_threshold: int) -> dict[str, Any]:
    ref_edges = interior_edges(reference)
    candidate_edges = interior_edges(aligned)
    reference_edge_pixels = count_mask(ref_edges)
    # A small dilation tolerates renderer antialiasing and sub-pixel glyph drift.
    candidate_edge_dilated = candidate_edges.filter(ImageFilter.MaxFilter(5))
    matched_edges = count_mask(ImageChops.multiply(ref_edges, candidate_edge_dilated))
    edge_recall = matched_edges / reference_edge_pixels if reference_edge_pixels else None

    background = estimated_background(reference)
    background_image = Image.new("RGB", reference.size, background)
    foreground_delta = max_channel_difference(ImageChops.difference(reference, background_image))
    foreground_threshold = max(18, pixel_threshold)
    foreground_mask = foreground_delta.point(lambda value: 255 if value > foreground_threshold else 0, mode="L")
    foreground_pixels = count_mask(foreground_mask)
    comparison_delta = max_channel_difference(ImageChops.difference(reference, aligned))
    foreground_match_threshold = max(32, pixel_threshold * 2)
    matching_pixels = comparison_delta.point(lambda value: 255 if value <= foreground_match_threshold else 0, mode="L")
    matched_foreground = count_mask(ImageChops.multiply(foreground_mask, matching_pixels))
    foreground_recall = matched_foreground / foreground_pixels if foreground_pixels else None

    blank = Image.new("RGB", reference.size, "white")
    blank_mae = sum(ImageStat.Stat(ImageChops.difference(reference, blank)).mean) / (3 * 255.0)
    aligned_mae = sum(ImageStat.Stat(ImageChops.difference(reference, aligned)).mean) / (3 * 255.0)
    blank_improvement = (blank_mae - aligned_mae) / blank_mae if blank_mae > 0 else None
    render_stddev = sum(ImageStat.Stat(aligned.convert("L")).stddev) / 1.0
    return {
        "reference_background_rgb": list(background),
        "reference_edge_pixels": reference_edge_pixels,
        "matched_reference_edge_pixels": matched_edges,
        "edge_recall": edge_recall,
        "reference_foreground_pixels": foreground_pixels,
        "matched_reference_foreground_pixels": matched_foreground,
        "foreground_recall": foreground_recall,
        "foreground_threshold_255": foreground_threshold,
        "foreground_match_threshold_255": foreground_match_threshold,
        "blank_baseline_mae": blank_mae,
        "aligned_mae_for_blank_baseline": aligned_mae,
        "blank_baseline_improvement": blank_improvement,
        "aligned_render_luma_stddev": render_stddev,
    }


def heatmap_from_difference(maximum_channel: Image.Image) -> Image.Image:
    """False-colour map: black=equal, blue/cyan/yellow/red=increasing error."""
    palette: list[int] = []
    for value in range(256):
        t = value / 255.0
        if t == 0:
            red = green = blue = 0
        elif t < 0.25:
            red, green, blue = 0, round(4 * t * 255), 255
        elif t < 0.5:
            red, green, blue = 0, 255, round((2 - 4 * t) * 255)
        elif t < 0.75:
            red, green, blue = round((4 * t - 2) * 255), 255, 0
        else:
            red, green, blue = 255, round((4 - 4 * t) * 255), 0
        palette.extend((max(0, min(255, red)), max(0, min(255, green)), max(0, min(255, blue))))
    indexed = maximum_channel.copy()
    indexed.putpalette(palette)
    return indexed.convert("RGB")


class ReportBuilder:
    def __init__(self) -> None:
        self.checks: list[dict[str, Any]] = []

    def add(self, check_id: str, status: str, severity: str, message: str, **evidence: Any) -> None:
        check = {"id": check_id, "status": status.upper(), "severity": severity.upper(), "message": message}
        if evidence:
            check["evidence"] = evidence
        self.checks.append(check)

    def finish(self, **payload: Any) -> dict[str, Any]:
        hard = [{"id": c["id"], "message": c["message"], "severity": c["severity"]} for c in self.checks if c["severity"] == "HARD" and c["status"] == "FAIL"]
        warnings = [{"id": c["id"], "message": c["message"], "severity": c["severity"]} for c in self.checks if c["status"] == "WARN"]
        return {
            "schema_version": "1.0",
            "kind": "sci-diagram-pptx-visual-report",
            "tool": "sci-diagram-pptx/compare_render.py",
            "tool_version": TOOL_VERSION,
            "generated_at": utc_now(),
            "purpose": "QA_ONLY_DO_NOT_USE_AS_DELIVERABLE_VISUAL",
            "status": "FAIL" if hard else "WARN" if warnings else "PASS",
            "hard_failures": hard,
            "warnings": warnings,
            "checks": self.checks,
            **payload,
        }


def compare_images(
    reference_path: Path,
    render_path: Path,
    output_dir: Path,
    max_shift: int = 16,
    pixel_threshold: int = 16,
    max_mae: float = 0.08,
    max_mismatch_ratio: float = 0.25,
    max_edge_mae: float = 0.15,
    aspect_tolerance: float = 0.005,
    min_edge_recall: float = 0.55,
    min_foreground_recall: float = 0.55,
    min_blank_improvement: float = 0.15,
    min_render_stddev: float = 0.75,
    max_accepted_shift: int | None = None,
) -> dict[str, Any]:
    output = output_dir.expanduser().resolve()
    output.mkdir(parents=True, exist_ok=True)
    report = ReportBuilder()
    reference, reference_meta = open_image(reference_path)
    render, render_meta = open_image(render_path)
    report.add("visual.inputs_readable", "PASS", "HARD", "Reference and render were decoded by Pillow")

    if reference_meta["frame_count"] != 1 or render_meta["frame_count"] != 1:
        report.add("visual.single_frame_inputs", "WARN", "HARD", "A multi-frame input was detected; only the first frame was compared", reference_frames=reference_meta["frame_count"], render_frames=render_meta["frame_count"])
    else:
        report.add("visual.single_frame_inputs", "PASS", "INFO", "Both inputs are single-frame images")

    if reference_meta["has_icc_profile"] or render_meta["has_icc_profile"]:
        report.add("visual.color_management", "WARN", "INFO", "An ICC profile is present; this Pillow-only comparison does not perform profile conversion", reference_has_icc=reference_meta["has_icc_profile"], render_has_icc=render_meta["has_icc_profile"])
    else:
        report.add("visual.color_management", "PASS", "INFO", "Neither input declares an ICC profile")

    reference_ratio = reference.width / reference.height
    render_ratio = render.width / render.height
    aspect_error = abs(render_ratio / reference_ratio - 1.0)
    report.add(
        "visual.aspect_ratio",
        "PASS" if aspect_error <= aspect_tolerance else "FAIL",
        "HARD",
        "Reference and render aspect ratios agree" if aspect_error <= aspect_tolerance else "Reference and render aspect ratios differ; resized metrics are diagnostic only",
        reference_aspect_ratio=reference_ratio,
        render_aspect_ratio=render_ratio,
        relative_error=aspect_error,
        tolerance=aspect_tolerance,
    )

    resized = render.resize(reference.size, RESAMPLING.LANCZOS) if render.size != reference.size else render.copy()
    alignment = best_translation(reference, resized, max_shift)
    aligned = shifted_canvas(resized, alignment["dx_px"], alignment["dy_px"], reference.size)
    accepted_shift = max_accepted_shift if max_accepted_shift is not None else max(2, round(min(reference.size) * 0.005))
    if alignment["structure_stddev"] is not None and alignment["structure_stddev"] < 2.0:
        report.add("visual.alignment_confidence", "WARN", "HARD", "Reference has too little luminance structure to validate translation alignment", alignment=alignment)
    elif alignment["at_search_boundary"]:
        report.add("visual.alignment_confidence", "WARN", "HARD", "Best translation lies on the search boundary; a larger offset may exist", alignment=alignment)
    else:
        report.add("visual.alignment_confidence", "PASS", "INFO", "Translation alignment converged inside the configured search window", alignment=alignment)
    alignment_magnitude = max(abs(alignment["dx_px"]), abs(alignment["dy_px"]))
    report.add(
        "visual.alignment_magnitude",
        "PASS" if alignment_magnitude <= accepted_shift else "FAIL",
        "HARD",
        "Required alignment translation is within the acceptance limit" if alignment_magnitude <= accepted_shift else "Required alignment translation exceeds the acceptance limit; large layout displacement cannot be hidden by registration",
        dx_px=alignment["dx_px"], dy_px=alignment["dy_px"], magnitude_chebyshev_px=alignment_magnitude, threshold_px=accepted_shift,
    )

    metrics, max_difference = image_metrics(reference, aligned, pixel_threshold)
    structural = structural_recall_metrics(reference, aligned, pixel_threshold)
    metrics.update(structural)
    heatmap_path = output / "heatmap.png"
    heatmap_from_difference(max_difference).save(heatmap_path, format="PNG")
    report.add("visual.heatmap_written", "PASS", "HARD", "QA-only false-colour heatmap was written", path=str(heatmap_path))
    report.add(
        "visual.mae_threshold",
        "PASS" if metrics["mae"] <= max_mae else "FAIL",
        "HARD",
        "Mean absolute error is within the configured limit" if metrics["mae"] <= max_mae else "Mean absolute error exceeds the configured limit",
        actual=metrics["mae"], threshold=max_mae,
    )
    report.add(
        "visual.mismatch_ratio_threshold",
        "PASS" if metrics["mismatch_pixel_ratio"] <= max_mismatch_ratio else "FAIL",
        "HARD",
        "Mismatch-pixel ratio is within the configured limit" if metrics["mismatch_pixel_ratio"] <= max_mismatch_ratio else "Mismatch-pixel ratio exceeds the configured limit",
        actual=metrics["mismatch_pixel_ratio"], threshold=max_mismatch_ratio, pixel_threshold=pixel_threshold,
    )
    report.add(
        "visual.edge_difference",
        "PASS" if metrics["edge_mae"] <= max_edge_mae else "WARN",
        "SOFT",
        "Edge difference is within the diagnostic limit" if metrics["edge_mae"] <= max_edge_mae else "Edge difference is high; inspect fine geometry, glyphs, and connector routing",
        actual=metrics["edge_mae"], threshold=max_edge_mae,
    )
    if structural["edge_recall"] is None:
        report.add("visual.edge_recall", "WARN", "HARD", "Reference edge mask is empty; edge preservation cannot be proven", reference_edge_pixels=structural["reference_edge_pixels"])
    else:
        report.add(
            "visual.edge_recall",
            "PASS" if structural["edge_recall"] >= min_edge_recall else "FAIL",
            "HARD",
            "Reference-edge recall is within the configured limit" if structural["edge_recall"] >= min_edge_recall else "Reference-edge recall is too low; the render may be blank or structurally incomplete",
            actual=structural["edge_recall"], threshold=min_edge_recall,
        )
    if structural["foreground_recall"] is None:
        report.add("visual.foreground_recall", "WARN", "HARD", "Reference foreground mask is empty; foreground preservation cannot be proven", reference_foreground_pixels=structural["reference_foreground_pixels"])
    else:
        report.add(
            "visual.foreground_recall",
            "PASS" if structural["foreground_recall"] >= min_foreground_recall else "FAIL",
            "HARD",
            "Reference-foreground recall is within the configured limit" if structural["foreground_recall"] >= min_foreground_recall else "Reference-foreground recall is too low; visible source content is missing or inaccurate",
            actual=structural["foreground_recall"], threshold=min_foreground_recall,
        )
    if structural["blank_baseline_improvement"] is None:
        report.add("visual.blank_baseline_guard", "WARN", "HARD", "Reference is indistinguishable from a blank white canvas under MAE; blank-render rejection is unproven")
    else:
        report.add(
            "visual.blank_baseline_guard",
            "PASS" if structural["blank_baseline_improvement"] >= min_blank_improvement else "FAIL",
            "HARD",
            "Render improves materially over a blank white slide" if structural["blank_baseline_improvement"] >= min_blank_improvement else "Render does not improve enough over a blank white slide",
            actual_improvement=structural["blank_baseline_improvement"], threshold=min_blank_improvement,
        )
    report.add(
        "visual.render_not_blank",
        "PASS" if structural["aligned_render_luma_stddev"] >= min_render_stddev else "FAIL",
        "HARD",
        "Aligned render contains measurable luminance structure" if structural["aligned_render_luma_stddev"] >= min_render_stddev else "Aligned render is blank or nearly uniform",
        actual_luma_stddev=structural["aligned_render_luma_stddev"], threshold=min_render_stddev,
    )

    return report.finish(
        input={
            "reference": reference_meta,
            "render": render_meta,
            "reference_sha256": reference_meta["sha256"],
            "render_sha256": render_meta["sha256"],
        },
        alignment={
            "method": "render-resize-to-reference-then-luma-translation-search",
            "render_resized": render.size != reference.size,
            "original_render_size": list(render.size),
            "comparison_size": list(reference.size),
            **alignment,
        },
        thresholds={
            "aspect_tolerance": aspect_tolerance,
            "max_shift_px": max_shift,
            "max_accepted_shift_px": accepted_shift,
            "pixel_threshold_255": pixel_threshold,
            "max_mae": max_mae,
            "max_mismatch_ratio": max_mismatch_ratio,
            "max_edge_mae": max_edge_mae,
            "min_edge_recall": min_edge_recall,
            "min_foreground_recall": min_foreground_recall,
            "min_blank_improvement": min_blank_improvement,
            "min_render_stddev": min_render_stddev,
        },
        metrics=metrics,
        artifacts={"report": str(output / "report.json"), "heatmap": str(heatmap_path)},
        limitations=[
            "Pixel metrics do not prove semantic correctness, topology, editability, or formula object type.",
            "Renderer font hinting and antialiasing can create legitimate pixel differences.",
            "Alignment corrects only uniform resize and translation; it does not hide rotation, crop, or local distortion.",
            "The heatmap is QA evidence only and must not be inserted into the deliverable PPTX.",
        ],
    )


class JsonArgumentParser(argparse.ArgumentParser):
    def error(self, message: str) -> None:
        print(json.dumps({"status": "FAIL", "hard_failures": [{"id": "cli.invalid_arguments", "message": message}], "warnings": [], "checks": []}, ensure_ascii=False), file=sys.stderr)
        raise SystemExit(2)


def build_parser() -> argparse.ArgumentParser:
    parser = JsonArgumentParser(description="Compare a source diagram with a slide render and write QA-only evidence.")
    parser.add_argument("--reference", type=Path, required=True, help="locked source image")
    parser.add_argument("--render", type=Path, required=True, help="rendered Slide 1 image")
    parser.add_argument("--output-dir", type=Path, required=True, help="directory for report.json and heatmap.png")
    parser.add_argument("--max-shift", type=int, default=16, help="maximum translation search in reference pixels (default: 16)")
    parser.add_argument("--pixel-threshold", type=int, default=16, help="per-pixel max-channel mismatch threshold 0..255 (default: 16)")
    parser.add_argument("--max-mae", type=float, default=0.08, help="hard normalized MAE limit (default: 0.08)")
    parser.add_argument("--max-mismatch-ratio", type=float, default=0.25, help="hard mismatch-pixel ratio limit (default: 0.25)")
    parser.add_argument("--max-edge-mae", type=float, default=0.15, help="diagnostic edge-MAE warning limit (default: 0.15)")
    parser.add_argument("--aspect-tolerance", type=float, default=0.005, help="hard relative aspect-ratio tolerance (default: 0.005)")
    parser.add_argument("--min-edge-recall", type=float, default=0.55, help="hard source-edge recall minimum (default: 0.55)")
    parser.add_argument("--min-foreground-recall", type=float, default=0.55, help="hard source-foreground recall minimum (default: 0.55)")
    parser.add_argument("--min-blank-improvement", type=float, default=0.15, help="hard MAE improvement over a blank white render (default: 0.15)")
    parser.add_argument("--min-render-stddev", type=float, default=0.75, help="hard minimum render luma standard deviation (default: 0.75)")
    parser.add_argument("--max-accepted-shift", type=int, help="hard accepted translation in pixels (default: max(2, 0.5%% of shorter reference side))")
    return parser


def operational_failure(message: str) -> dict[str, Any]:
    check = {"id": "visual.operational", "status": "FAIL", "severity": "HARD", "message": message}
    return {
        "schema_version": "1.0",
        "kind": "sci-diagram-pptx-visual-report",
        "tool": "sci-diagram-pptx/compare_render.py",
        "tool_version": TOOL_VERSION,
        "generated_at": utc_now(),
        "purpose": "QA_ONLY_DO_NOT_USE_AS_DELIVERABLE_VISUAL",
        "status": "FAIL",
        "hard_failures": [{"id": check["id"], "message": message, "severity": "HARD"}],
        "warnings": [],
        "checks": [check],
    }


def main(argv: Iterable[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if args.max_shift < 0 or not 0 <= args.pixel_threshold <= 255:
        build_parser().error("--max-shift must be non-negative and --pixel-threshold must be 0..255")
    for name in ("max_mae", "max_mismatch_ratio", "max_edge_mae", "aspect_tolerance", "min_edge_recall", "min_foreground_recall", "min_blank_improvement"):
        value = getattr(args, name)
        if not 0 <= value <= 1:
            build_parser().error(f"--{name.replace('_', '-')} must be between 0 and 1")
    if args.min_render_stddev < 0:
        build_parser().error("--min-render-stddev must be non-negative")
    if args.max_accepted_shift is not None and args.max_accepted_shift < 0:
        build_parser().error("--max-accepted-shift must be non-negative")
    output = args.output_dir.expanduser().resolve()
    try:
        payload = compare_images(
            args.reference, args.render, output, args.max_shift, args.pixel_threshold,
            args.max_mae, args.max_mismatch_ratio, args.max_edge_mae, args.aspect_tolerance,
            args.min_edge_recall, args.min_foreground_recall, args.min_blank_improvement, args.min_render_stddev,
            args.max_accepted_shift,
        )
        write_json(payload, output / "report.json")
        return 1 if payload["hard_failures"] else 0
    except (CompareError, OSError, ValueError) as exc:
        payload = operational_failure(str(exc))
        try:
            output.mkdir(parents=True, exist_ok=True)
            write_json(payload, output / "report.json")
        except OSError:
            print(json.dumps(payload, ensure_ascii=False), file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
