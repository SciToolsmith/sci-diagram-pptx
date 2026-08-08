#!/usr/bin/env python3
"""Smoke tests for source, panel-crop, and scene-plan preflight.

Pass real source images as positional arguments.  They are read only: this test
does not copy or modify them.  Panel-crop tests create only temporary synthetic
PNGs under a system temp directory and never generate a PPTX.
"""

from __future__ import annotations

import argparse
import contextlib
import copy
import hashlib
import io
import json
import sys
import tempfile
from pathlib import Path
from typing import Any

REPO_ROOT = Path(__file__).resolve().parents[1]
SKILL_DIR = REPO_ROOT / "skills" / "sci-diagram-pptx"
SCRIPTS_DIR = SKILL_DIR / "scripts"
sys.path.insert(0, str(SCRIPTS_DIR))

import source_preflight
import panel_crop
import validate_scene_plan

from PIL import Image


def _valid_plan(sha256: str, width_px: int, height_px: int) -> dict[str, Any]:
    return {
        "schema_version": "1.0",
        "kind": "sci-diagram-pptx-scene-plan",
        "source": {
            "sha256": sha256,
            "format": "PNG",
            "width_px": width_px,
            "height_px": height_px,
        },
        "canvas": {
            "coordinate_system": "normalized-top-left",
            "width_inches": 10.0,
            "height_inches": 10.0 * height_px / width_px,
        },
        "objects": [
            {
                "id": "node-a",
                "type": "shape",
                "source_bbox": {"x": 0.1, "y": 0.2, "width": 0.25, "height": 0.2},
                "bbox": {"x": 0.1, "y": 0.2, "width": 0.25, "height": 0.2},
                "z_order": 1,
                "confidence": 0.98,
                "reconstruction": {"mode": "native-exact", "expected_ooxml_kind": "p:sp"},
                "content": {"text": "Input", "language": "en"},
                "style": {"shape_geometry": "roundRect", "fill": {}, "line": {}, "text": {}},
            },
            {
                "id": "node-b",
                "type": "shape",
                "source_bbox": {"x": 0.65, "y": 0.2, "width": 0.25, "height": 0.2},
                "bbox": {"x": 0.65, "y": 0.2, "width": 0.25, "height": 0.2},
                "z_order": 2,
                "confidence": 0.96,
                "reconstruction": {"mode": "native-exact", "expected_ooxml_kind": "p:sp"},
                "content": {"text": "x → y", "math_source": "x → y", "language": "und"},
                "style": {"shape_geometry": "textbox", "fill": {}, "line": {}, "text": {}},
            },
        ],
        "connections": [
            {
                "id": "edge-a-b",
                "from": {"object_id": "node-a", "anchor": "right"},
                "to": {"object_id": "node-b", "anchor": "left"},
                "source_route": [{"x": 0.35, "y": 0.3}, {"x": 0.65, "y": 0.3}],
                "route": [{"x": 0.35, "y": 0.3}, {"x": 0.65, "y": 0.3}],
                "z_order": 0,
                "confidence": 0.94,
                "reconstruction": {"mode": "native-exact", "expected_ooxml_kind": "p:cxnSp"},
                "style": {
                    "kind": "attached-connector",
                    "arrow_at": "end",
                    "dash": "solid",
                    "line_width_px": 1.5,
                    "line_color": "#334155",
                },
            }
        ],
        "approvals": {"degradations": []},
    }


def _check(condition: bool, name: str, checks: list[dict[str, Any]], **details: Any) -> None:
    entry = {"name": name, "pass": bool(condition), **details}
    checks.append(entry)


def _file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _test_panel_crop(validator: validate_scene_plan.LightweightSchemaValidator, checks: list[dict[str, Any]]) -> None:
    with tempfile.TemporaryDirectory(prefix="sci-diagram-panel-crop-test-") as temporary_dir:
        temporary = Path(temporary_dir)
        parent_path = temporary / "oriented-parent.png"
        crop_path = temporary / "panel.png"
        manifest_path = temporary / "panel-manifest.json"

        synthetic = Image.new("RGB", (6, 4))
        for y in range(synthetic.height):
            for x in range(synthetic.width):
                synthetic.putpixel((x, y), ((x * 37) % 256, (y * 71) % 256, ((x + y) * 29) % 256))
        exif = Image.Exif()
        exif[274] = 6  # raw 6x4 -> displayed 4x6 after a 90-degree clockwise transform
        synthetic.save(parent_path, format="PNG", exif=exif)

        crop_args = [
            str(parent_path),
            "--bbox",
            "1",
            "1",
            "2",
            "3",
            "--label",
            "panel-A",
            "--selected-by",
            "user",
            "--selected-at",
            "2026-08-08T11:00:00Z",
            "--evidence",
            "The user explicitly selected panel-A in the active conversation.",
            "--output-image",
            str(crop_path),
            "--output-manifest",
            str(manifest_path),
        ]
        captured_out, captured_err = io.StringIO(), io.StringIO()
        with contextlib.redirect_stdout(captured_out), contextlib.redirect_stderr(captured_err):
            crop_rc = panel_crop.main(crop_args)
        crop_manifest = json.loads(manifest_path.read_text(encoding="utf-8")) if manifest_path.exists() else {}
        with Image.open(crop_path) as cropped_image:
            cropped_image.load()
            cropped_size = cropped_image.size
        _check(
            crop_rc == 0
            and crop_manifest.get("kind") == "sci-diagram-pptx-panel-crop-manifest"
            and crop_manifest.get("tool") == "sci-diagram-pptx/panel_crop.py"
            and crop_manifest.get("tool_version") == "1.0"
            and crop_manifest.get("status") == "PASS"
            and crop_manifest.get("hard_failures") == []
            and crop_manifest.get("parent", {}).get("raw_width_px") == 6
            and crop_manifest.get("parent", {}).get("raw_height_px") == 4
            and crop_manifest.get("parent", {}).get("display_width_px") == 4
            and crop_manifest.get("parent", {}).get("display_height_px") == 6
            and crop_manifest.get("parent", {}).get("exif_orientation") == 6
            and cropped_size == (2, 3)
            and crop_manifest.get("crop", {}).get("sha256") == _file_sha256(crop_path)
            and all(item.get("status") == "PASS" for item in crop_manifest.get("checks", [])),
            "panel-crop-exif-oriented-exact-pixels",
            checks,
            exit=crop_rc,
            stderr=captured_err.getvalue(),
            manifest=crop_manifest,
        )

        original_crop_hash = _file_sha256(crop_path)
        original_manifest_hash = _file_sha256(manifest_path)
        with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
            overwrite_rc = panel_crop.main(crop_args)
        _check(
            overwrite_rc != 0
            and _file_sha256(crop_path) == original_crop_hash
            and _file_sha256(manifest_path) == original_manifest_hash,
            "panel-crop-refuses-existing-outputs",
            checks,
            exit=overwrite_rc,
        )

        invalid_crop_path = temporary / "invalid-panel.png"
        invalid_manifest_path = temporary / "invalid-panel-manifest.json"
        invalid_args = list(crop_args)
        bbox_at = invalid_args.index("--bbox")
        invalid_args[bbox_at + 1 : bbox_at + 5] = ["3", "1", "2", "3"]
        invalid_args[invalid_args.index("--output-image") + 1] = str(invalid_crop_path)
        invalid_args[invalid_args.index("--output-manifest") + 1] = str(invalid_manifest_path)
        with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
            invalid_rc = panel_crop.main(invalid_args)
        _check(
            invalid_rc != 0 and not invalid_crop_path.exists() and not invalid_manifest_path.exists(),
            "panel-crop-rejects-out-of-bounds-bbox",
            checks,
            exit=invalid_rc,
        )

        panel_plan = _valid_plan(
            crop_manifest["crop"]["sha256"],
            crop_manifest["crop"]["width_px"],
            crop_manifest["crop"]["height_px"],
        )
        panel_plan["source"].update(
            {
                "parent_sha256": crop_manifest["parent"]["sha256"],
                "parent_width_px": crop_manifest["parent"]["display_width_px"],
                "parent_height_px": crop_manifest["parent"]["display_height_px"],
                "panel_manifest_sha256": _file_sha256(manifest_path),
                "panel_bbox": crop_manifest["panel"]["bbox"],
                "panel_label": crop_manifest["panel"]["label"],
                "panel_selection": crop_manifest["selection"],
            }
        )
        panel_schema_errors = validator.validate(panel_plan)
        panel_semantic_errors, panel_semantic_warnings, _ = validate_scene_plan.validate_semantics(panel_plan, 0.75, True)
        lineage_report, lineage_errors, lineage_warnings = validate_scene_plan.verify_panel_manifest(panel_plan, manifest_path)
        _check(
            not panel_schema_errors
            and not panel_semantic_errors
            and not panel_semantic_warnings
            and lineage_report.get("status") == "pass"
            and not lineage_errors
            and not lineage_warnings
            and lineage_report.get("comparisons_checked") == 12,
            "panel-manifest-lineage-match",
            checks,
            schema_errors=panel_schema_errors,
            semantic_errors=panel_semantic_errors,
            lineage_report=lineage_report,
            lineage_errors=lineage_errors,
        )

        panel_plan_path = temporary / "panel-scene-plan.json"
        validation_report_path = temporary / "panel-scene-plan-validation.json"
        panel_plan_path.write_text(json.dumps(panel_plan), encoding="utf-8")
        with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
            validation_rc = validate_scene_plan.main(
                [
                    str(panel_plan_path),
                    "--expected-source-sha256",
                    crop_manifest["crop"]["sha256"],
                    "--panel-manifest",
                    str(manifest_path),
                    "--output",
                    str(validation_report_path),
                ]
            )
        validation_report = json.loads(validation_report_path.read_text(encoding="utf-8"))
        _check(
            validation_rc == 0
            and validation_report.get("status") == "PASS"
            and validation_report.get("panel_manifest_match", {}).get("status") == "pass"
            and any(
                item.get("id") == "panel-lineage-manifest" and item.get("status") == "PASS"
                for item in validation_report.get("checks", [])
            ),
            "panel-manifest-validator-cli-roundtrip",
            checks,
            exit=validation_rc,
            report=validation_report,
        )

        missing_report, missing_errors, _ = validate_scene_plan.verify_panel_manifest(panel_plan, None)
        _check(
            missing_report.get("status") == "fail"
            and any(item.get("code") == "panel_manifest.required" for item in missing_errors),
            "panel-manifest-required-for-lineage",
            checks,
            errors=missing_errors,
        )

        tampered_plan = copy.deepcopy(panel_plan)
        tampered_plan["source"]["panel_bbox"]["x"] += 1
        tampered_report, tampered_errors, _ = validate_scene_plan.verify_panel_manifest(tampered_plan, manifest_path)
        _check(
            tampered_report.get("status") == "fail"
            and any(item.get("code") == "panel_manifest.metadata_mismatch" for item in tampered_errors),
            "panel-manifest-rejects-lineage-mismatch",
            checks,
            errors=tampered_errors,
        )

        wrong_hash_plan = copy.deepcopy(panel_plan)
        wrong_hash_plan["source"]["panel_manifest_sha256"] = "0" * 64
        hash_report, hash_errors, _ = validate_scene_plan.verify_panel_manifest(wrong_hash_plan, manifest_path)
        _check(
            hash_report.get("status") == "fail"
            and any(item.get("code") == "panel_manifest.hash_mismatch" for item in hash_errors),
            "panel-manifest-rejects-file-hash-mismatch",
            checks,
            errors=hash_errors,
        )

        failed_manifest_path = temporary / "failed-panel-manifest.json"
        failed_manifest = copy.deepcopy(crop_manifest)
        failed_manifest["kind"] = "not-a-panel-crop-manifest"
        failed_manifest["status"] = "FAIL"
        failed_manifest["ok"] = False
        failed_manifest["hard_failures"] = [{"id": "synthetic.failure", "message": "negative fixture"}]
        failed_manifest_path.write_text(json.dumps(failed_manifest), encoding="utf-8")
        failed_manifest_plan = copy.deepcopy(panel_plan)
        failed_manifest_plan["source"]["panel_manifest_sha256"] = _file_sha256(failed_manifest_path)
        failed_report, failed_errors, _ = validate_scene_plan.verify_panel_manifest(
            failed_manifest_plan, failed_manifest_path
        )
        failed_codes = {item.get("code") for item in failed_errors}
        _check(
            failed_report.get("status") == "fail"
            and {"panel_manifest.kind", "panel_manifest.status", "panel_manifest.hard_failures"}.issubset(failed_codes),
            "panel-manifest-rejects-non-pass-or-wrong-kind",
            checks,
            errors=failed_errors,
        )

        standalone_plan = _valid_plan("e" * 64, 100, 70)
        unexpected_report, unexpected_errors, _ = validate_scene_plan.verify_panel_manifest(standalone_plan, manifest_path)
        _check(
            unexpected_report.get("status") == "fail"
            and any(item.get("code") == "panel_manifest.unexpected" for item in unexpected_errors),
            "panel-manifest-rejected-for-standalone-plan",
            checks,
            errors=unexpected_errors,
        )


def main() -> int:
    parser = argparse.ArgumentParser(description="Run read-only sci-diagram-pptx preflight smoke tests.")
    parser.add_argument("sources", nargs="*", type=Path, help="existing PNG sources to preflight read-only")
    parser.add_argument("--expect-count", type=int, help="fail unless this many source paths are supplied")
    args = parser.parse_args()
    checks: list[dict[str, Any]] = []
    source_results: list[dict[str, Any]] = []
    if args.expect_count is not None:
        _check(len(args.sources) == args.expect_count, "source-count", checks, expected=args.expect_count, actual=len(args.sources))

    seed_sha = "a" * 64
    seed_width, seed_height = 1000, 700
    for source in args.sources:
        manifest = source_preflight.build_manifest(source, None, 100_000_000)
        source_info = manifest["source"]
        source_results.append(
            {
                "path": str(source.expanduser().resolve()),
                "ok": manifest["ok"],
                "format": source_info.get("format"),
                "width_px": source_info.get("width_px"),
                "height_px": source_info.get("height_px"),
                "sha256": source_info.get("sha256"),
                "warning_count": len(manifest["warnings"]),
            }
        )
        _check(manifest["ok"], f"source-preflight:{source.name}", checks)
        _check(source_info.get("format") == "PNG", f"source-is-png:{source.name}", checks, actual=source_info.get("format"))
        _check(isinstance(source_info.get("sha256"), str) and len(source_info["sha256"]) == 64, f"source-hash:{source.name}", checks)
        _check(
            all(
                source_info.get(key) is not None
                for key in ("width_px", "height_px", "aspect_ratio", "display_aspect_ratio", "exif", "alpha", "orientation")
            ),
            f"source-metadata:{source.name}",
            checks,
        )
        if manifest["ok"] and seed_sha == "a" * 64:
            seed_sha = source_info["sha256"]
            seed_width, seed_height = source_info["width_px"], source_info["height_px"]

    if args.sources and source_results and source_results[0]["ok"]:
        wrong_hash = "0" * 64 if source_results[0]["sha256"] != "0" * 64 else "1" * 64
        mismatch_manifest = source_preflight.build_manifest(args.sources[0], wrong_hash, 100_000_000)
        _check(
            mismatch_manifest["ok"] is False
            and any(item.get("code") == "source.hash_mismatch" for item in mismatch_manifest["errors"]),
            "source-preflight-hash-mismatch-rejected",
            checks,
        )

    schema_path = SCRIPTS_DIR / "scene-plan.schema.json"
    schema = json.loads(schema_path.read_text(encoding="utf-8"))
    validator = validate_scene_plan.LightweightSchemaValidator(schema)
    example_path = SKILL_DIR / "references" / "scene-plan.example.json"
    bundled_example = json.loads(example_path.read_text(encoding="utf-8"))
    example_schema_errors = validator.validate(bundled_example)
    example_semantic_errors, example_semantic_warnings, _ = validate_scene_plan.validate_semantics(
        bundled_example, 0.75, True
    )
    _check(
        not example_schema_errors and not example_semantic_errors and not example_semantic_warnings,
        "bundled-scene-plan-example",
        checks,
        example=str(example_path.resolve()),
        schema_errors=example_schema_errors,
        semantic_errors=example_semantic_errors,
        semantic_warnings=example_semantic_warnings,
    )
    valid_plan = _valid_plan(seed_sha, seed_width, seed_height)
    schema_errors = validator.validate(valid_plan)
    semantic_errors, semantic_warnings, _ = validate_scene_plan.validate_semantics(valid_plan, 0.75, False)
    _check(not schema_errors, "valid-plan-schema", checks, errors=schema_errors)
    _check(not semantic_errors and not semantic_warnings, "valid-plan-semantics", checks, errors=semantic_errors, warnings=semantic_warnings)
    _test_panel_crop(validator, checks)

    source_report, source_errors, _ = validate_scene_plan.verify_source(
        valid_plan,
        Path("scene-plan.json").resolve(),
        None,
        None,
        seed_sha,
        False,
    )
    _check(source_report.get("status") == "pass" and not source_errors, "source-hash-match", checks, report=source_report, errors=source_errors)

    # Exercise the documented end-to-end CLI, including --source-manifest and
    # --output.  Only short-lived JSON files are written under a temp directory.
    if args.sources and source_results and source_results[0]["ok"]:
        with tempfile.TemporaryDirectory(prefix="sci-diagram-pptx-test-") as temporary_dir:
            temporary = Path(temporary_dir)
            manifest_path = temporary / "source-manifest.json"
            plan_path = temporary / "scene-plan.json"
            report_path = temporary / "scene-plan-validation.json"
            preflight_rc = source_preflight.main([str(args.sources[0]), "--output", str(manifest_path)])
            plan_path.write_text(json.dumps(valid_plan), encoding="utf-8")
            validation_rc = validate_scene_plan.main(
                [
                    str(plan_path),
                    "--source-manifest",
                    str(manifest_path),
                    "--low-confidence-as-error",
                    "--output",
                    str(report_path),
                ]
            )
            cli_report = json.loads(report_path.read_text(encoding="utf-8")) if report_path.exists() else {}
            _check(
                preflight_rc == 0
                and validation_rc == 0
                and cli_report.get("ok") is True
                and cli_report.get("status") == "PASS"
                and cli_report.get("tool") == "sci-diagram-pptx/validate_scene_plan.py"
                and cli_report.get("tool_version") == "1.0"
                and cli_report.get("validator_version") == "1.0"
                and cli_report.get("hard_failures") == []
                and isinstance(cli_report.get("scene_plan_sha256"), str)
                and len(cli_report["scene_plan_sha256"]) == 64
                and bool(cli_report.get("checks"))
                and all(item.get("status") == "PASS" for item in cli_report.get("checks", [])),
                "documented-cli-roundtrip",
                checks,
                preflight_exit=preflight_rc,
                validation_exit=validation_rc,
                status=cli_report.get("status"),
                tool=cli_report.get("tool"),
                tool_version=cli_report.get("tool_version"),
                hard_failure_count=len(cli_report.get("hard_failures", [])),
                check_count=len(cli_report.get("checks", [])),
                scene_plan_sha256=cli_report.get("scene_plan_sha256"),
                source_match=cli_report.get("source_match"),
                errors=cli_report.get("errors"),
            )

    invalid_plan = copy.deepcopy(valid_plan)
    invalid_plan["objects"][1]["id"] = "node-a"
    invalid_plan["objects"][1]["bbox"] = {"x": 0.9, "y": 0.2, "width": 0.25, "height": 0.2}
    invalid_plan["objects"][1]["z_order"] = 1
    invalid_plan["objects"][1]["confidence"] = 1.5
    invalid_plan["objects"][1]["reconstruction"] = {"mode": "isolated-raster", "expected_ooxml_kind": "p:pic"}
    invalid_plan["connections"][0]["to"]["object_id"] = "missing-node"
    invalid_schema_errors = validator.validate(invalid_plan)
    invalid_semantic_errors, _, _ = validate_scene_plan.validate_semantics(invalid_plan, 0.75, False)
    error_codes = {item["code"] for item in invalid_schema_errors + invalid_semantic_errors}
    expected_codes = {
        "schema.maximum",
        "object.id_duplicate",
        "bbox.out_of_bounds_x",
        "z_order.duplicate",
        "connection.object_missing",
        "degradation.approval_required",
    }
    _check(expected_codes.issubset(error_codes), "invalid-plan-rejected", checks, expected=sorted(expected_codes), actual=sorted(error_codes))

    missing_contract = copy.deepcopy(valid_plan)
    del missing_contract["kind"]
    del missing_contract["objects"][0]["source_bbox"]
    del missing_contract["objects"][0]["content"]
    del missing_contract["objects"][1]["style"]["shape_geometry"]
    del missing_contract["objects"][1]["reconstruction"]["expected_ooxml_kind"]
    del missing_contract["connections"][0]["source_route"]
    del missing_contract["connections"][0]["style"]["arrow_at"]
    del missing_contract["connections"][0]["reconstruction"]["expected_ooxml_kind"]
    missing_errors = validator.validate(missing_contract)
    missing_paths = {item["path"] for item in missing_errors if item["code"] == "schema.required"}
    _check(
        {
            "/kind",
            "/objects/0/source_bbox",
            "/objects/0/content",
            "/objects/1/style/shape_geometry",
            "/objects/1/reconstruction/expected_ooxml_kind",
            "/connections/0/source_route",
            "/connections/0/style/arrow_at",
            "/connections/0/reconstruction/expected_ooxml_kind",
        }.issubset(missing_paths),
        "required-contract-fields",
        checks,
        actual=sorted(missing_paths),
    )
    missing_semantic_errors, _, _ = validate_scene_plan.validate_semantics(missing_contract, 0.75, True)
    missing_semantic_codes = {item["code"] for item in missing_semantic_errors}
    _check(
        {"object.content_contract", "object.style_contract", "connection.style_contract"}.issubset(missing_semantic_codes),
        "semantic-content-style-contracts",
        checks,
        actual=sorted(missing_semantic_codes),
    )

    partial_panel = copy.deepcopy(valid_plan)
    partial_panel["source"]["panel_label"] = "panel A"
    panel_errors, _, _ = validate_scene_plan.validate_semantics(partial_panel, 0.75, True)
    panel_codes = {item["code"] for item in panel_errors}
    _check(
        {
            "source.panel_parent_hash_required",
            "source.panel_parent_width_required",
            "source.panel_parent_height_required",
            "source.panel_manifest_hash_required",
            "source.panel_bbox_required",
            "source.panel_selection_required",
        }.issubset(panel_codes),
        "partial-panel-metadata-rejected",
        checks,
        actual=sorted(panel_codes),
    )

    valid_panel = copy.deepcopy(valid_plan)
    valid_panel["source"].update(
        {
            "parent_sha256": "c" * 64,
            "parent_width_px": seed_width + 400,
            "parent_height_px": seed_height + 500,
            "panel_manifest_sha256": "d" * 64,
            "panel_bbox": {"x": 100, "y": 200, "width": seed_width, "height": seed_height},
            "panel_label": "A",
            "panel_selection": {
                "selection_source": "user-explicit",
                "selected_by": "user",
                "selected_at": "2026-08-08T10:25:00Z",
                "evidence": "The user explicitly selected panel A in the active conversation.",
            },
        }
    )
    valid_panel_schema_errors = validator.validate(valid_panel)
    valid_panel_semantic_errors, _, _ = validate_scene_plan.validate_semantics(valid_panel, 0.75, True)
    _check(
        not valid_panel_schema_errors and not valid_panel_semantic_errors,
        "complete-panel-lineage-accepted",
        checks,
        schema_errors=valid_panel_schema_errors,
        semantic_errors=valid_panel_semantic_errors,
    )

    out_of_parent = copy.deepcopy(valid_panel)
    out_of_parent["source"]["panel_bbox"]["x"] = out_of_parent["source"]["parent_width_px"] - seed_width + 1
    out_of_parent["source"]["panel_bbox"]["y"] = out_of_parent["source"]["parent_height_px"] - seed_height + 1
    out_of_parent_errors, _, _ = validate_scene_plan.validate_semantics(out_of_parent, 0.75, True)
    out_of_parent_codes = {item["code"] for item in out_of_parent_errors}
    _check(
        {"source.panel_bbox_out_of_parent_x", "source.panel_bbox_out_of_parent_y"}.issubset(out_of_parent_codes),
        "panel-bounds-contained-by-parent",
        checks,
        actual=sorted(out_of_parent_codes),
    )

    agent_selected_panel = copy.deepcopy(valid_panel)
    agent_selected_panel["source"]["panel_selection"]["selection_source"] = "agent-inferred"
    agent_selection_schema_errors = validator.validate(agent_selected_panel)
    agent_selection_semantic_errors, _, _ = validate_scene_plan.validate_semantics(agent_selected_panel, 0.75, True)
    _check(
        any(item["code"] == "schema.const" for item in agent_selection_schema_errors)
        and any(item["code"] == "source.panel_selection_source" for item in agent_selection_semantic_errors),
        "agent-inferred-panel-selection-rejected",
        checks,
        schema_errors=agent_selection_schema_errors,
        semantic_errors=agent_selection_semantic_errors,
    )

    approved_degradation = copy.deepcopy(valid_plan)
    approved_degradation["objects"][0]["reconstruction"] = {
        "mode": "isolated-raster",
        "expected_ooxml_kind": "p:pic",
        "degradation_id": "deg-raster-a",
        "rationale": "intrinsically raster scientific inset",
    }
    approved_degradation["approvals"]["degradations"] = [
        {
            "id": "deg-raster-a",
            "entity_ids": ["node-a"],
            "reason": "intrinsically raster scientific inset",
            "approved": True,
            "approved_by": "user",
            "approved_at": "2026-08-08T10:30:00Z",
            "evidence": "Explicit approval recorded in the active conversation.",
            "approval_source": "user-explicit",
        }
    ]
    approved_schema_errors = validator.validate(approved_degradation)
    approved_semantic_errors, _, _ = validate_scene_plan.validate_semantics(approved_degradation, 0.75, True)
    _check(
        not approved_schema_errors and not approved_semantic_errors,
        "explicit-degradation-approval-accepted",
        checks,
        schema_errors=approved_schema_errors,
        semantic_errors=approved_semantic_errors,
    )
    approved_connection = copy.deepcopy(valid_plan)
    approved_connection["connections"][0]["reconstruction"] = {
        "mode": "native-approximation",
        "expected_ooxml_kind": "p:sp",
        "degradation_id": "deg-edge-line",
        "rationale": "attachment cannot round-trip in the target runtime",
    }
    approved_connection["approvals"]["degradations"] = [
        {
            "id": "deg-edge-line",
            "entity_ids": ["edge-a-b"],
            "reason": "connector attachment becomes a native line",
            "approved": True,
            "approved_by": "user",
            "approved_at": "2026-08-08T10:31:00Z",
            "evidence": "Explicit approval recorded in the active conversation.",
            "approval_source": "user-explicit",
        }
    ]
    connection_schema_errors = validator.validate(approved_connection)
    connection_semantic_errors, _, _ = validate_scene_plan.validate_semantics(approved_connection, 0.75, True)
    _check(
        not connection_schema_errors and not connection_semantic_errors,
        "connection-degradation-approval-accepted",
        checks,
        schema_errors=connection_schema_errors,
        semantic_errors=connection_semantic_errors,
    )
    missing_evidence = copy.deepcopy(approved_degradation)
    del missing_evidence["approvals"]["degradations"][0]["evidence"]
    evidence_errors = validator.validate(missing_evidence)
    _check(
        any(item["code"] == "schema.required" and item["path"].endswith("/evidence") for item in evidence_errors),
        "degradation-evidence-required",
        checks,
        errors=evidence_errors,
    )

    low_confidence = copy.deepcopy(valid_plan)
    low_confidence["objects"][0]["confidence"] = 0.5
    low_errors, _, _ = validate_scene_plan.validate_semantics(low_confidence, 0.75, True)
    _check(
        any(item["code"] == "confidence.below_threshold" for item in low_errors),
        "low-confidence-default-hard-failure",
        checks,
        errors=low_errors,
    )

    mismatch_report, mismatch_errors, _ = validate_scene_plan.verify_source(
        valid_plan,
        Path("scene-plan.json").resolve(),
        None,
        None,
        "b" * 64,
        False,
    )
    _check(mismatch_report.get("status") == "fail" and any(item["code"] == "source.hash_mismatch" for item in mismatch_errors), "source-hash-mismatch-rejected", checks)

    passed = all(item["pass"] for item in checks)
    payload = {
        "ok": passed,
        "kind": "sci-diagram-pptx-preflight-tests",
        "source_results": source_results,
        "checks": checks,
        "summary": {"passed": sum(item["pass"] for item in checks), "total": len(checks)},
    }
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    return 0 if passed else 1


if __name__ == "__main__":
    raise SystemExit(main())
