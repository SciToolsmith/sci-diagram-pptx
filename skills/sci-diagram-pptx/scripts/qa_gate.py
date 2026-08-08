#!/usr/bin/env python3
"""Aggregate sci-diagram-pptx QA reports and enforce hard gates."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable


TOOL_VERSION = "1.0"
VALID_STATUS = {"PASS", "WARN", "FAIL"}
VALID_SEVERITY = {"HARD", "SOFT", "INFO"}
SHA256_RE = re.compile(r"^[0-9a-fA-F]{64}$")
EXPECTED_REPORT_KINDS = {
    "scene_plan_report": "sci-diagram-pptx-scene-plan-validation",
    "pptx_audit": "sci-diagram-pptx-pptx-audit",
    "visual_report": "sci-diagram-pptx-visual-report",
    "manual_review_attestation": "sci-diagram-pptx-manual-review-attestation",
    "overflow_report": "sci-diagram-pptx-overflow-report",
    "render_report": "sci-diagram-pptx-render-report",
}
REPORT_CONTRACTS: dict[str, dict[str, Any]] = {
    "scene_plan_report": {
        "tool": "sci-diagram-pptx/validate_scene_plan.py",
        "version": "1.0",
        "checks": {
            "schema-load", "scene-plan-load", "json-schema-validation", "scene-kind-and-panel-contract",
            "unique-entity-ids", "normalized-source-and-slide-bboxes", "deterministic-z-order",
            "connection-references", "confidence-threshold", "scene-entity-reconstruction-contract",
            "explicit-degradation-approval", "source-identity-match", "panel-lineage-manifest",
        },
    },
    "pptx_audit": {
        "tool": "sci-diagram-pptx/audit_pptx.py",
        "version": "1.0",
        "checks": {
            "inputs.source_manifest", "inputs.source_manifest_kind", "inputs.scene_plan", "inputs.scene_plan_kind",
            "inputs.source_identity", "package.readable", "presentation.slide_size", "presentation.slide_count",
            "presentation.visible_slides", "package.orphan_slide_parts", "package.no_external_relationships",
            "package.no_macros", "package.no_ole_or_embedded_packages", "package.no_svg_emf_wmf_shortcuts",
            "slide1.no_full_slide_image_or_image_background", "slide1.no_page_sized_background_shapes",
            "slide1.no_hidden_or_offcanvas_source_images", "slide1.raster_objects_against_plan",
            "slide1.image_bearers_match_isolated_raster_plan",
            "slide1.has_native_objects", "plan.no_unsupported_objects", "plan.reconstruction_modes",
            "slide1.expected_ooxml_kind_counts", "slide1.named_object_ooxml_kinds",
            "slide1.named_object_content_and_bbox", "slide1.named_object_shape_geometry", "slide1.connection_semantics",
            "slide1.math_source_semantics",
            "slide1.connectors_against_plan", "slide1.connector_attachment", "slide1.text_against_plan",
            "slide1.omml_equations", "presentation.relationship_parse",
            "slide2.reference_image_relationship", "slide2.reference_source_identity", "slide2.reference_no_crop",
            "slide2.reference_visible", "slide2.reference_contain_center", "slide2.reference_only_content",
            "presentation.aspect_ratio", "plan.degradations_approved",
        },
    },
    "visual_report": {
        "tool": "sci-diagram-pptx/compare_render.py",
        "version": "1.0",
        "checks": {
            "visual.inputs_readable", "visual.single_frame_inputs", "visual.aspect_ratio", "visual.alignment_confidence", "visual.alignment_magnitude",
            "visual.heatmap_written", "visual.mae_threshold", "visual.mismatch_ratio_threshold",
            "visual.edge_recall", "visual.foreground_recall", "visual.blank_baseline_guard",
            "visual.render_not_blank",
        },
    },
    "overflow_report": {
        "tool": "sci-diagram-pptx/overflow_check.py",
        "version": "1.0",
        "checks": {"overflow.inputs_readable", "overflow.helper_exit", "overflow.explicit_result"},
    },
    "render_report": {
        "tool": "sci-diagram-pptx/render_evidence.py",
        "version": "1.0",
        "checks": {"render.inputs_readable", "render.fresh_output_directory", "render.helper_exit", "render.explicit_result", "render.slide_files"},
    },
}


class GateError(RuntimeError):
    """An input or output error that prevents aggregation."""


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def load_report(path: Path, label: str) -> tuple[dict[str, Any], dict[str, Any]]:
    resolved = path.expanduser().resolve()
    try:
        raw = resolved.read_text(encoding="utf-8")
        value = json.loads(raw)
    except (OSError, json.JSONDecodeError) as exc:
        raise GateError(f"Cannot read {label} JSON '{resolved}': {exc}") from exc
    if not isinstance(value, dict):
        raise GateError(f"{label} must contain one JSON object")
    provenance = {"label": label, "path": str(resolved), "sha256": sha256_file(resolved)}
    return value, provenance


def write_json(payload: dict[str, Any], output: Path) -> None:
    target = output.expanduser().resolve()
    target.parent.mkdir(parents=True, exist_ok=True)
    temporary = target.with_name(target.name + ".tmp")
    temporary.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    os.replace(temporary, target)


def status_of(value: Any) -> str | None:
    if not isinstance(value, str):
        return None
    normalized = value.strip().upper()
    aliases = {"OK": "PASS", "PASSED": "PASS", "WARNING": "WARN", "FAILED": "FAIL", "ERROR": "FAIL"}
    normalized = aliases.get(normalized, normalized)
    return normalized if normalized in VALID_STATUS else None


def severity_of(value: Any) -> str:
    if not isinstance(value, str):
        return "INFO"
    normalized = value.strip().upper()
    return normalized if normalized in VALID_SEVERITY else "INFO"


def message_of(item: Any, fallback: str) -> str:
    if isinstance(item, str):
        return item
    if isinstance(item, dict):
        for key in ("message", "detail", "reason", "code", "id"):
            if isinstance(item.get(key), str) and item[key]:
                return item[key]
    return fallback


def id_of(item: Any, fallback: str) -> str:
    if isinstance(item, dict):
        for key in ("id", "code"):
            if isinstance(item.get(key), str) and item[key]:
                return item[key]
    return fallback


def normalize_report(label: str, report: dict[str, Any]) -> dict[str, Any]:
    """Normalize one report while preserving uncertainty as explicit WARN."""
    checks: list[dict[str, Any]] = []
    hard_failures: list[dict[str, Any]] = []
    warnings: list[dict[str, Any]] = []
    schema_missing: list[str] = []
    for key in ("status", "hard_failures", "warnings", "checks"):
        if key not in report:
            schema_missing.append(key)

    raw_status = status_of(report.get("status"))
    if raw_status is None:
        # Compatibility with earlier source/validator reports is intentionally
        # one-way: a negative `ok` proves failure, but positive `ok` alone does
        # not prove all current gates passed.
        if report.get("ok") is False:
            raw_status = "FAIL"
        else:
            raw_status = "WARN"
            warnings.append({
                "source": label,
                "id": "report.status_unrecognized",
                "message": "Report has no recognized PASS/WARN/FAIL status; it was not treated as PASS",
                "severity": "HARD",
            })

    raw_checks = report.get("checks")
    if isinstance(raw_checks, list):
        for index, item in enumerate(raw_checks):
            if not isinstance(item, dict):
                checks.append({
                    "source": label,
                    "id": f"report.malformed_check.{index}",
                    "status": "WARN",
                    "severity": "HARD",
                    "message": "A report check is not an object and could not be evaluated",
                })
                continue
            check_status = status_of(item.get("status"))
            if check_status is None:
                check_status = "WARN"
                check_message = message_of(item, "Check status is absent or unrecognized; it was not treated as PASS")
            else:
                check_message = message_of(item, "No check message supplied")
            check_severity = severity_of(item.get("severity")) if "severity" in item else ("HARD" if check_status in {"WARN", "FAIL"} else "INFO")
            normalized = {
                "source": label,
                "id": id_of(item, f"check.{index}"),
                "status": check_status,
                "severity": check_severity,
                "message": check_message,
            }
            if "evidence" in item:
                normalized["evidence"] = item["evidence"]
            checks.append(normalized)
            if normalized["status"] == "FAIL" and normalized["severity"] == "HARD":
                hard_failures.append({key: normalized[key] for key in ("source", "id", "message", "severity")})
            elif normalized["status"] == "WARN" or (normalized["status"] == "FAIL" and normalized["severity"] != "HARD"):
                warnings.append({key: normalized[key] for key in ("source", "id", "message", "severity")})
    else:
        warnings.append({
            "source": label,
            "id": "report.checks_unavailable",
            "message": "Report checks are absent or malformed; gate coverage is unproven",
            "severity": "HARD",
        })

    raw_hard = report.get("hard_failures")
    if isinstance(raw_hard, list):
        for index, item in enumerate(raw_hard):
            hard_failures.append({
                "source": label,
                "id": id_of(item, f"reported_hard_failure.{index}"),
                "message": message_of(item, "Upstream report declared a hard failure"),
                "severity": "HARD",
            })
    elif raw_hard is not None:
        warnings.append({
            "source": label,
            "id": "report.hard_failures_malformed",
            "message": "Report hard_failures is not an array; failure coverage is unproven",
            "severity": "HARD",
        })

    # Support the legacy source_preflight `errors` array.  Errors prove failure;
    # their absence does not prove PASS.
    raw_errors = report.get("errors")
    if isinstance(raw_errors, list):
        for index, item in enumerate(raw_errors):
            hard_failures.append({
                "source": label,
                "id": id_of(item, f"reported_error.{index}"),
                "message": message_of(item, "Upstream report declared an error"),
                "severity": "HARD",
            })

    raw_warnings = report.get("warnings")
    if isinstance(raw_warnings, list):
        for index, item in enumerate(raw_warnings):
            warnings.append({
                "source": label,
                "id": id_of(item, f"reported_warning.{index}"),
                "message": message_of(item, "Upstream report declared a warning"),
                "severity": severity_of(item.get("severity")) if isinstance(item, dict) and "severity" in item else "HARD",
            })
    elif raw_warnings is not None:
        warnings.append({
            "source": label,
            "id": "report.warnings_malformed",
            "message": "Report warnings is not an array",
            "severity": "INFO",
        })

    if raw_status == "FAIL" and not hard_failures:
        hard_failures.append({
            "source": label,
            "id": "report.status_fail",
            "message": "Upstream report status is FAIL",
            "severity": "HARD",
        })
    if raw_status == "WARN" and not warnings:
        warnings.append({
            "source": label,
            "id": "report.status_warn",
            "message": "Upstream report status is WARN",
            "severity": "HARD",
        })
    if schema_missing:
        warnings.append({
            "source": label,
            "id": "report.schema_incomplete",
            "message": "Report does not expose every unified aggregation field",
            "severity": "HARD",
            "missing_fields": schema_missing,
        })

    # A HARD warning represents unresolved evidence.  Preserve its WARN status in
    # checks/warnings, but also make it a delivery-blocking hard failure.
    for warning in warnings:
        if warning.get("severity") == "HARD":
            hard_failures.append({
                "source": warning.get("source", label),
                "id": warning.get("id", "hard_warning"),
                "message": f"Unresolved HARD warning: {warning.get('message', 'warning')}",
                "severity": "HARD",
            })

    return {
        "status": raw_status,
        "checks": checks,
        "hard_failures": deduplicate_issues(hard_failures),
        "warnings": deduplicate_issues(warnings),
        "schema_missing": schema_missing,
        "kind": report.get("kind"),
    }


def normalize_attestation(label: str, report: dict[str, Any]) -> dict[str, Any]:
    """Validate the dedicated manual-review schema and hash every evidence file."""
    checks: list[dict[str, Any]] = []
    hard_failures: list[dict[str, Any]] = []

    def check(check_id: str, passed: bool, message_pass: str, message_fail: str, **evidence: Any) -> None:
        item: dict[str, Any] = {
            "source": label,
            "id": check_id,
            "status": "PASS" if passed else "FAIL",
            "severity": "HARD",
            "message": message_pass if passed else message_fail,
        }
        if evidence:
            item["evidence"] = evidence
        checks.append(item)
        if not passed:
            hard_failures.append({"source": label, "id": check_id, "message": item["message"], "severity": "HARD"})

    check("attestation.status", report.get("status") == "PASS", "Attestation status is PASS", "Attestation status must be exactly PASS", actual=report.get("status"))
    check(
        "attestation.identity_fields",
        valid_sha(report.get("source_sha256")) is not None and valid_sha(report.get("pptx_sha256")) is not None,
        "Attestation contains valid source and PPTX SHA-256 fields",
        "Attestation source_sha256 and pptx_sha256 must be 64 hexadecimal characters",
    )
    check(
        "attestation.review_flags",
        report.get("full_size_visual_review") is True and report.get("overflow_check_passed") is True,
        "Attestation confirms full-size review and a passing overflow check",
        "Attestation must set full_size_visual_review and overflow_check_passed to true",
        full_size_visual_review=report.get("full_size_visual_review"),
        overflow_check_passed=report.get("overflow_check_passed"),
    )

    reviewer = report.get("reviewer")
    reviewed_at = report.get("reviewed_at")
    reviewed_at_valid = False
    if isinstance(reviewed_at, str) and reviewed_at.strip():
        try:
            parsed = datetime.fromisoformat(reviewed_at.replace("Z", "+00:00"))
            reviewed_at_valid = parsed.tzinfo is not None
        except ValueError:
            reviewed_at_valid = False
    check(
        "attestation.reviewer_and_time",
        isinstance(reviewer, str) and bool(reviewer.strip()) and reviewed_at_valid,
        "Attestation identifies the reviewer and a timezone-aware review timestamp",
        "Attestation requires a non-empty reviewer and timezone-aware ISO-8601 reviewed_at",
        reviewer=reviewer,
        reviewed_at=reviewed_at,
    )

    evidence = report.get("evidence") if isinstance(report.get("evidence"), dict) else {}
    evidence_hashes: dict[str, str | None] = {}
    for evidence_name in ("slide_1_render", "slide_2_render", "overflow_report"):
        record = evidence.get(evidence_name) if isinstance(evidence.get(evidence_name), dict) else {}
        raw_path = record.get("path")
        declared_sha = valid_sha(record.get("sha256"))
        resolved: Path | None = None
        actual_sha: str | None = None
        path_ok = isinstance(raw_path, str) and bool(raw_path) and Path(raw_path).expanduser().is_absolute()
        if path_ok:
            resolved = Path(raw_path).expanduser().resolve()
            if resolved.is_file():
                try:
                    actual_sha = sha256_file(resolved)
                except OSError:
                    actual_sha = None
        evidence_hashes[evidence_name] = actual_sha
        passed = path_ok and resolved is not None and resolved.is_file() and declared_sha is not None and actual_sha == declared_sha
        overflow_details: dict[str, Any] = {}
        if evidence_name == "overflow_report":
            exit_ok = isinstance(record.get("exit_code"), int) and not isinstance(record.get("exit_code"), bool) and record.get("exit_code") == 0
            overflow_payload: dict[str, Any] | None = None
            if resolved is not None and resolved.is_file():
                try:
                    candidate = json.loads(resolved.read_text(encoding="utf-8"))
                    overflow_payload = candidate if isinstance(candidate, dict) else None
                except (OSError, json.JSONDecodeError, UnicodeDecodeError):
                    overflow_payload = None
            overflow_schema_ok = bool(
                overflow_payload
                and overflow_payload.get("kind") == "sci-diagram-pptx-overflow-report"
                and overflow_payload.get("schema_version") == "1.0"
                and overflow_payload.get("status") == "PASS"
                and valid_sha(overflow_payload.get("pptx_sha256")) == valid_sha(report.get("pptx_sha256"))
            )
            overflow_details = {
                "exit_code_ok": exit_ok,
                "overflow_schema_ok": overflow_schema_ok,
                "overflow_kind": overflow_payload.get("kind") if overflow_payload else None,
                "overflow_schema_version": overflow_payload.get("schema_version") if overflow_payload else None,
                "overflow_status": overflow_payload.get("status") if overflow_payload else None,
                "overflow_pptx_sha256": overflow_payload.get("pptx_sha256") if overflow_payload else None,
            }
            passed = passed and exit_ok and overflow_schema_ok
        check(
            f"attestation.evidence.{evidence_name}",
            bool(passed),
            f"{evidence_name} exists and its bytes match the attested SHA-256",
            f"{evidence_name} must be an absolute regular file whose SHA-256 matches the attestation" + ("; overflow JSON must have the supported kind/schema, PASS status, matching PPTX SHA-256, and exit_code 0" if evidence_name == "overflow_report" else ""),
            path=str(resolved) if resolved else raw_path,
            declared_sha256=declared_sha,
            actual_sha256=actual_sha,
            exit_code=record.get("exit_code") if evidence_name == "overflow_report" else None,
            overflow_validation=overflow_details if evidence_name == "overflow_report" else None,
        )

    return {
        "status": "PASS" if not hard_failures else "FAIL",
        "checks": checks,
        "hard_failures": deduplicate_issues(hard_failures),
        "warnings": [],
        "schema_missing": [],
        "kind": report.get("kind"),
        "evidence_hashes": evidence_hashes,
    }


def deduplicate_issues(items: list[dict[str, Any]]) -> list[dict[str, Any]]:
    seen: set[tuple[str, str, str]] = set()
    result: list[dict[str, Any]] = []
    for item in items:
        key = (str(item.get("source", "")), str(item.get("id", "")), str(item.get("message", "")))
        if key not in seen:
            seen.add(key)
            result.append(item)
    return result


def nested_value(value: dict[str, Any], *paths: tuple[str, ...]) -> Any:
    for path in paths:
        current: Any = value
        for key in path:
            if not isinstance(current, dict) or key not in current:
                break
            current = current[key]
        else:
            if current is not None:
                return current
    return None


def valid_sha(value: Any) -> str | None:
    return value.lower() if isinstance(value, str) and SHA256_RE.fullmatch(value) else None


def identity_sha(report: dict[str, Any], label: str, identity: str) -> str | None:
    paths: dict[tuple[str, str], tuple[tuple[str, ...], ...]] = {
        ("scene_plan_report", "source"): (("source_sha256",), ("identity", "source_sha256"), ("source_match", "actual_sha256"), ("source_match", "planned_sha256")),
        ("pptx_audit", "source"): (("input", "source_sha256"),),
        ("visual_report", "source"): (("input", "reference_sha256"), ("input", "reference", "sha256")),
        ("manual_review_attestation", "source"): (("source_sha256",), ("identity", "source_sha256")),
        ("scene_plan_report", "scene_plan"): (("scene_plan_sha256",), ("identity", "scene_plan_sha256"), ("input", "scene_plan_sha256")),
        ("pptx_audit", "scene_plan"): (("input", "scene_plan_sha256"),),
        ("pptx_audit", "pptx"): (("input", "pptx_sha256"),),
        ("manual_review_attestation", "pptx"): (("pptx_sha256",), ("identity", "pptx_sha256")),
        ("overflow_report", "pptx"): (("pptx_sha256",),),
        ("render_report", "pptx"): (("pptx_sha256",),),
    }
    candidates = paths.get((label, identity), ())
    return valid_sha(nested_value(report, *candidates))


def binding_check(
    check_id: str,
    identity: str,
    participants: list[str],
    reports: dict[str, dict[str, Any]],
) -> tuple[dict[str, Any], dict[str, Any] | None]:
    values = {label: identity_sha(reports[label], label, identity) for label in participants}
    missing = [label for label, value in values.items() if value is None]
    unique = sorted({value for value in values.values() if value is not None})
    if missing:
        status = "FAIL"
        message = f"{identity} SHA-256 is missing or invalid in required reports"
    elif len(unique) != 1:
        status = "FAIL"
        message = f"{identity} SHA-256 differs across required reports"
    else:
        status = "PASS"
        message = f"{identity} SHA-256 is consistently bound across required reports"
    check = {
        "source": "aggregate",
        "id": check_id,
        "status": status,
        "severity": "HARD",
        "message": message,
        "evidence": {"values": values, "missing": missing},
    }
    failure = None if status == "PASS" else {"source": "aggregate", "id": check_id, "message": message, "severity": "HARD"}
    return check, failure


def aggregate_reports(inputs: list[tuple[str, Path]]) -> dict[str, Any]:
    all_checks: list[dict[str, Any]] = []
    hard_failures: list[dict[str, Any]] = []
    warnings: list[dict[str, Any]] = []
    provenance: list[dict[str, Any]] = []
    report_summaries: dict[str, Any] = {}
    raw_reports: dict[str, dict[str, Any]] = {}

    for label, path in inputs:
        report, source = load_report(path, label)
        raw_reports[label] = report
        provenance.append(source)
        normalized = normalize_attestation(label, report) if label == "manual_review_attestation" else normalize_report(label, report)
        report_summaries[label] = {
            "status": normalized["status"],
            "kind": normalized["kind"],
            "check_count": len(normalized["checks"]),
            "hard_failure_count": len(normalized["hard_failures"]),
            "warning_count": len(normalized["warnings"]),
            "schema_missing": normalized["schema_missing"],
        }
        if label == "manual_review_attestation":
            report_summaries[label]["evidence_hashes"] = normalized.get("evidence_hashes", {})
        all_checks.extend(normalized["checks"])
        hard_failures.extend(normalized["hard_failures"])
        warnings.extend(normalized["warnings"])
        all_checks.append({
            "source": "aggregate",
            "id": f"aggregate.{label}.status",
            "status": normalized["status"],
            "severity": "HARD" if normalized["status"] == "FAIL" else "INFO",
            "message": f"{label} upstream status is {normalized['status']}",
        })

        expected_kind = EXPECTED_REPORT_KINDS.get(label)
        actual_kind = report.get("kind")
        kind_ok = actual_kind == expected_kind
        kind_check = {
            "source": "aggregate",
            "id": f"aggregate.{label}.kind",
            "status": "PASS" if kind_ok else "FAIL",
            "severity": "HARD",
            "message": f"{label} kind is recognized" if kind_ok else f"{label} kind is absent or unexpected",
            "evidence": {"expected": expected_kind, "actual": actual_kind},
        }
        all_checks.append(kind_check)
        if not kind_ok:
            hard_failures.append({"source": "aggregate", "id": kind_check["id"], "message": kind_check["message"], "severity": "HARD"})

        contract = REPORT_CONTRACTS.get(label)
        if contract:
            actual_tool = report.get("tool")
            actual_version = report.get("tool_version") or report.get("validator_version")
            producer_ok = actual_tool == contract["tool"] and actual_version == contract["version"]
            producer_check = {
                "source": "aggregate",
                "id": f"aggregate.{label}.producer",
                "status": "PASS" if producer_ok else "FAIL",
                "severity": "HARD",
                "message": f"{label} producer and version are supported" if producer_ok else f"{label} producer or version is absent/unsupported",
                "evidence": {"expected_tool": contract["tool"], "actual_tool": actual_tool, "expected_version": contract["version"], "actual_version": actual_version},
            }
            all_checks.append(producer_check)
            if not producer_ok:
                hard_failures.append({"source": "aggregate", "id": producer_check["id"], "message": producer_check["message"], "severity": "HARD"})

            check_map: dict[str, list[str]] = {}
            for child_check in normalized["checks"]:
                check_map.setdefault(str(child_check.get("id")), []).append(str(child_check.get("status")))
            missing_checks = sorted(check_id for check_id in contract["checks"] if check_id not in check_map)
            nonpassing_checks = {
                check_id: statuses
                for check_id, statuses in check_map.items()
                if check_id in contract["checks"] and (len(statuses) != 1 or statuses[0] != "PASS")
            }
            coverage_ok = not missing_checks and not nonpassing_checks
            coverage_check = {
                "source": "aggregate",
                "id": f"aggregate.{label}.mandatory_checks",
                "status": "PASS" if coverage_ok else "FAIL",
                "severity": "HARD",
                "message": f"{label} contains one passing instance of every mandatory check" if coverage_ok else f"{label} mandatory gate coverage is missing, duplicated, or non-passing",
                "evidence": {"required_count": len(contract["checks"]), "missing": missing_checks, "nonpassing_or_duplicate": nonpassing_checks},
            }
            all_checks.append(coverage_check)
            if not coverage_ok:
                hard_failures.append({"source": "aggregate", "id": coverage_check["id"], "message": coverage_check["message"], "severity": "HARD"})

    attestation = raw_reports.get("manual_review_attestation", {})
    review = attestation.get("review") if isinstance(attestation.get("review"), dict) else attestation
    full_size_review = review.get("full_size_visual_review") is True
    overflow_passed = review.get("overflow_check_passed") is True
    attestation_check = {
        "source": "aggregate",
        "id": "aggregate.manual_review_evidence",
        "status": "PASS" if full_size_review and overflow_passed else "FAIL",
        "severity": "HARD",
        "message": "Manual full-size visual review and overflow check are attested" if full_size_review and overflow_passed else "Manual review attestation must set full_size_visual_review and overflow_check_passed to true",
        "evidence": {"full_size_visual_review": full_size_review, "overflow_check_passed": overflow_passed},
    }
    all_checks.append(attestation_check)
    if attestation_check["status"] == "FAIL":
        hard_failures.append({"source": "aggregate", "id": attestation_check["id"], "message": attestation_check["message"], "severity": "HARD"})

    bindings: dict[str, Any] = {}
    binding_specs = (
        ("aggregate.source_identity_binding", "source", ["scene_plan_report", "pptx_audit", "visual_report", "manual_review_attestation"]),
        ("aggregate.scene_plan_identity_binding", "scene_plan", ["scene_plan_report", "pptx_audit"]),
        ("aggregate.pptx_identity_binding", "pptx", ["pptx_audit", "manual_review_attestation", "overflow_report", "render_report"]),
    )
    for check_id, identity, participants in binding_specs:
        check, failure = binding_check(check_id, identity, participants, raw_reports)
        all_checks.append(check)
        bindings[identity] = check["evidence"]["values"]
        if failure:
            hard_failures.append(failure)

    visual_render_sha = valid_sha(nested_value(raw_reports["visual_report"], ("input", "render_sha256"), ("input", "render", "sha256")))
    attested_slide1_sha = valid_sha(nested_value(raw_reports["manual_review_attestation"], ("evidence", "slide_1_render", "sha256")))
    attested_slide2_sha = valid_sha(nested_value(raw_reports["manual_review_attestation"], ("evidence", "slide_2_render", "sha256")))
    render_rows = raw_reports["render_report"].get("renders") if isinstance(raw_reports["render_report"].get("renders"), list) else []
    render_by_slide: dict[int, list[str]] = {}
    for row in render_rows:
        if isinstance(row, dict) and isinstance(row.get("slide_index"), int):
            digest = valid_sha(row.get("sha256"))
            if digest:
                render_by_slide.setdefault(row["slide_index"], []).append(digest)
    rendered_slide1_sha = render_by_slide.get(1, [None])[0] if len(render_by_slide.get(1, [])) == 1 else None
    rendered_slide2_sha = render_by_slide.get(2, [None])[0] if len(render_by_slide.get(2, [])) == 1 else None
    render_binding_ok = (
        visual_render_sha is not None
        and attested_slide1_sha is not None
        and rendered_slide1_sha is not None
        and visual_render_sha == attested_slide1_sha == rendered_slide1_sha
    )
    render_binding_check = {
        "source": "aggregate",
        "id": "aggregate.slide1_render_identity_binding",
        "status": "PASS" if render_binding_ok else "FAIL",
        "severity": "HARD",
        "message": "Render producer, visual comparison, and manual review use the same Slide 1 bytes" if render_binding_ok else "Slide 1 SHA-256 is missing/duplicated or differs across render producer, visual report, and attestation",
        "evidence": {"render_report_slide_1_sha256": rendered_slide1_sha, "visual_report_render_sha256": visual_render_sha, "attested_slide_1_render_sha256": attested_slide1_sha},
    }
    all_checks.append(render_binding_check)
    if not render_binding_ok:
        hard_failures.append({"source": "aggregate", "id": render_binding_check["id"], "message": render_binding_check["message"], "severity": "HARD"})

    slide2_binding_ok = rendered_slide2_sha is not None and attested_slide2_sha is not None and rendered_slide2_sha == attested_slide2_sha
    slide2_binding_check = {
        "source": "aggregate",
        "id": "aggregate.slide2_render_identity_binding",
        "status": "PASS" if slide2_binding_ok else "FAIL",
        "severity": "HARD",
        "message": "Render producer and manual review use the same Slide 2 bytes" if slide2_binding_ok else "Slide 2 SHA-256 is missing/duplicated or differs between render producer and attestation",
        "evidence": {"render_report_slide_2_sha256": rendered_slide2_sha, "attested_slide_2_render_sha256": attested_slide2_sha},
    }
    all_checks.append(slide2_binding_check)
    if not slide2_binding_ok:
        hard_failures.append({"source": "aggregate", "id": slide2_binding_check["id"], "message": slide2_binding_check["message"], "severity": "HARD"})

    overflow_input = next((item for item in provenance if item["label"] == "overflow_report"), {})
    attested_overflow_path = nested_value(raw_reports["manual_review_attestation"], ("evidence", "overflow_report", "path"))
    attested_overflow_sha = valid_sha(nested_value(raw_reports["manual_review_attestation"], ("evidence", "overflow_report", "sha256")))
    overflow_path_matches = isinstance(attested_overflow_path, str) and str(Path(attested_overflow_path).expanduser().resolve()) == overflow_input.get("path")
    overflow_evidence_ok = overflow_path_matches and attested_overflow_sha is not None and attested_overflow_sha == overflow_input.get("sha256")
    overflow_evidence_check = {
        "source": "aggregate",
        "id": "aggregate.overflow_evidence_binding",
        "status": "PASS" if overflow_evidence_ok else "FAIL",
        "severity": "HARD",
        "message": "Attestation references the exact aggregated overflow report file" if overflow_evidence_ok else "Attestation overflow evidence path/hash does not match --overflow-report",
        "evidence": {"aggregated_path": overflow_input.get("path"), "aggregated_sha256": overflow_input.get("sha256"), "attested_path": attested_overflow_path, "attested_sha256": attested_overflow_sha},
    }
    all_checks.append(overflow_evidence_check)
    if not overflow_evidence_ok:
        hard_failures.append({"source": "aggregate", "id": overflow_evidence_check["id"], "message": overflow_evidence_check["message"], "severity": "HARD"})

    hard_failures = deduplicate_issues(hard_failures)
    warnings = deduplicate_issues(warnings)
    status = "FAIL" if hard_failures else "WARN" if warnings else "PASS"
    status_counts = {candidate: sum(1 for check in all_checks if check["status"] == candidate) for candidate in sorted(VALID_STATUS)}
    return {
        "schema_version": "1.0",
        "kind": "sci-diagram-pptx-qa-summary",
        "tool": "sci-diagram-pptx/qa_gate.py",
        "tool_version": TOOL_VERSION,
        "generated_at": utc_now(),
        "status": status,
        "delivery_authorization": not hard_failures,
        "hard_failures": hard_failures,
        "warnings": warnings,
        "checks": all_checks,
        "summary": {
            "report_count": len(inputs),
            "check_count": len(all_checks),
            "hard_failure_count": len(hard_failures),
            "warning_count": len(warnings),
            "check_status_counts": status_counts,
        },
        "reports": report_summaries,
        "inputs": provenance,
        "identity_bindings": bindings,
    }


class JsonArgumentParser(argparse.ArgumentParser):
    def error(self, message: str) -> None:
        print(json.dumps({"status": "FAIL", "hard_failures": [{"id": "cli.invalid_arguments", "message": message}], "warnings": [], "checks": []}, ensure_ascii=False), file=sys.stderr)
        raise SystemExit(2)


def build_parser() -> argparse.ArgumentParser:
    parser = JsonArgumentParser(description="Aggregate sci-diagram-pptx QA reports and enforce hard failures.")
    parser.add_argument("--scene-plan-report", type=Path, required=True, help="scene-plan validation JSON")
    parser.add_argument("--pptx-audit", type=Path, required=True, help="audit_pptx JSON")
    parser.add_argument("--visual-report", type=Path, required=True, help="compare_render report.json")
    parser.add_argument("--overflow-report", type=Path, required=True, help="overflow_check JSON")
    parser.add_argument("--render-report", type=Path, required=True, help="render_evidence JSON")
    parser.add_argument("--manual-review-attestation", type=Path, required=True, help="human/full-runtime review attestation JSON bound to source and PPTX hashes")
    parser.add_argument("--output", type=Path, required=True, help="write aggregate JSON here")
    return parser


def operational_failure(message: str) -> dict[str, Any]:
    check = {"source": "aggregate", "id": "aggregate.operational", "status": "FAIL", "severity": "HARD", "message": message}
    return {
        "schema_version": "1.0",
        "kind": "sci-diagram-pptx-qa-summary",
        "tool": "sci-diagram-pptx/qa_gate.py",
        "tool_version": TOOL_VERSION,
        "generated_at": utc_now(),
        "status": "FAIL",
        "delivery_authorization": False,
        "hard_failures": [{"source": "aggregate", "id": check["id"], "message": message, "severity": "HARD"}],
        "warnings": [],
        "checks": [check],
    }


def main(argv: Iterable[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        payload = aggregate_reports([
            ("scene_plan_report", args.scene_plan_report),
            ("pptx_audit", args.pptx_audit),
            ("visual_report", args.visual_report),
            ("overflow_report", args.overflow_report),
            ("render_report", args.render_report),
            ("manual_review_attestation", args.manual_review_attestation),
        ])
        write_json(payload, args.output)
        return 1 if payload["hard_failures"] else 0
    except (GateError, OSError) as exc:
        payload = operational_failure(str(exc))
        try:
            write_json(payload, args.output)
        except OSError:
            print(json.dumps(payload, ensure_ascii=False), file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
