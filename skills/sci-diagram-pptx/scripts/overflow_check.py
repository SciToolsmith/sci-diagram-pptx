#!/usr/bin/env python3
"""Fail-closed wrapper around the current Presentations slides_test.py helper."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable


TOOL_VERSION = "1.0"
PASS_SENTENCE = "Test passed. No overflow detected."


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


def build_report(pptx_path: Path, helper_path: Path, timeout: int, helper_python_path: Path | None = None) -> dict[str, Any]:
    pptx = pptx_path.expanduser().resolve()
    helper = helper_path.expanduser().resolve()
    helper_python = (helper_python_path or Path(sys.executable)).expanduser().resolve()
    checks: list[dict[str, Any]] = []

    def add(check_id: str, status: str, severity: str, message: str, **evidence: Any) -> None:
        item: dict[str, Any] = {"id": check_id, "status": status, "severity": severity, "message": message}
        if evidence:
            item["evidence"] = evidence
        checks.append(item)

    inputs_ok = pptx.is_file() and helper.is_file() and helper.suffix.lower() == ".py" and helper_python.is_file() and os.access(helper_python, os.X_OK)
    add("overflow.inputs_readable", "PASS" if inputs_ok else "FAIL", "HARD", "PPTX, slides_test helper, and helper Python are readable/executable" if inputs_ok else "PPTX, slides_test helper, or helper Python is missing/invalid", pptx=str(pptx), helper=str(helper), helper_python=str(helper_python))
    command = [str(helper_python), str(helper), str(pptx)]
    stdout = stderr = ""
    return_code: int | None = None
    timed_out = False
    if inputs_ok:
        try:
            process = subprocess.run(command, capture_output=True, text=True, timeout=timeout, check=False)
            return_code = process.returncode
            stdout, stderr = process.stdout, process.stderr
        except subprocess.TimeoutExpired as exc:
            timed_out = True
            stdout = exc.stdout.decode("utf-8", "replace") if isinstance(exc.stdout, bytes) else (exc.stdout or "")
            stderr = exc.stderr.decode("utf-8", "replace") if isinstance(exc.stderr, bytes) else (exc.stderr or "")
    exit_ok = return_code == 0 and not timed_out
    add("overflow.helper_exit", "PASS" if exit_ok else "FAIL", "HARD", "slides_test helper exited successfully" if exit_ok else "slides_test helper failed or timed out", return_code=return_code, timed_out=timed_out, timeout_seconds=timeout)

    combined = "\n".join((stdout, stderr)).strip()
    explicit_pass = PASS_SENTENCE in stdout and "ERROR:" not in combined
    explicit_fail = "ERROR: Slides with content overflowing original canvas" in combined
    if explicit_fail:
        result_status, result_message = "FAIL", "slides_test explicitly reported overflowing content"
    elif exit_ok and explicit_pass and not stderr.strip():
        result_status, result_message = "PASS", "slides_test explicitly reported no overflow"
    else:
        result_status, result_message = "WARN", "slides_test output was not an unambiguous clean PASS; it was not accepted"
    add("overflow.explicit_result", result_status, "HARD", result_message, pass_sentence=PASS_SENTENCE)

    hard_failures = [
        {"id": check["id"], "message": check["message"], "severity": "HARD"}
        for check in checks if check["severity"] == "HARD" and check["status"] == "FAIL"
    ]
    warnings = [
        {"id": check["id"], "message": check["message"], "severity": check["severity"]}
        for check in checks if check["status"] == "WARN"
    ]
    status = "FAIL" if hard_failures else "WARN" if warnings else "PASS"
    python_version = None
    if helper_python.is_file() and os.access(helper_python, os.X_OK):
        try:
            version_process = subprocess.run([str(helper_python), "--version"], capture_output=True, text=True, timeout=10, check=False)
            python_version = (version_process.stdout or version_process.stderr).strip() or None
        except (OSError, subprocess.TimeoutExpired):
            pass
    return {
        "schema_version": "1.0",
        "kind": "sci-diagram-pptx-overflow-report",
        "tool": "sci-diagram-pptx/overflow_check.py",
        "tool_version": TOOL_VERSION,
        "generated_at": utc_now(),
        "status": status,
        "hard_failures": hard_failures,
        "warnings": warnings,
        "checks": checks,
        "pptx": str(pptx),
        "pptx_sha256": sha256_file(pptx) if pptx.is_file() else None,
        "slides_test_script": str(helper),
        "slides_test_script_sha256": sha256_file(helper) if helper.is_file() else None,
        "helper_python": str(helper_python),
        "helper_python_sha256": sha256_file(helper_python) if helper_python.is_file() else None,
        "helper_python_version": python_version,
        "command": command,
        "return_code": return_code,
        "stdout": stdout[-100_000:],
        "stderr": stderr[-100_000:],
    }


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Run Presentations slides_test.py and emit a hash-bound overflow report.")
    parser.add_argument("pptx", type=Path)
    parser.add_argument("--slides-test-script", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--helper-python", type=Path, default=Path(sys.executable), help="Python interpreter used to run slides_test.py")
    parser.add_argument("--timeout", type=int, default=300)
    return parser


def main(argv: Iterable[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if args.timeout <= 0:
        build_parser().error("--timeout must be positive")
    try:
        report = build_report(args.pptx, args.slides_test_script, args.timeout, args.helper_python)
        write_json(report, args.output)
        return 0 if report["status"] == "PASS" else 1
    except OSError as exc:
        payload = {
            "schema_version": "1.0", "kind": "sci-diagram-pptx-overflow-report",
            "tool": "sci-diagram-pptx/overflow_check.py", "tool_version": TOOL_VERSION,
            "generated_at": utc_now(), "status": "FAIL",
            "hard_failures": [{"id": "overflow.operational", "message": str(exc), "severity": "HARD"}],
            "warnings": [], "checks": [{"id": "overflow.operational", "status": "FAIL", "severity": "HARD", "message": str(exc)}],
        }
        try:
            write_json(payload, args.output)
        except OSError:
            print(json.dumps(payload, ensure_ascii=False), file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
