#!/usr/bin/env python3
"""Repeated, parity-checked S1 conditional audit profile on a generated workbook."""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import os
import statistics
import subprocess
import sys
import tempfile
import time
from datetime import datetime, timezone
from pathlib import Path
from decimal import Decimal
from typing import Any

import psutil
from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule, FormulaRule


WORKSPACE = Path(__file__).resolve().parents[3]
AUDIT_ROOT = WORKSPACE / "audit-core"
PYTHON_SOURCE = Path(r"C:\Users\AZI\Desktop\VibeCode\3.open-data-audit")
DEFAULT_BINARY = Path(r"F:\RustBuild\audit\release\open-data-audit.exe")
DEFAULT_REPORT = AUDIT_ROOT / "probe" / "reports" / "s1_conditional_profile_20261030.json"
PERIOD = "2026-09"
BATCH_ID = "s1-conditional-profile"
AUDIT_TIME = "2026-10-30T12:00:00+08:00"


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def load_module(name: str, path: Path):
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load module: {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


def generate_fixture(path: Path, rows: int) -> dict[str, Any]:
    book = Workbook()
    sheet = book.active
    sheet.title = "S1Conditional"
    sheet.append(["id", "positive", "signed", "formula_target", "left", "right", "mapped", "text"])
    for index in range(1, rows + 1):
        sheet.append([
            index,
            index % 1000,
            (index % 11) - 5,
            index % 500,
            index * 2,
            index + 1,
            index % 400,
            f"group-{index % 50}",
        ])
    final_row = rows + 1
    sheet.conditional_formatting.add(
        f"B2:B{final_row}", CellIsRule(operator="greaterThan", formula=["500"])
    )
    sheet.conditional_formatting.add(
        f"C2:C{final_row}", CellIsRule(operator="lessThan", formula=["0"])
    )
    sheet.conditional_formatting.add(
        f"D2:D{final_row}", FormulaRule(formula=["E2>F2"])
    )
    sheet.conditional_formatting.add(
        f"G2:G{final_row}", FormulaRule(formula=["AND($B2>100,$B2<300)"])
    )
    book.save(path)
    book.close()
    return {
        "sheets": 1,
        "rows_including_header": final_row,
        "columns": 8,
        "nonempty_cells": (rows + 1) * 8,
        "conditional_rules": 4,
        "scanned_cells_expected": rows * 4,
    }


def peak_child_rss(process: subprocess.Popen[bytes]) -> float:
    try:
        child = psutil.Process(process.pid)
    except psutil.Error:
        return 0.0
    peak = 0
    while process.poll() is None:
        try:
            peak = max(peak, child.memory_info().rss)
        except psutil.Error:
            break
        time.sleep(0.005)
    try:
        peak = max(peak, child.memory_info().rss)
    except psutil.Error:
        pass
    return peak / (1024 * 1024)


def run_python_worker(workbook: Path) -> dict[str, Any]:
    helper = load_module(
        "s1_conditional_profile_worker_helper",
        AUDIT_ROOT / "tests" / "compare" / "compare_s1_conditional_audit.py",
    )
    sys.dont_write_bytecode = True
    sys.path.insert(0, str(PYTHON_SOURCE / "core" / "src"))
    from base_audit.conditional_evaluators import ConditionalFormatService

    service = ConditionalFormatService(mode="PYTHON")
    issues, stats = service.extract_issues(
        workbook_path=workbook,
        ranges=None,
        structure_ranges=[],
        period=PERIOD,
        batch_id=BATCH_ID,
        audit_time=AUDIT_TIME,
        source_file=workbook,
        mapping=None,
    )
    return {
        "issues": [helper.normalize_issue(issue) for issue in issues],
        "internal_elapsed_ms": round(float(stats["total_time"]) * 1000, 3),
        "external_elapsed_ms": None,
        "peak_rss_mb": None,
        "stage_ms": {
            "rule_read": round(float(stats.get("rule_read_time", 0)) * 1000, 3),
            "rule_normalize": round(float(stats.get("rule_normalize_time", 0)) * 1000, 3),
            "evaluation": round(float(stats.get("evaluation_time", 0)) * 1000, 3),
            "issue_build": round(float(stats.get("result_build_time", 0)) * 1000, 3),
            "context_build": round(float(stats.get("context_build_time", 0)) * 1000, 3),
        },
        "stats": stats,
    }


def json_safe(value: Any) -> Any:
    if isinstance(value, Decimal):
        return int(value) if value == value.to_integral_value() else float(value)
    if isinstance(value, dict):
        return {str(key): json_safe(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [json_safe(item) for item in value]
    return value


def run_python_process(workbook: Path, output_dir: Path, index: int, normalize_issue) -> dict[str, Any]:
    stdout_path = output_dir / f"python-{index}.stdout.json"
    stderr_path = output_dir / f"python-{index}.stderr.log"
    command = [sys.executable, str(Path(__file__).resolve()), "--python-worker", str(workbook)]
    environment = os.environ.copy()
    environment["PYTHONIOENCODING"] = "utf-8"
    started = time.perf_counter()
    with stdout_path.open("wb") as stdout, stderr_path.open("wb") as stderr:
        process = subprocess.Popen(command, stdout=stdout, stderr=stderr, env=environment)
        peak_rss_mb = peak_child_rss(process)
        return_code = process.wait()
    external_elapsed_ms = (time.perf_counter() - started) * 1000
    if return_code != 0:
        detail = stderr_path.read_text(encoding="utf-8", errors="replace")
        raise RuntimeError(f"Python conditional service failed ({return_code}): {detail[-4000:]}")
    result = json.loads(stdout_path.read_text(encoding="utf-8"))
    result["issues"] = [normalize_issue(issue) for issue in result.get("issues", [])]
    result["external_elapsed_ms"] = round(external_elapsed_ms, 3)
    result["peak_rss_mb"] = round(peak_rss_mb, 3)
    return result


def run_rust(binary: Path, workbook: Path, output_dir: Path, normalize_issue) -> dict[str, Any]:
    stdout_path = output_dir / "rust.stdout.json"
    stderr_path = output_dir / "rust.stderr.log"
    command = [
        str(binary),
        "conditional-audit",
        "--workbook", str(workbook),
        "--source-file", str(workbook),
        "--period", PERIOD,
        "--batch-id", BATCH_ID,
        "--audit-time", AUDIT_TIME,
        "--developer-view",
    ]
    started = time.perf_counter()
    with stdout_path.open("wb") as stdout, stderr_path.open("wb") as stderr:
        process = subprocess.Popen(command, stdout=stdout, stderr=stderr)
        peak_rss_mb = peak_child_rss(process)
        return_code = process.wait()
    external_elapsed_ms = (time.perf_counter() - started) * 1000
    if return_code != 0:
        detail = stderr_path.read_text(encoding="utf-8", errors="replace")
        raise RuntimeError(f"Rust conditional-audit failed ({return_code}): {detail[-4000:]}")
    report = json.loads(stdout_path.read_text(encoding="utf-8"))
    events = [
        json.loads(line)
        for line in stderr_path.read_text(encoding="utf-8").splitlines()
        if line.strip()
    ]
    completed = next(
        (
            event for event in reversed(events)
            if event.get("step") == "conditional_format_issue_extraction"
            and event.get("status") == "completed"
        ),
        None,
    )
    if completed is None:
        raise RuntimeError("Rust output has no completed conditional-audit timing event")
    return {
        "issues": [normalize_issue(issue) for issue in report.get("issues", [])],
        "internal_elapsed_ms": completed["elapsed_ms"],
        "external_elapsed_ms": round(external_elapsed_ms, 3),
        "peak_rss_mb": round(peak_rss_mb, 3),
        "stage_ms": {
            "read": completed.get("read_time_ms"),
            "evaluation": completed.get("evaluation_time_ms"),
            "issue_build": completed.get("issue_build_time_ms"),
            "reported_total": completed.get("total_time_ms"),
        },
        "rules_total": completed.get("rules_total"),
        "rules_supported": completed.get("rules_supported"),
        "scanned_cells": completed.get("scanned_cells"),
        "issue_count": completed.get("issues"),
        "diagnostics": completed.get("diagnostics"),
    }


def parity(python_rows: list[dict[str, Any]], rust_rows: list[dict[str, Any]]) -> dict[str, Any]:
    py_by_id = {row["issue_id"]: row for row in python_rows}
    rust_by_id = {row["issue_id"]: row for row in rust_rows}
    differences: dict[str, int] = {}
    for issue_id in py_by_id.keys() & rust_by_id.keys():
        for field in py_by_id[issue_id].keys() | rust_by_id[issue_id].keys():
            if py_by_id[issue_id].get(field) != rust_by_id[issue_id].get(field):
                differences[field] = differences.get(field, 0) + 1
    return {
        "python_issues": len(py_by_id),
        "rust_issues": len(rust_by_id),
        "missing_in_rust": len(py_by_id.keys() - rust_by_id.keys()),
        "missing_in_python": len(rust_by_id.keys() - py_by_id.keys()),
        "field_differences": differences,
    }


def summary(values: list[float]) -> dict[str, float]:
    return {
        "median": round(statistics.median(values), 3),
        "min": round(min(values), 3),
        "max": round(max(values), 3),
    }


def main() -> int:
    if len(sys.argv) >= 3 and sys.argv[1] == "--python-worker":
        result = run_python_worker(Path(sys.argv[2]))
        print(json.dumps(json_safe(result), ensure_ascii=False))
        return 0

    parser = argparse.ArgumentParser()
    parser.add_argument("--binary", type=Path, default=DEFAULT_BINARY)
    parser.add_argument("--rows", type=int, default=4000)
    parser.add_argument("--rounds", type=int, default=5)
    parser.add_argument("--warmup", type=int, default=1)
    parser.add_argument("--report", type=Path, default=DEFAULT_REPORT)
    args = parser.parse_args()
    if args.rows < 1 or args.rounds < 1 or args.warmup < 0:
        parser.error("rows/rounds must be positive and warmup must be nonnegative")

    source_core = PYTHON_SOURCE / "core"
    if not source_core.is_dir():
        raise FileNotFoundError(f"Python reference project not found: {source_core}")
    if not args.binary.is_file():
        raise FileNotFoundError(f"Rust executable not found: {args.binary}")
    helper = load_module(
        "s1_conditional_profile_compare_helper",
        AUDIT_ROOT / "tests" / "compare" / "compare_s1_conditional_audit.py",
    )
    records: list[dict[str, Any]] = []
    parity_records: list[dict[str, Any]] = []
    python_peaks: list[float] = []
    python_results: list[dict[str, Any]] = []
    rust_results: list[dict[str, Any]] = []
    with tempfile.TemporaryDirectory(prefix="audit_s1_cf_profile_") as temp_name:
        temp_dir = Path(temp_name)
        workbook = temp_dir / "S1ConditionalProfile.xlsx"
        fixture = generate_fixture(workbook, args.rows)
        original_hash = sha256(workbook)

        python_run_index = 0

        def measured_python(label: str) -> dict[str, Any]:
            nonlocal python_run_index
            python_run_index += 1
            run_dir = temp_dir / f"python-{label}-{python_run_index}"
            run_dir.mkdir(exist_ok=True)
            result = run_python_process(
                workbook, run_dir, python_run_index, helper.normalize_issue
            )
            if label == "measured":
                python_results.append(result)
                python_peaks.append(result["peak_rss_mb"])
            return result

        def measured_rust(label: str, index: int) -> dict[str, Any]:
            run_dir = temp_dir / f"rust-{label}-{index}"
            run_dir.mkdir(exist_ok=True)
            result = run_rust(args.binary.resolve(), workbook, run_dir, helper.normalize_issue)
            if label == "measured":
                rust_results.append(result)
            return result

        for index in range(args.warmup):
            py_result = measured_python("warmup")
            rust_result = measured_rust("warmup", index)
            check = parity(py_result["issues"], rust_result["issues"])
            if check["missing_in_rust"] or check["missing_in_python"] or check["field_differences"]:
                raise AssertionError(f"warmup output mismatch: {check}")

        for index in range(args.rounds):
            if index % 2 == 0:
                rust_result = measured_rust("measured", index)
                py_result = measured_python("measured")
            else:
                py_result = measured_python("measured")
                rust_result = measured_rust("measured", index)
            check = parity(py_result["issues"], rust_result["issues"])
            if check["missing_in_rust"] or check["missing_in_python"] or check["field_differences"]:
                raise AssertionError(f"round {index + 1} output mismatch: {check}")
            parity_records.append(check)
            records.append({
                "round": index + 1,
                "order": "rust_then_python" if index % 2 == 0 else "python_then_rust",
                "python": {key: py_result[key] for key in (
                    "internal_elapsed_ms", "external_elapsed_ms", "peak_rss_mb", "stage_ms"
                )},
                "rust": {key: rust_result[key] for key in (
                    "internal_elapsed_ms", "external_elapsed_ms", "peak_rss_mb", "stage_ms",
                    "rules_total", "rules_supported", "scanned_cells", "issue_count", "diagnostics"
                )},
                "parity": check,
            })

        final_hash = sha256(workbook)
        if final_hash != original_hash:
            raise AssertionError("source workbook hash changed during profiling")
        input_bytes = workbook.stat().st_size

    python_internal = [item["internal_elapsed_ms"] for item in python_results]
    rust_internal = [item["internal_elapsed_ms"] for item in rust_results]
    python_external = [item["external_elapsed_ms"] for item in python_results]
    rust_external = [item["external_elapsed_ms"] for item in rust_results]
    report = {
        "schema_version": 1,
        "suite": "audit.s1_conditional_issue_extraction_profile",
        "status": "passed",
        "created_at": datetime.now(timezone.utc).isoformat(),
        "environment": {
            "platform": sys.platform,
            "python_version": sys.version.split()[0],
            "python_reference_root": str(PYTHON_SOURCE),
            "rust_binary": str(args.binary.resolve()),
            "rust_binary_sha256": sha256(args.binary.resolve()),
            "rust_cli_elapsed_includes_process_startup": True,
            "internal_times_exclude_rust_json_output_serialization": True,
            "peak_rss_sampling_interval_ms": 5,
            "peak_rss_note": "Sampled child-process working set; measurements run in separate Python/Rust processes and do not execute the sampled business code on the sampling thread.",
        },
        "input": {
            **fixture,
            "bytes": input_bytes,
            "sha256_before": original_hash,
            "sha256_after": final_hash,
        },
        "runs": {"warmup": args.warmup, "measured": args.rounds},
        "summary_ms": {
            "python_internal_service": summary(python_internal),
            "rust_internal_audit": summary(rust_internal),
            "python_external_call": summary(python_external),
            "rust_external_cli": summary(rust_external),
            "python_to_rust_internal_ratio": round(
                statistics.median(python_internal) / statistics.median(rust_internal), 4
            ),
        },
        "peak_rss_mb": {
            "python_service_sampled_peak_median": round(statistics.median(python_peaks), 3),
            "rust_cli_sampled_peak_median": round(
                statistics.median(item["peak_rss_mb"] for item in rust_results), 3
            ),
        },
        "correctness": {
            "all_measured_rounds_exact_issue_parity": all(
                item["missing_in_rust"] == 0
                and item["missing_in_python"] == 0
                and not item["field_differences"]
                for item in parity_records
            ),
            "rounds": parity_records,
            "input_unchanged": final_hash == original_hash,
        },
        "measured_rounds": records,
    }
    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
