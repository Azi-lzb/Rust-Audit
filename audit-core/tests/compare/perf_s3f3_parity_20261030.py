#!/usr/bin/env python3
"""Focused S3-F3 feedback-import profile with full workbook parity checks."""

from __future__ import annotations

import argparse
import hashlib
import importlib
import json
import platform
import statistics
import subprocess
import sys
import tempfile
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook


WORKSPACE = Path(__file__).resolve().parents[3]
AUDIT_ROOT = WORKSPACE / "audit-core"
DEFAULT_PYTHON_SOURCE = Path(r"C:\Users\AZI\Desktop\VibeCode\3.open-data-audit")
DEFAULT_BINARY = Path(r"F:\RustBuild\audit\release\open-data-audit.exe")
DEFAULT_REPORT = AUDIT_ROOT / "probe" / "reports" / "s3f3_parity_optimized_20261030.json"
DEFAULT_BASELINE_REPORT = AUDIT_ROOT / "probe" / "reports" / "s3f3_parity_profile_20261030.json"
HEADERS = [
    "机构代码", "机构名称", "指标代码", "指标名称", "本期金额（亿元）",
    "上期值", "增减额（亿元）", "变动原因",
]


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def build_inputs(root: Path, rows: int) -> tuple[Path, Path]:
    """Use the same unique-key, no-conflict fixture shape as the prior baseline."""
    target_book = Workbook()
    target_sheet = target_book.active
    target_sheet.title = "说明文件"
    target_sheet.append(["xx说明"])
    target_sheet.append(HEADERS)

    feedback_book = Workbook()
    feedback_sheet = feedback_book.active
    feedback_sheet.title = "机构反馈"
    feedback_sheet.append(HEADERS)
    for index in range(rows):
        key = (f"ORG{index:05d}", f"机构{index:05d}", f"I{index:05d}", f"指标{index:05d}")
        amount = 100.0 + index
        change = index % 23 - 11
        target_sheet.append([*key, amount, 90.0 + index, change, "待机构填写"])
        feedback_sheet.append([*key, amount, 90.0 + index, change, f"已核实原因{index:05d}"])

    target_path = root / "target_base.xlsx"
    feedback_path = root / "feedback.xlsx"
    target_book.save(target_path)
    feedback_book.save(feedback_path)
    target_book.close()
    feedback_book.close()
    return target_path, feedback_path


def load_python_api(source_root: Path):
    package_root = source_root / "core" / "src"
    sys.path.insert(0, str(package_root))
    from base_audit.systems.s3_central_statistics.explanation_exporter import (
        import_explanation_feedback,
    )

    return import_explanation_feedback


def workbook_snapshot(path: Path, *, data_only: bool = False) -> dict[str, Any]:
    workbook = load_workbook(path, data_only=data_only, read_only=False)
    try:
        sheets: list[dict[str, Any]] = []
        for sheet in workbook.worksheets:
            cells: dict[str, Any] = {}
            formula_count = 0
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.value is not None:
                        cells[cell.coordinate] = cell.value
                    if isinstance(cell.value, str) and cell.value.startswith("="):
                        formula_count += 1
            sheets.append(
                {
                    "name": sheet.title,
                    "state": sheet.sheet_state,
                    "max_row": sheet.max_row,
                    "max_column": sheet.max_column,
                    "cells": cells,
                    "formula_count": formula_count,
                    "merged_ranges": sorted(str(item) for item in sheet.merged_cells.ranges),
                    "freeze_panes": str(sheet.freeze_panes) if sheet.freeze_panes else None,
                    "auto_filter": sheet.auto_filter.ref,
                    "hidden_rows": sorted(
                        index for index, dimension in sheet.row_dimensions.items() if dimension.hidden
                    ),
                    "hidden_columns": sorted(
                        key for key, dimension in sheet.column_dimensions.items() if dimension.hidden
                    ),
                }
            )
        return {"sheet_names": workbook.sheetnames, "sheets": sheets}
    finally:
        workbook.close()


def compare_workbooks(left: Path, right: Path) -> dict[str, Any]:
    left_snapshot = workbook_snapshot(left)
    right_snapshot = workbook_snapshot(right)
    differences: list[str] = []
    if left_snapshot["sheet_names"] != right_snapshot["sheet_names"]:
        differences.append("sheet_names")
    compared_cells = 0
    for sheet_index, (left_sheet, right_sheet) in enumerate(
        zip(left_snapshot["sheets"], right_snapshot["sheets"])
    ):
        for key in (
            "name", "state", "max_row", "max_column", "formula_count", "merged_ranges",
            "freeze_panes", "auto_filter", "hidden_rows", "hidden_columns",
        ):
            if left_sheet[key] != right_sheet[key]:
                differences.append(f"sheet[{sheet_index}].{key}")
        left_cells = left_sheet["cells"]
        right_cells = right_sheet["cells"]
        for coordinate in sorted(set(left_cells) | set(right_cells)):
            compared_cells += 1
            if left_cells.get(coordinate) != right_cells.get(coordinate):
                differences.append(f"sheet[{sheet_index}].{coordinate}")

    left_values = workbook_snapshot(left, data_only=True)
    right_values = workbook_snapshot(right, data_only=True)
    cached_value_differences = 0
    for left_sheet, right_sheet in zip(left_values["sheets"], right_values["sheets"]):
        if left_sheet["cells"] != right_sheet["cells"]:
            cached_value_differences += 1
    return {
        "compared_nonempty_cells": compared_cells,
        "difference_count": len(differences),
        "difference_examples": differences[:20],
        "cached_value_sheet_difference_count": cached_value_differences,
        "python_structure": {
            "sheet_names": left_snapshot["sheet_names"],
            "dimensions": [(sheet["max_row"], sheet["max_column"]) for sheet in left_snapshot["sheets"]],
        },
        "rust_structure": {
            "sheet_names": right_snapshot["sheet_names"],
            "dimensions": [(sheet["max_row"], sheet["max_column"]) for sheet in right_snapshot["sheets"]],
        },
    }


def verify_output(target: Path, original_hash: str, rows: int) -> dict[str, Any]:
    backups = sorted(target.parent.glob(f"{target.stem}_导入前备份_*.xlsx"))
    if len(backups) != 1:
        raise AssertionError(f"expected exactly one backup beside {target.name}, got {len(backups)}")
    backup_hash = sha256(backups[0])
    if backup_hash != original_hash:
        raise AssertionError(f"pre-import backup differs from original target for {target.name}")

    snapshot = workbook_snapshot(target)
    if snapshot["sheet_names"] != ["说明文件"]:
        raise AssertionError(f"unexpected sheet names for {target.name}: {snapshot['sheet_names']}")
    sheet = snapshot["sheets"][0]
    if (sheet["max_row"], sheet["max_column"]) != (rows + 2, len(HEADERS)):
        raise AssertionError(f"unexpected result dimensions for {target.name}: {sheet['max_row']}x{sheet['max_column']}")
    cells = sheet["cells"]
    for index in range(rows):
        coordinate = f"H{index + 3}"
        expected = f"已核实原因{index:05d}"
        if cells.get(coordinate) != expected:
            raise AssertionError(f"feedback write-back mismatch at {target.name}!{coordinate}")
    return {"backup_path": str(backups[0]), "backup_sha256_matches_original": True}


def assert_only_feedback_cells_changed(
    before: dict[str, Any], after: dict[str, Any], rows: int
) -> int:
    if before["sheet_names"] != after["sheet_names"] or len(before["sheets"]) != len(after["sheets"]):
        raise AssertionError("S3-F3 changed the target sheet structure")
    expected = {f"H{index + 3}" for index in range(rows)}
    changes: set[str] = set()
    for old_sheet, new_sheet in zip(before["sheets"], after["sheets"]):
        if (old_sheet["max_row"], old_sheet["max_column"]) != (
            new_sheet["max_row"], new_sheet["max_column"]
        ):
            raise AssertionError("S3-F3 changed the target dimensions")
        for coordinate in set(old_sheet["cells"]) | set(new_sheet["cells"]):
            if old_sheet["cells"].get(coordinate) != new_sheet["cells"].get(coordinate):
                changes.add(coordinate)
    unexpected = changes - expected
    missing = expected - changes
    if unexpected or missing:
        raise AssertionError(
            f"unexpected S3-F3 mutation set: unexpected={sorted(unexpected)[:10]}, missing={sorted(missing)[:10]}"
        )
    return len(changes)


def rust_call(binary: Path, target: Path, feedback: Path) -> tuple[float, dict[str, Any]]:
    command = [
        str(binary), "s3-import-explanation-feedback",
        "--target-workbook", str(target), "--feedback-workbook", str(feedback), "--developer-view",
    ]
    started = time.perf_counter_ns()
    completed = subprocess.run(command, capture_output=True, text=True, encoding="utf-8", timeout=120)
    wall_ms = (time.perf_counter_ns() - started) / 1_000_000
    if completed.returncode != 0:
        raise RuntimeError(f"Rust S3-F3 failed ({completed.returncode}): {completed.stderr[-4000:]}")
    try:
        report = json.loads(completed.stdout)
    except json.JSONDecodeError as error:
        raise RuntimeError(f"Rust S3-F3 did not return JSON: {completed.stdout[-1000:]}") from error
    return wall_ms, report


def monitored_run(command: list[str], *, cwd: Path | None = None) -> dict[str, Any]:
    try:
        import psutil
    except ImportError as error:  # pragma: no cover - workstation dependency
        raise RuntimeError("psutil is required for process-tree peak working-set measurement") from error

    with tempfile.TemporaryFile() as stdout_file, tempfile.TemporaryFile() as stderr_file:
        process = subprocess.Popen(command, cwd=cwd, stdout=stdout_file, stderr=stderr_file)
        try:
            root = psutil.Process(process.pid)
        except psutil.Error:
            root = None
        peak_bytes = 0
        peak_process_count = 0
        while process.poll() is None:
            try:
                processes = [root, *root.children(recursive=True)] if root is not None else []
            except psutil.Error:
                processes = [root] if root is not None else []
            current_bytes = 0
            live_count = 0
            for item in processes:
                if item is None:
                    continue
                try:
                    info = item.memory_info()
                    current_bytes += int(getattr(info, "wset", info.rss))
                    live_count += 1
                except psutil.Error:
                    pass
            peak_bytes = max(peak_bytes, current_bytes)
            peak_process_count = max(peak_process_count, live_count)
            time.sleep(0.005)
        process.wait()
        stdout_file.seek(0)
        stderr_file.seek(0)
        stdout = stdout_file.read().decode("utf-8", errors="replace")
        stderr = stderr_file.read().decode("utf-8", errors="replace")
    if process.returncode != 0:
        raise RuntimeError(f"monitored command failed ({process.returncode}): {stderr[-4000:]}")
    return {
        "return_code": process.returncode,
        "stdout": stdout,
        "stderr": stderr,
        "sampled_peak_process_tree_working_set_bytes": peak_bytes,
        "peak_live_process_count": peak_process_count,
    }


def python_worker(source_root: Path, target: Path, feedback: Path) -> None:
    function = load_python_api(source_root)
    started = time.perf_counter_ns()
    result = function(target, [feedback])
    elapsed_ms = (time.perf_counter_ns() - started) / 1_000_000
    print(json.dumps({
        "elapsed_ms": elapsed_ms,
        "imported": getattr(result, "imported", None),
        "unmatched": getattr(result, "unmatched", None),
        "conflicts": getattr(result, "conflicts", None),
    }))


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--python-source", type=Path, default=DEFAULT_PYTHON_SOURCE)
    parser.add_argument("--binary", type=Path, default=DEFAULT_BINARY)
    parser.add_argument("--rows", type=int, default=1000)
    parser.add_argument("--rounds", type=int, default=3)
    parser.add_argument("--report", type=Path, default=DEFAULT_REPORT)
    parser.add_argument("--baseline-report", type=Path, default=DEFAULT_BASELINE_REPORT)
    parser.add_argument("--worker", action="store_true")
    parser.add_argument("--target", type=Path)
    parser.add_argument("--feedback", type=Path)
    args = parser.parse_args()

    if args.worker:
        if args.target is None or args.feedback is None:
            parser.error("--worker requires --target and --feedback")
        python_worker(args.python_source, args.target, args.feedback)
        return 0

    if args.rows <= 0 or args.rounds <= 0:
        parser.error("--rows and --rounds must be positive")
    if not args.python_source.is_dir():
        parser.error(f"Python source project does not exist: {args.python_source}")
    if not args.binary.is_file():
        parser.error(f"Rust Release binary does not exist: {args.binary}")

    python_function = load_python_api(args.python_source)
    python_module_path = (
        args.python_source / "core" / "src" / "base_audit" / "systems" /
        "s3_central_statistics" / "explanation_exporter.py"
    )
    python_source_hash_before = sha256(python_module_path)
    binary_hash = sha256(args.binary)
    rounds: list[dict[str, Any]] = []

    with tempfile.TemporaryDirectory(prefix="audit-s3f3-parity-") as temporary:
        root = Path(temporary)
        target_base, feedback = build_inputs(root, args.rows)
        original_hash = sha256(target_base)
        feedback_hash = sha256(feedback)
        base_snapshot = workbook_snapshot(target_base)
        input_info = {
            "target": {
                "bytes": target_base.stat().st_size,
                "sheets": [
                    {"name": sheet["name"], "rows": sheet["max_row"], "columns": sheet["max_column"]}
                    for sheet in base_snapshot["sheets"]
                ],
                "sha256": original_hash,
            },
            "feedback": {
                "bytes": feedback.stat().st_size,
                "sheets": [
                    {"name": sheet["name"], "rows": sheet["max_row"], "columns": sheet["max_column"]}
                    for sheet in workbook_snapshot(feedback)["sheets"]
                ],
                "sha256": feedback_hash,
            },
        }

        plan_probe_dir = root / "plan-probe"
        plan_probe_dir.mkdir()
        plan_target = plan_probe_dir / "target.xlsx"
        plan_feedback = plan_probe_dir / "feedback.xlsx"
        plan_candidate = plan_probe_dir / "candidate.xlsx"
        plan_path = plan_probe_dir / "plan.json"
        plan_target.write_bytes(target_base.read_bytes())
        plan_feedback.write_bytes(feedback.read_bytes())
        plan_command = [
            str(args.binary), "s3-plan-explanation-feedback-import",
            "--target-workbook", str(plan_target),
            "--feedback-workbook", str(plan_feedback),
            "--candidate-output-workbook", str(plan_candidate),
            "--plan-output", str(plan_path), "--developer-view",
        ]
        plan_started = time.perf_counter_ns()
        plan_completed = subprocess.run(
            plan_command, capture_output=True, text=True, encoding="utf-8", timeout=120
        )
        plan_wall_ms = (time.perf_counter_ns() - plan_started) / 1_000_000
        if plan_completed.returncode != 0:
            raise RuntimeError(f"Rust S3-F3 plan probe failed: {plan_completed.stderr[-4000:]}")
        plan_report = json.loads(plan_completed.stdout)
        if (plan_report.get("imported"), plan_report.get("unmatched"), plan_report.get("conflicts")) != (
            args.rows, 0, 0
        ):
            raise AssertionError(f"unexpected plan-only S3-F3 summary: {plan_report}")
        if sha256(plan_target) != original_hash or sha256(plan_feedback) != feedback_hash:
            raise AssertionError("S3-F3 plan-only operation changed an input")
        plan_probe = {
            "scope": "separate plan-only CLI invocation on identical fixture before timed in-place A/B rounds",
            "cli_wall_ms": plan_wall_ms,
            "internal_total_ms": plan_report.get("total_time_ms"),
            "target_read_ms": plan_report.get("target_read_time_ms"),
            "feedback_read_ms": plan_report.get("feedback_read_time_ms"),
            "conflict_report_write_ms": plan_report.get("conflict_write_time_ms"),
            "plan_file_write_ms": plan_report.get("plan_write_time_ms"),
            "matches_expected": True,
            "inputs_unchanged": True,
        }

        for round_index in range(-1, args.rounds):
            label = "warmup" if round_index == -1 else f"round-{round_index + 1}"
            py_target = root / f"{label}-python" / "target.xlsx"
            rust_target = root / f"{label}-rust" / "target.xlsx"
            py_target.parent.mkdir(parents=True)
            rust_target.parent.mkdir(parents=True)
            py_target.write_bytes(target_base.read_bytes())
            rust_target.write_bytes(target_base.read_bytes())

            def call_python() -> tuple[float, dict[str, Any]]:
                started = time.perf_counter_ns()
                report = python_function(py_target, [feedback])
                elapsed = (time.perf_counter_ns() - started) / 1_000_000
                payload = {
                    "imported": getattr(report, "imported", None),
                    "unmatched": getattr(report, "unmatched", None),
                    "conflicts": getattr(report, "conflicts", None),
                }
                return elapsed, payload

            if round_index < 0 or round_index % 2 == 0:
                python_ms, python_report = call_python()
                rust_ms, rust_report = rust_call(args.binary, rust_target, feedback)
            else:
                rust_ms, rust_report = rust_call(args.binary, rust_target, feedback)
                python_ms, python_report = call_python()

            if python_report != {
                "imported": args.rows, "unmatched": 0, "conflicts": 0,
            }:
                raise AssertionError(f"unexpected Python S3-F3 summary: {python_report}")
            rust_summary = {
                "imported": rust_report.get("imported"),
                "unmatched": rust_report.get("unmatched"),
                "conflicts": rust_report.get("conflicts"),
            }
            if rust_summary != python_report:
                raise AssertionError(f"Python/Rust S3-F3 summaries differ: {python_report} != {rust_summary}")

            py_backup = verify_output(py_target, original_hash, args.rows)
            rust_backup = verify_output(rust_target, original_hash, args.rows)
            py_mutations = assert_only_feedback_cells_changed(base_snapshot, workbook_snapshot(py_target), args.rows)
            rust_mutations = assert_only_feedback_cells_changed(base_snapshot, workbook_snapshot(rust_target), args.rows)
            comparison = compare_workbooks(py_target, rust_target)
            if comparison["difference_count"] or comparison["cached_value_sheet_difference_count"]:
                raise AssertionError(f"S3-F3 output parity failed: {comparison}")

            if round_index >= 0:
                rounds.append({
                    "round": round_index + 1,
                    "order": "python-first" if round_index % 2 == 0 else "rust-first",
                    "python_service_ms": python_ms,
                    "rust_cli_wall_ms": rust_ms,
                    "rust_internal_total_ms": rust_report.get("total_time_ms"),
                    "rust_stages_ms": {
                        "target_read": rust_report.get("target_read_time_ms"),
                        "feedback_read": rust_report.get("feedback_read_time_ms"),
                        "writeback": rust_report.get("writeback_time_ms"),
                        "total": rust_report.get("total_time_ms"),
                    },
                    "summary": python_report,
                    "backup_checks": {"python": py_backup, "rust": rust_backup},
                    "updated_target_mutations": {"python": py_mutations, "rust": rust_mutations},
                    "output_comparison": comparison,
                })

        if sha256(feedback) != feedback_hash or sha256(target_base) != original_hash:
            raise AssertionError("S3-F3 changed a source fixture")

        memory_python_target = root / "memory-python" / "target.xlsx"
        memory_rust_target = root / "memory-rust" / "target.xlsx"
        memory_python_target.parent.mkdir(parents=True)
        memory_rust_target.parent.mkdir(parents=True)
        memory_python_target.write_bytes(target_base.read_bytes())
        memory_rust_target.write_bytes(target_base.read_bytes())

        python_memory = monitored_run([
            sys.executable, str(Path(__file__).resolve()),
            "--python-source", str(args.python_source), "--worker",
            "--target", str(memory_python_target), "--feedback", str(feedback),
        ])
        rust_memory = monitored_run([
            str(args.binary), "s3-import-explanation-feedback",
            "--target-workbook", str(memory_rust_target),
            "--feedback-workbook", str(feedback), "--developer-view",
        ])
        python_memory_report = json.loads(python_memory["stdout"])
        rust_memory_report = json.loads(rust_memory["stdout"])
        verify_output(memory_python_target, original_hash, args.rows)
        verify_output(memory_rust_target, original_hash, args.rows)
        memory_comparison = compare_workbooks(memory_python_target, memory_rust_target)
        if memory_comparison["difference_count"]:
            raise AssertionError(f"memory probe outputs differ: {memory_comparison}")

    python_source_hash_after = sha256(python_module_path)
    if python_source_hash_after != python_source_hash_before:
        raise AssertionError("read-only Python source changed during S3-F3 profiling")

    python_median = statistics.median(item["python_service_ms"] for item in rounds)
    rust_wall_median = statistics.median(item["rust_cli_wall_ms"] for item in rounds)
    rust_internal_values = [item["rust_internal_total_ms"] for item in rounds]
    rust_internal_median = statistics.median(rust_internal_values) if all(
        value is not None for value in rust_internal_values
    ) else None
    baseline_comparison = None
    if (
        args.baseline_report.is_file()
        and args.baseline_report.resolve() != args.report.resolve()
        and rust_internal_median is not None
    ):
        baseline = json.loads(args.baseline_report.read_text(encoding="utf-8-sig"))
        old_internal = baseline.get("rust_internal_total_median_ms")
        old_cli = baseline.get("rust_cli_wall_median_ms")
        if old_internal is not None and old_cli is not None:
            baseline_comparison = {
                "baseline_report": str(args.baseline_report),
                "baseline_rust_binary_sha256": baseline.get("rust_binary_sha256"),
                "baseline_python_service_median_ms": baseline.get("python_service_median_ms"),
                "baseline_rust_cli_wall_median_ms": old_cli,
                "baseline_rust_internal_total_median_ms": old_internal,
                "rust_cli_wall_ms_saved": old_cli - rust_wall_median,
                "rust_internal_ms_saved": old_internal - rust_internal_median,
                "rust_cli_wall_reduction_percent": round((old_cli - rust_wall_median) / old_cli * 100, 2),
                "rust_internal_reduction_percent": round((old_internal - rust_internal_median) / old_internal * 100, 2),
            }
    rust_orchestration_residual_ms = [
        item["rust_stages_ms"]["total"]
        - sum(item["rust_stages_ms"][key] for key in ("target_read", "feedback_read", "writeback"))
        for item in rounds
    ]
    rust_orchestration_residual_median_ms = statistics.median(rust_orchestration_residual_ms)
    report = {
        "schema_version": 1,
        "date": datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds"),
        "feature": "S3-F3",
        "scenario": f"{args.rows}-unique-feedback-rows-no-conflicts",
        "platform": f"{platform.system()}-{platform.release()}-{platform.machine()}",
        "python_version": sys.version.split()[0],
        "python_source": str(args.python_source),
        "python_module_sha256_before": python_source_hash_before,
        "python_module_sha256_after": python_source_hash_after,
        "rust_binary": str(args.binary),
        "rust_binary_sha256": binary_hash,
        "timing_scope": {
            "python": "in-process import_explanation_feedback service call; excludes interpreter/module startup and fixture/output verification",
            "rust_wall": "Release CLI process wall time; includes process startup and excludes output verification",
            "rust_internal": "S3-F3 total_time_ms reported inside Rust execution",
            "warmup_runs_excluded": 1,
            "round_order": "alternates Python-first and Rust-first",
        },
        "fixture": {
            "feedback_rows": args.rows,
            "target_workbooks_per_round": 2,
            "feedback_workbooks_per_round": 1,
            "matching_keys": args.rows,
            "unmatched_rows": 0,
            "conflicts": 0,
            "inputs": input_info,
        },
        "python_service_ms": [item["python_service_ms"] for item in rounds],
        "rust_cli_wall_ms": [item["rust_cli_wall_ms"] for item in rounds],
        "rust_internal_total_ms": rust_internal_values,
        "python_service_median_ms": python_median,
        "rust_cli_wall_median_ms": rust_wall_median,
        "rust_internal_total_median_ms": rust_internal_median,
        "python_over_rust_cli_speedup": round(python_median / rust_wall_median, 3),
        "python_over_rust_internal_speedup": round(python_median / rust_internal_median, 3)
        if rust_internal_median else None,
        "rounds": rounds,
        "baseline_comparison": baseline_comparison,
        "plan_only_probe": plan_probe,
        "rust_orchestration_residual_ms": rust_orchestration_residual_ms,
        "rust_orchestration_residual_median_ms": rust_orchestration_residual_median_ms,
        "process_tree_peak_memory": {
            "measurement": "sampled aggregate working set of parent and live child processes in a separate run; process startup/import included; 5 ms sampling interval",
            "python_working_set_bytes": python_memory["sampled_peak_process_tree_working_set_bytes"],
            "python_peak_live_process_count": python_memory["peak_live_process_count"],
            "rust_cli_working_set_bytes": rust_memory["sampled_peak_process_tree_working_set_bytes"],
            "rust_peak_live_process_count": rust_memory["peak_live_process_count"],
            "memory_probe_service_ms": python_memory_report["elapsed_ms"],
        },
        "correctness": {
            "all_round_summaries_match": True,
            "all_1000_feedback_rows_applied": True,
            "only_expected_reason_cells_changed": True,
            "backups_are_byte_identical_to_original_targets": True,
            "python_rust_output_value_formula_and_key_structure_differences": 0,
            "feedback_input_unchanged": True,
            "python_source_unchanged": True,
            "fixture_scope": "unique keys with no conflicts; conflict behavior remains covered by existing business tests, not this performance sample",
        },
        "status": "passed",
    }
    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    markdown_path = args.report.with_suffix(".md")
    stage_median = {
        key: statistics.median(
            item["rust_stages_ms"][key] for item in rounds
            if item["rust_stages_ms"][key] is not None
        ) if all(item["rust_stages_ms"][key] is not None for item in rounds) else None
        for key in ("target_read", "feedback_read", "writeback", "total")
    }
    report["rust_stage_medians_ms"] = stage_median
    args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    if baseline_comparison:
        result_judgement = (
            f"本次 Rust 内部总耗时较优化前减少 {baseline_comparison['rust_internal_ms_saved']:.0f} ms "
            f"（{baseline_comparison['rust_internal_reduction_percent']:.2f}%），CLI 墙钟减少 "
            f"{baseline_comparison['rust_cli_wall_ms_saved']:.2f} ms "
            f"（{baseline_comparison['rust_cli_wall_reduction_percent']:.2f}%）。"
            "S3-F3 原地执行现在在内存中把匹配计划交给 Excel-Core，不再序列化临时计划 JSON 后立即读回；"
            "显式的“生成计划”流程仍保留 JSON 文件输出。相同业务输入下输出及备份检查均无差异。"
        )
        implementation_note = "本轮只对原地执行路径移除了临时计划 JSON 的磁盘往返；用户可见的计划导出路径未改变。"
    else:
        result_judgement = (
            "本样本 Python 服务与 Rust 内部耗时接近；Rust CLI 墙钟包含启动，因此整体 CLI 比较略慢。"
            "该无冲突合成负载不覆盖冲突明细生成、真实业务大文件或 GUI 常驻服务，不能替代人工验收。"
        )
        implementation_note = "本轮没有修改 Rust 产品代码。"
    report_md = f"""# S3-F3 机构反馈导入性能与输出对拍（2026-10-30）

## 范围

沿用既有 1,000 条唯一反馈键、无冲突的合成负载。目标工作簿先复制到独立临时目录；Python 与 Rust 各处理 1,000 行，均须先生成与原目标字节一致的备份，再原地更新说明文件。输入共 2 个 XLSX：目标 {input_info['target']['bytes']:,} bytes、1 张表 {args.rows + 2}×8；反馈 {input_info['feedback']['bytes']:,} bytes、1 张表 {args.rows + 1}×8。反馈输入和基准目标在运行前后 SHA-256 未变。

使用 Release 二进制 SHA-256 `{binary_hash}`；预热一次后交错运行 {args.rounds} 轮。Python 计时不含解释器/模块导入，Rust CLI 墙钟含进程启动；两端计时均不含输出对拍。峰值工作集在独立进程运行中以 5 ms 采样，不含测试控制器。

## 耗时

| 指标 | Python | Rust CLI | Rust 内部 |
|---|---:|---:|---:|
| 三轮中位数 | {python_median:.2f} ms | {rust_wall_median:.2f} ms | {rust_internal_median if rust_internal_median is not None else '未提供'} ms |
| Python / Rust | — | {python_median / rust_wall_median:.2f}× | {python_median / rust_internal_median:.2f}× | 

Rust 内部阶段中位数：目标读取 {stage_median['target_read']} ms、反馈读取 {stage_median['feedback_read']} ms、Excel-Core 写回 {stage_median['writeback']} ms、报告总计 {stage_median['total']} ms。扣除这三个阶段后的端到端剩余部分中位数为 {rust_orchestration_residual_median_ms} ms，包含计划/哈希/备份/发布等未分别计时步骤，不能将它进一步归因于单一操作。

另在相同输入上单独运行只读的计划命令：CLI 墙钟 {plan_probe['cli_wall_ms']:.2f} ms、内部总计 {plan_probe['internal_total_ms']} ms；目标读取 {plan_probe['target_read_ms']} ms、反馈读取 {plan_probe['feedback_read_ms']} ms、冲突清单写入 {plan_probe['conflict_report_write_ms']} ms、计划文件写入 {plan_probe['plan_file_write_ms']} ms。此诊断运行不修改目标，也不与端到端结果相加。

| 轮次 | Python 服务 | Rust CLI 墙钟 | Rust 内部总计 | 顺序 | 值/结构差异 |
|---:|---:|---:|---:|---|---:|
"""
    for item in rounds:
        report_md += (
            f"| {item['round']} | {item['python_service_ms']:.2f} ms | {item['rust_cli_wall_ms']:.2f} ms | "
            f"{item['rust_internal_total_ms']} ms | {item['order']} | {item['output_comparison']['difference_count']} |\n"
        )
    python_mb = report["process_tree_peak_memory"]["python_working_set_bytes"] / (1024 * 1024)
    rust_mb = report["process_tree_peak_memory"]["rust_cli_working_set_bytes"] / (1024 * 1024)
    report_md += f"""
## 正确性和资源

每轮均导入 {args.rows} 条、未匹配 0、冲突 0；更新后的目标只改变 H3:H{args.rows + 2} 的说明值，其它业务值和表结构保持不变。Python/Rust 输出逐格比较，包含公式文本与公式计数、工作表名/顺序/尺寸、合并区域、冻结窗格、筛选、隐藏行列及 data-only 读值，均为 0 差异。每份更新目标旁的备份与原目标 SHA-256 完全一致；反馈文件、源目标样本和 Python 源文件未变化。

独立进程树峰值工作集约 Python {python_mb:.2f} MB（{report['process_tree_peak_memory']['python_peak_live_process_count']} 个进程）、Rust {rust_mb:.2f} MB（{report['process_tree_peak_memory']['rust_peak_live_process_count']} 个进程）。采样值会漏掉短于 5 ms 的瞬时峰值，且 Python/Rust 启动边界不同；作为同机负载观察，不代表所有生产环境内存峰值。

## 判断

{result_judgement}

逐轮原始数据见同名 JSON，复跑驱动为 `tests/compare/perf_s3f3_parity_20261030.py`。{implementation_note}
"""
    markdown_path.write_text(report_md, encoding="utf-8")
    print(json.dumps({
        "feature": "S3-F3",
        "status": "passed",
        "python_median_ms": python_median,
        "rust_cli_wall_median_ms": rust_wall_median,
        "rust_internal_median_ms": rust_internal_median,
        "speedup_cli": round(python_median / rust_wall_median, 3),
        "output_differences": sum(item["output_comparison"]["difference_count"] for item in rounds),
        "peak_working_set_mb": {"python": round(python_mb, 2), "rust": round(rust_mb, 2)},
        "report": str(args.report),
        "markdown": str(markdown_path),
    }, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
