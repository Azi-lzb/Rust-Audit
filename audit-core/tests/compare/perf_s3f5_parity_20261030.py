#!/usr/bin/env python3
"""S3-F5 focused performance profile with value-level parity for every output."""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
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


WORKSPACE = Path(__file__).resolve().parents[3]
AUDIT_ROOT = WORKSPACE / "audit-core"
DEFAULT_PYTHON_SOURCE = Path(r"C:\Users\AZI\Desktop\VibeCode\3.open-data-audit")
DEFAULT_REPORT = AUDIT_ROOT / "probe" / "reports" / "s3f5_parity_profile_20261030.json"


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


def load_perf_driver():
    return load_module(
        "audit_perf_compare_20261030",
        AUDIT_ROOT / "tests" / "compare" / "perf_compare_20261030.py",
    )


def python_render(source_root: Path, comparison: Path, common: Path, forms: Path, output: Path) -> dict[str, Any]:
    sys.path.insert(0, str(source_root / "core" / "src"))
    from base_audit.systems.s3_central_statistics.service import render_financial_forms

    started = time.perf_counter()
    result = render_financial_forms(
        comparison_file=comparison,
        output_dir=output,
        config_path=[common, forms],
    )
    elapsed_ms = (time.perf_counter() - started) * 1000
    files = sorted(output.glob("*.xlsx"), key=lambda path: path.name.casefold())
    return {
        "elapsed_ms": round(elapsed_ms, 2),
        "service_elapsed_ms": round(float(getattr(result, "seconds", elapsed_ms / 1000)) * 1000, 2),
        "files": files,
        "stats": getattr(result, "stats", {}),
    }


def rust_command(binary: Path, comparison: Path, common: Path, forms: Path, output: Path) -> list[str]:
    return [
        str(binary),
        "s3-generate-financial-forms",
        "--comparison-workbook",
        str(comparison),
        "--common-config",
        str(common),
        "--forms-config",
        str(forms),
        "--output-directory",
        str(output),
        "--developer-view",
    ]


def rust_render(binary: Path, comparison: Path, common: Path, forms: Path, output: Path) -> dict[str, Any]:
    started = time.perf_counter()
    completed = subprocess.run(
        rust_command(binary, comparison, common, forms, output),
        check=True,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    wall_ms = (time.perf_counter() - started) * 1000
    report = json.loads(completed.stdout)
    events = [json.loads(line) for line in completed.stderr.splitlines() if line.strip()]
    completed_event = next(
        (event for event in reversed(events) if event.get("status") == "completed"),
        {},
    )
    report["wall_ms_including_cli_startup"] = round(wall_ms, 2)
    report["stage_timings_ms"] = {
        key: completed_event.get(key)
        for key in ("read_time_ms", "planning_time_ms", "editing_time_ms", "elapsed_ms")
    }
    report["per_workbook_edit_ms"] = [
        float(edit["elapsed_ms"])
        for edit in report.get("edits", [])
        if isinstance(edit.get("elapsed_ms"), (int, float))
    ]
    report["progress"] = completed_event
    return report


def values_equal(expected: Any, actual: Any, tolerance: float) -> bool:
    if isinstance(expected, bool) or isinstance(actual, bool):
        return type(expected) is type(actual) and expected == actual
    if isinstance(expected, (int, float)) and isinstance(actual, (int, float)):
        return abs(float(expected) - float(actual)) <= tolerance
    return type(expected) is type(actual) and expected == actual


def compare_workbooks(expected_path: Path, actual_path: Path, tolerance: float = 1e-12) -> dict[str, Any]:
    from openpyxl import load_workbook

    expected_formula = load_workbook(expected_path, read_only=False, data_only=False)
    actual_formula = load_workbook(actual_path, read_only=False, data_only=False)
    expected_values = load_workbook(expected_path, read_only=False, data_only=True)
    actual_values = load_workbook(actual_path, read_only=False, data_only=True)
    differences: list[dict[str, Any]] = []
    stats = {
        "sheets": 0,
        "cells_checked": 0,
        "formula_cells": 0,
        "numeric_cells": 0,
        "hidden_rows_checked": 0,
        "merged_regions_checked": 0,
    }

    def record(kind: str, location: str, expected: Any, actual: Any) -> None:
        if len(differences) < 50:
            differences.append({"kind": kind, "location": location, "expected": str(expected), "actual": str(actual)})

    try:
        if expected_formula.sheetnames != actual_formula.sheetnames:
            record("sheet_names", "workbook", expected_formula.sheetnames, actual_formula.sheetnames)
        stats["sheets"] = min(len(expected_formula.worksheets), len(actual_formula.worksheets))
        for expected_sheet, actual_sheet, expected_cache_sheet, actual_cache_sheet in zip(
            expected_formula.worksheets,
            actual_formula.worksheets,
            expected_values.worksheets,
            actual_values.worksheets,
        ):
            name = expected_sheet.title
            if expected_sheet.sheet_state != actual_sheet.sheet_state:
                record("sheet_state", name, expected_sheet.sheet_state, actual_sheet.sheet_state)
            expected_merges = sorted(str(item) for item in expected_sheet.merged_cells.ranges)
            actual_merges = sorted(str(item) for item in actual_sheet.merged_cells.ranges)
            stats["merged_regions_checked"] += len(expected_merges)
            if expected_merges != actual_merges:
                record("merged_ranges", name, expected_merges, actual_merges)
            if (expected_sheet.max_row, expected_sheet.max_column) != (actual_sheet.max_row, actual_sheet.max_column):
                record(
                    "dimensions",
                    name,
                    (expected_sheet.max_row, expected_sheet.max_column),
                    (actual_sheet.max_row, actual_sheet.max_column),
                )
            max_row = max(expected_sheet.max_row, actual_sheet.max_row)
            max_column = max(expected_sheet.max_column, actual_sheet.max_column)
            for row in range(1, max_row + 1):
                expected_hidden = bool(expected_sheet.row_dimensions[row].hidden)
                actual_hidden = bool(actual_sheet.row_dimensions[row].hidden)
                stats["hidden_rows_checked"] += 1
                if expected_hidden != actual_hidden:
                    record("row_hidden", f"{name}!{row}", expected_hidden, actual_hidden)
                for column in range(1, max_column + 1):
                    expected = expected_sheet.cell(row, column)
                    actual = actual_sheet.cell(row, column)
                    expected_cache = expected_cache_sheet.cell(row, column)
                    actual_cache = actual_cache_sheet.cell(row, column)
                    address = f"{name}!{expected.coordinate}"
                    stats["cells_checked"] += 1
                    if expected.data_type == "f" or actual.data_type == "f":
                        stats["formula_cells"] += 1
                        if expected.data_type != actual.data_type or expected.value != actual.value:
                            record("formula", address, expected.value, actual.value)
                    elif isinstance(expected.value, (int, float)) and not isinstance(expected.value, bool):
                        stats["numeric_cells"] += 1
                        if not values_equal(expected.value, actual.value, tolerance):
                            record("value", address, expected.value, actual.value)
                    elif not values_equal(expected.value, actual.value, tolerance):
                        record("value", address, expected.value, actual.value)
                    if expected.data_type == "f" or actual.data_type == "f":
                        if not values_equal(expected_cache.value, actual_cache.value, tolerance):
                            record("formula_cache", address, expected_cache.value, actual_cache.value)
    finally:
        expected_formula.close()
        actual_formula.close()
        expected_values.close()
        actual_values.close()

    return {"stats": stats, "differences": differences, "difference_count": len(differences)}


def input_info(path: Path) -> dict[str, Any]:
    from openpyxl import load_workbook

    workbook = load_workbook(path, read_only=True, data_only=False)
    try:
        sheets = [
            {"name": sheet.title, "rows": sheet.max_row, "columns": sheet.max_column}
            for sheet in workbook.worksheets
        ]
        return {"bytes": path.stat().st_size, "sheets": sheets}
    finally:
        workbook.close()


def run_monitored_tree(command: list[str], cwd: Path | None = None) -> dict[str, Any]:
    try:
        import psutil
    except ImportError as error:  # pragma: no cover - workstation dependency
        raise RuntimeError("psutil is required for process-tree peak working-set measurement") from error

    with tempfile.TemporaryFile() as stdout_file, tempfile.TemporaryFile() as stderr_file:
        process = subprocess.Popen(command, cwd=cwd, stdout=stdout_file, stderr=stderr_file)
        try:
            root_process = psutil.Process(process.pid)
        except psutil.Error:
            root_process = None
        peak_tree_bytes = 0
        peak_process_count = 0
        while process.poll() is None:
            processes = []
            if root_process is not None:
                try:
                    processes = [root_process, *root_process.children(recursive=True)]
                except psutil.Error:
                    processes = [root_process]
            current_bytes = 0
            alive_count = 0
            for item in processes:
                try:
                    current_bytes += int(getattr(item.memory_info(), "wset", item.memory_info().rss))
                    alive_count += 1
                except psutil.Error:
                    pass
            peak_tree_bytes = max(peak_tree_bytes, current_bytes)
            peak_process_count = max(peak_process_count, alive_count)
            time.sleep(0.005)
        process.wait()
        stdout_file.seek(0)
        stderr_file.seek(0)
        stdout = stdout_file.read().decode("utf-8", errors="replace")
        stderr = stderr_file.read().decode("utf-8", errors="replace")
    return {
        "return_code": process.returncode,
        "stdout": stdout,
        "stderr": stderr,
        "sampled_peak_process_tree_working_set_bytes": peak_tree_bytes,
        "peak_live_process_count": peak_process_count,
    }


def python_memory_worker(source_root: Path, comparison: Path, common: Path, forms: Path, output: Path) -> int:
    run = python_render(source_root, comparison, common, forms, output)
    print(json.dumps({"status": "completed", "outputs": len(run["files"])}, ensure_ascii=True))
    return 0


def median_stage(reports: list[dict[str, Any]], field: str) -> dict[str, float]:
    names = sorted({name for report in reports for name in report.get(field, {}) if isinstance(report.get(field, {}).get(name), (int, float))})
    return {name: round(statistics.median(float(report.get(field, {}).get(name, 0)) for report in reports), 2) for name in names}


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--audit-binary", type=Path)
    parser.add_argument("--source-root", type=Path, default=DEFAULT_PYTHON_SOURCE)
    parser.add_argument("--report", type=Path, default=DEFAULT_REPORT)
    parser.add_argument("--institutions", type=int, default=30)
    parser.add_argument("--indicators", type=int, default=40)
    parser.add_argument("--forms", type=int, default=4)
    parser.add_argument("--rounds", type=int, default=3)
    parser.add_argument("--memory-worker", choices=("python",))
    parser.add_argument("--comparison", type=Path)
    parser.add_argument("--common", type=Path)
    parser.add_argument("--forms-config", type=Path)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()

    if args.memory_worker:
        if not all((args.comparison, args.common, args.forms_config, args.output)):
            parser.error("memory worker requires the comparison/config/output paths")
        return python_memory_worker(
            args.source_root.resolve(),
            args.comparison.resolve(),
            args.common.resolve(),
            args.forms_config.resolve(),
            args.output.resolve(),
        )

    if args.audit_binary is None or not args.audit_binary.is_file():
        parser.error("--audit-binary must point to an existing Release executable")
    if not args.source_root.is_dir():
        parser.error(f"Python source directory does not exist: {args.source_root}")
    if min(args.institutions, args.indicators, args.forms, args.rounds) < 1:
        parser.error("institution, indicator, form and round counts must all be positive")

    driver = load_perf_driver()
    source_root = args.source_root.resolve()
    input_files: dict[str, Path] = {}
    python_times: list[float] = []
    rust_reports: list[dict[str, Any]] = []
    output_checks: list[dict[str, Any]] = []
    with tempfile.TemporaryDirectory(prefix="audit-s3f5-parity-profile-") as temporary:
        root = Path(temporary)
        fixture = root / "fixture"
        fixture.mkdir()
        common, forms, comparison = driver.build_s3_f5(
            fixture, args.institutions, args.indicators, args.forms
        )
        # The S3-F5 perf fixture builder names tabs with A1 so the fixture's
        # RMB comparison rows match and populate the financial forms.
        input_files = {"comparison": comparison, "common_config": common, "forms_config": forms}
        input_hashes_before = {name: sha256(path) for name, path in input_files.items()}
        file_info = {name: input_info(path) for name, path in input_files.items()}

        warm_python = python_render(source_root, comparison, common, forms, root / "warm-python")
        warm_rust = rust_render(args.audit_binary.resolve(), comparison, common, forms, root / "warm-rust")
        if len(warm_python["files"]) != args.institutions or warm_rust.get("institution_count") != args.institutions:
            raise AssertionError(
                f"warm-up output count mismatch: Python={len(warm_python['files'])}, "
                f"Rust={warm_rust.get('institution_count')}"
            )
        expected_matches = args.institutions * args.indicators * args.forms
        if int(warm_rust.get("matched_indicators", 0)) != expected_matches:
            raise AssertionError(
                "S3-F5 synthetic workload did not match every form row; "
                f"matched={warm_rust.get('matched_indicators')}, "
                f"expected={expected_matches}"
            )

        for index in range(args.rounds):
            python_output = root / f"python-{index + 1}"
            rust_output = root / f"rust-{index + 1}"
            if index % 2 == 0:
                python_run = python_render(source_root, comparison, common, forms, python_output)
                rust_run = rust_render(args.audit_binary.resolve(), comparison, common, forms, rust_output)
            else:
                rust_run = rust_render(args.audit_binary.resolve(), comparison, common, forms, rust_output)
                python_run = python_render(source_root, comparison, common, forms, python_output)

            python_times.append(python_run["elapsed_ms"])
            rust_reports.append(rust_run)
            python_names = {path.name for path in python_run["files"]}
            rust_files = [Path(path) for path in rust_run.get("output_files", [])]
            rust_names = {path.name for path in rust_files}
            name_differences = {
                "missing_in_rust": sorted(python_names - rust_names),
                "unexpected_in_rust": sorted(rust_names - python_names),
            }
            if name_differences["missing_in_rust"] or name_differences["unexpected_in_rust"]:
                raise AssertionError(f"round {index + 1} output file names differ: {name_differences}")
            python_by_name = {path.name: path for path in python_run["files"]}
            round_workbook_stats = []
            for rust_path in rust_files:
                checked = compare_workbooks(python_by_name[rust_path.name], rust_path)
                if checked["difference_count"]:
                    raise AssertionError(
                        f"round {index + 1} workbook {rust_path.name} has differences: "
                        f"{checked['differences'][:10]!r}"
                    )
                round_workbook_stats.append({"file": rust_path.name, **checked["stats"]})
            numeric_cells = sum(item["numeric_cells"] for item in round_workbook_stats)
            expected_numeric_cells = expected_matches * 4
            if numeric_cells < expected_numeric_cells:
                raise AssertionError(
                    f"round {index + 1} populated only {numeric_cells} numeric form cells; "
                    f"expected at least {expected_numeric_cells}"
                )
            output_checks.append(
                {
                    "round": index + 1,
                    "python_output_count": len(python_run["files"]),
                    "rust_output_count": len(rust_files),
                    "matched_names": len(python_names & rust_names),
                    "workbooks_value_formula_structure_checked": len(round_workbook_stats),
                    "difference_count": 0,
                    "per_workbook_stats": round_workbook_stats,
                    "rust_report": {
                        key: rust_run.get(key)
                        for key in (
                            "comparison_rows",
                            "institution_count",
                            "matched_indicators",
                            "warnings",
                            "read_time_ms",
                            "planning_time_ms",
                            "editing_time_ms",
                            "elapsed_ms",
                        )
                    },
                    "per_workbook_edit_ms": {
                        "count": len(rust_run.get("per_workbook_edit_ms", [])),
                        "median": round(statistics.median(rust_run["per_workbook_edit_ms"]), 2)
                        if rust_run.get("per_workbook_edit_ms") else None,
                        "minimum": min(rust_run["per_workbook_edit_ms"])
                        if rust_run.get("per_workbook_edit_ms") else None,
                        "maximum": max(rust_run["per_workbook_edit_ms"])
                        if rust_run.get("per_workbook_edit_ms") else None,
                    },
                    "numeric_cells_checked": numeric_cells,
                }
            )

        input_hashes_after = {name: sha256(path) for name, path in input_files.items()}
        if input_hashes_after != input_hashes_before:
            raise AssertionError("S3-F5 modified a comparison or configuration input")

        memory_python_output = root / "memory-python"
        memory_command = [
            sys.executable,
            str(Path(__file__).resolve()),
            "--memory-worker",
            "python",
            "--source-root",
            str(source_root),
            "--comparison",
            str(comparison),
            "--common",
            str(common),
            "--forms-config",
            str(forms),
            "--output",
            str(memory_python_output),
        ]
        memory_python = run_monitored_tree(memory_command, cwd=WORKSPACE)
        if memory_python["return_code"] != 0:
            raise RuntimeError(f"Python process-tree memory run failed: {memory_python['stderr'][-4000:]}")
        memory_rust = run_monitored_tree(
            rust_command(args.audit_binary.resolve(), comparison, common, forms, root / "memory-rust"),
        )
        if memory_rust["return_code"] != 0:
            raise RuntimeError(f"Rust memory run failed: {memory_rust['stderr'][-4000:]}")

        baseline_hashes_after_all = {name: sha256(path) for name, path in input_files.items()}
        if baseline_hashes_after_all != input_hashes_before:
            raise AssertionError("S3-F5 changed an input during its separate memory runs")

    python_median = statistics.median(python_times)
    rust_wall_times = [float(report["wall_ms_including_cli_startup"]) for report in rust_reports]
    rust_internal_times = [float(report["elapsed_ms"]) for report in rust_reports]
    report = {
        "schema_version": 1,
        "date": datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds"),
        "feature": "S3-F5",
        "scenario": "30-institutions-40-indicators-4-forms",
        "platform": platform.platform(),
        "python_version": platform.python_version(),
        "python_source": str(source_root),
        "rust_binary": str(args.audit_binary.resolve()),
        "rust_binary_sha256": sha256(args.audit_binary),
        "timing_scope": {
            "python": "render_financial_forms service wall time; includes its worker-process pool startup and output generation",
            "rust_wall": "CLI subprocess wall time including process startup",
            "rust_internal": "S3-F05 CLI elapsed_ms measured inside command execution",
            "warmup_runs_excluded": 1,
            "round_order": "alternating Python-first and Rust-first",
        },
        "fixture": {
            "institutions": args.institutions,
            "indicators_per_institution": args.indicators,
            "forms": args.forms,
            "comparison_rows": args.institutions * args.indicators,
            "output_workbooks": args.institutions,
            "matched_form_rows": expected_matches,
            "inputs": file_info,
            "input_sha256_before": input_hashes_before,
            "input_sha256_after": input_hashes_after,
        },
        "rounds": {
            "python_service_ms": python_times,
            "python_median_ms": round(python_median, 2),
            "rust_cli_wall_ms": rust_wall_times,
            "rust_cli_wall_median_ms": round(statistics.median(rust_wall_times), 2),
            "rust_internal_ms": rust_internal_times,
            "rust_internal_median_ms": round(statistics.median(rust_internal_times), 2),
            "rust_stage_median_ms": median_stage(rust_reports, "stage_timings_ms"),
            "python_over_rust_cli_wall": round(python_median / statistics.median(rust_wall_times), 3),
            "python_over_rust_internal": round(python_median / statistics.median(rust_internal_times), 3),
        },
        "correctness": {
            "warmup_python_outputs": len(warm_python["files"]),
            "warmup_rust_outputs": warm_rust.get("institution_count"),
            "rounds": output_checks,
            "all_output_names_match": True,
            "all_output_cell_values_formulas_and_key_structure_match": True,
            "input_sha256_unchanged": input_hashes_after == input_hashes_before,
        },
        "process_tree_peak_memory": {
            "measurement": "sampled aggregate working set of parent and live child processes in a separate run; process startup included",
            "interval_ms": 5,
            "python_working_set_bytes": memory_python["sampled_peak_process_tree_working_set_bytes"],
            "python_peak_live_process_count": memory_python["peak_live_process_count"],
            "rust_cli_working_set_bytes": memory_rust["sampled_peak_process_tree_working_set_bytes"],
            "rust_peak_live_process_count": memory_rust["peak_live_process_count"],
        },
        "status": "passed",
    }
    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    compact = {
        "status": report["status"],
        "report": str(args.report),
        "rounds": report["rounds"],
        "correctness": {
            "round_count": len(output_checks),
            "all_output_files_checked": all(item["workbooks_value_formula_structure_checked"] == args.institutions for item in output_checks),
            "all_differences_zero": all(item["difference_count"] == 0 for item in output_checks),
        },
        "memory": report["process_tree_peak_memory"],
    }
    print(json.dumps(compact, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
