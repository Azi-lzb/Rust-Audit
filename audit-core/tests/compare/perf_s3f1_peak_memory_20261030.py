#!/usr/bin/env python3
"""Memory-only S3-F1 process-tree profile; does not compare elapsed time."""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import itertools
import json
import math
import platform
import statistics
import subprocess
import sys
import tempfile
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import openpyxl
import psutil


WORKSPACE = Path(__file__).resolve().parents[3]
AUDIT_ROOT = WORKSPACE / "audit-core"
PERF_DRIVER = AUDIT_ROOT / "tests" / "compare" / "perf_compare_20261030.py"
PYTHON_SOURCE = Path(r"C:\Users\AZI\Desktop\VibeCode\3.open-data-audit")
DEFAULT_BINARY = Path(r"F:\RustBuild\audit\release\open-data-audit.exe")
DEFAULT_REPORT = AUDIT_ROOT / "probe" / "reports" / "s3f1_peak_memory_20261030.json"


def load_module(name: str, path: Path):
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load module: {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def git_revision() -> str:
    completed = subprocess.run(
        ["git", "rev-parse", "--short", "HEAD"],
        cwd=WORKSPACE,
        check=True,
        capture_output=True,
        text=True,
    )
    return completed.stdout.strip()


def measure_process(command: list[str], sample_interval_s: float = 0.005) -> tuple[dict[str, Any], dict[str, Any]]:
    process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    root = psutil.Process(process.pid)
    peak_tree_rss = 0
    peak_process_count = 0
    sample_count = 0
    while True:
        try:
            processes = [root, *root.children(recursive=True)]
        except psutil.Error:
            processes = []
        rss = 0
        live_count = 0
        for child in {item.pid: item for item in processes}.values():
            try:
                rss += int(child.memory_info().rss)
                live_count += 1
            except psutil.Error:
                continue
        peak_tree_rss = max(peak_tree_rss, rss)
        peak_process_count = max(peak_process_count, live_count)
        sample_count += 1
        if process.poll() is not None:
            break
        time.sleep(sample_interval_s)

    stdout, stderr = process.communicate()
    sample = {
        "return_code": process.returncode,
        "peak_sampled_process_tree_working_set_bytes": peak_tree_rss,
        "peak_sampled_process_tree_working_set_mb": round(peak_tree_rss / (1024 * 1024), 3),
        "peak_live_process_count": peak_process_count,
        "sample_interval_ms": round(sample_interval_s * 1000, 3),
        "memory_samples": sample_count,
    }
    if process.returncode != 0:
        raise RuntimeError(
            f"worker failed ({process.returncode}):\n{stderr.decode('utf-8', errors='replace')[-5000:]}"
        )
    try:
        result = json.loads(stdout.decode("utf-8", errors="replace"))
    except json.JSONDecodeError as error:
        raise RuntimeError(
            f"worker returned invalid JSON: {stdout.decode('utf-8', errors='replace')[-2000:]}"
        ) from error
    return sample, result


def compare_workbooks(left_path: Path, right_path: Path) -> dict[str, Any]:
    left = openpyxl.load_workbook(left_path, read_only=True, data_only=False)
    right = openpyxl.load_workbook(right_path, read_only=True, data_only=False)
    try:
        left_names = list(left.sheetnames)
        right_names = list(right.sheetnames)
        if left_names != right_names:
            return {
                "passed": False,
                "sheet_names_equal": False,
                "python_sheet_names": left_names,
                "rust_sheet_names": right_names,
                "cell_differences": -1,
            }

        differences = 0
        max_numeric_delta = 0.0
        rounded_10_decimal_differences = 0
        dimensions_equal = True
        formula_counts = {"python": 0, "rust": 0}
        compared_cells = 0
        for name in left_names:
            left_sheet = left[name]
            right_sheet = right[name]
            dimensions_equal &= (
                left_sheet.max_row == right_sheet.max_row
                and left_sheet.max_column == right_sheet.max_column
            )
            max_row = max(left_sheet.max_row, right_sheet.max_row)
            max_column = max(left_sheet.max_column, right_sheet.max_column)
            left_rows = left_sheet.iter_rows(
                min_row=1,
                max_row=max_row,
                min_col=1,
                max_col=max_column,
                values_only=True,
            )
            right_rows = right_sheet.iter_rows(
                min_row=1,
                max_row=max_row,
                min_col=1,
                max_col=max_column,
                values_only=True,
            )
            for left_values, right_values in itertools.zip_longest(left_rows, right_rows, fillvalue=()):
                for left_value, right_value in zip(left_values, right_values, strict=False):
                    compared_cells += 1
                    if isinstance(left_value, str) and left_value.startswith("="):
                        formula_counts["python"] += 1
                    if isinstance(right_value, str) and right_value.startswith("="):
                        formula_counts["rust"] += 1
                    if isinstance(left_value, (int, float)) and not isinstance(left_value, bool) and isinstance(right_value, (int, float)) and not isinstance(right_value, bool):
                        delta = abs(float(left_value) - float(right_value))
                        max_numeric_delta = max(max_numeric_delta, delta)
                        if round(float(left_value), 10) != round(float(right_value), 10):
                            rounded_10_decimal_differences += 1
                        if not math.isclose(float(left_value), float(right_value), rel_tol=0.0, abs_tol=1e-10):
                            differences += 1
                    elif left_value != right_value:
                        differences += 1
        return {
            "passed": differences == 0 and dimensions_equal and formula_counts["python"] == formula_counts["rust"],
            "sheet_names_equal": True,
            "dimensions_equal": dimensions_equal,
            "python_sheet_names": left_names,
            "cell_differences_outside_1e_10_tolerance": differences,
            "rounded_10_decimal_differences": rounded_10_decimal_differences,
            "max_numeric_delta": max_numeric_delta,
            "formula_counts": formula_counts,
            "compared_cells": compared_cells,
        }
    finally:
        left.close()
        right.close()


def python_worker(current_csv: Path, previous_csv: Path, common_config: Path, feature_config: Path, output: Path, python_source: Path) -> int:
    sys.dont_write_bytecode = True
    sys.path.insert(0, str(python_source / "core" / "src"))
    from base_audit.systems.s3_central_statistics.service import run_comparison

    run_comparison(
        current_csv=current_csv,
        previous_csv=previous_csv,
        output_path=output,
        config_path=[common_config, feature_config],
        write_flow_logs=False,
        expression_schema="FIVE_SEGMENT_V1",
    )
    print(json.dumps({"status": "completed", "output_path": str(output)}))
    return 0


def run_profile(binary: Path, python_source: Path, report_path: Path, rows: int, rounds: int) -> dict[str, Any]:
    if not binary.is_file():
        raise FileNotFoundError(f"Rust Audit binary not found: {binary}")
    if not python_source.is_dir():
        raise FileNotFoundError(f"Python reference project not found: {python_source}")
    perf_driver = load_module("s3f1_peak_memory_fixture", PERF_DRIVER)
    with tempfile.TemporaryDirectory(prefix="audit-s3f1-peak-memory-") as temp_name:
        root = Path(temp_name)
        config_dir = root / "config"
        config_dir.mkdir()
        common_config, feature_config, _ = perf_driver.build_s3_configs(config_dir, with_cross=False)
        current_csv = root / "current.csv"
        previous_csv = root / "previous.csv"
        perf_driver.build_s3_csv(current_csv, rows)
        perf_driver.build_s3_csv(previous_csv, rows)
        input_paths = [current_csv, previous_csv, common_config, feature_config]
        hashes_before = {path.name: sha256(path) for path in input_paths}
        inputs = {
            "current_csv": current_csv,
            "previous_csv": previous_csv,
            "common_config": common_config,
            "feature_config": feature_config,
        }
        rounds_data = []
        script = Path(__file__).resolve()
        for index in range(rounds):
            python_output = root / f"python-{index + 1}.xlsx"
            rust_output = root / f"rust-{index + 1}.xlsx"
            python_command = [
                sys.executable,
                str(script),
                "--worker",
                "python",
                str(inputs["current_csv"]),
                str(inputs["previous_csv"]),
                str(inputs["common_config"]),
                str(inputs["feature_config"]),
                str(python_output),
                str(python_source),
            ]
            rust_command = [
                str(binary),
                "s3-execution-compare",
                "--current-data",
                str(current_csv),
                "--previous-data",
                str(previous_csv),
                "--common-config",
                str(common_config),
                "--feature-config",
                str(feature_config),
                "--output-workbook",
                str(rust_output),
                "--expression-schema",
                "five-segment",
            ]
            if index % 2 == 0:
                python_memory, python_result = measure_process(python_command)
                rust_memory, rust_result = measure_process(rust_command)
                order = "Python-first"
            else:
                rust_memory, rust_result = measure_process(rust_command)
                python_memory, python_result = measure_process(python_command)
                order = "Rust-first"

            if python_result.get("status") != "completed":
                raise AssertionError(f"Python S3-F1 failed: {python_result!r}")
            if rust_result.get("status") != "completed":
                raise AssertionError(f"Rust S3-F1 failed: {rust_result!r}")
            comparison = compare_workbooks(python_output, rust_output)
            if not comparison["passed"]:
                raise AssertionError(f"S3-F1 output mismatch: {comparison!r}")
            rounds_data.append(
                {
                    "round": index + 1,
                    "order": order,
                    "python_process_tree": python_memory,
                    "rust_process_tree": rust_memory,
                    "output_comparison": comparison,
                    "python_output_bytes": python_output.stat().st_size,
                    "rust_output_bytes": rust_output.stat().st_size,
                }
            )

        hashes_after = {path.name: sha256(path) for path in input_paths}
        if hashes_before != hashes_after:
            raise AssertionError("S3-F1 modified input files")
        python_peaks = [row["python_process_tree"]["peak_sampled_process_tree_working_set_bytes"] for row in rounds_data]
        rust_peaks = [row["rust_process_tree"]["peak_sampled_process_tree_working_set_bytes"] for row in rounds_data]
        report = {
            "schema_version": 1,
            "generated_at_utc": datetime.now(timezone.utc).isoformat(),
            "git_revision": git_revision(),
            "feature": "S3-F1",
            "measurement": "process-tree peak working set only; no elapsed-time comparison performed",
            "platform": platform.platform(),
            "python_version": platform.python_version(),
            "python_source": str(python_source),
            "rust_binary": str(binary),
            "rust_binary_sha256": sha256(binary),
            "fixture": {
                "kind": "synthetic 10k two-period CSV from perf_compare_20261030.py",
                "current_rows": rows,
                "previous_rows": rows,
                "csv_columns": 14,
                "common_config_sheets": 4,
                "feature_config_sheets": 5,
                "input_sizes_bytes": {path.name: path.stat().st_size for path in input_paths},
                "input_sha256_before": hashes_before,
                "input_sha256_after": hashes_after,
            },
            "method": {
                "rounds": rounds,
                "round_order": "alternates Python-first and Rust-first",
                "sample_interval_ms": 5,
                "peak_metric": "maximum sampled sum of RSS across worker process and descendants",
                "timing_results_not_recorded": True,
                "measurement_limit": "sampling can miss sub-interval peaks; includes worker/CLI startup but excludes input generation and output comparison",
            },
            "summary": {
                "python_peak_working_set_median_mb": round(statistics.median(python_peaks) / (1024 * 1024), 3),
                "python_peak_working_set_max_mb": round(max(python_peaks) / (1024 * 1024), 3),
                "rust_peak_working_set_median_mb": round(statistics.median(rust_peaks) / (1024 * 1024), 3),
                "rust_peak_working_set_max_mb": round(max(rust_peaks) / (1024 * 1024), 3),
                "python_to_rust_peak_median_ratio": round(statistics.median(python_peaks) / statistics.median(rust_peaks), 3),
                "all_round_output_comparisons_passed": True,
                "all_input_hashes_unchanged": True,
            },
            "rounds": rounds_data,
        }
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return report


def main() -> int:
    if len(sys.argv) > 1 and sys.argv[1] == "--worker":
        if len(sys.argv) != 9 or sys.argv[2] != "python":
            raise SystemExit("usage: --worker python CURRENT PREVIOUS COMMON_CONFIG FEATURE_CONFIG OUTPUT PYTHON_SOURCE")
        return python_worker(
            Path(sys.argv[3]),
            Path(sys.argv[4]),
            Path(sys.argv[5]),
            Path(sys.argv[6]),
            Path(sys.argv[7]),
            Path(sys.argv[8]),
        )

    parser = argparse.ArgumentParser()
    parser.add_argument("--audit-binary", type=Path, default=DEFAULT_BINARY)
    parser.add_argument("--python-source", type=Path, default=PYTHON_SOURCE)
    parser.add_argument("--report", type=Path, default=DEFAULT_REPORT)
    parser.add_argument("--rows", type=int, default=10000)
    parser.add_argument("--rounds", type=int, default=3)
    args = parser.parse_args()
    if args.rows < 1 or args.rounds < 1:
        parser.error("--rows and --rounds must be positive")
    report = run_profile(args.audit_binary, args.python_source, args.report, args.rows, args.rounds)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
