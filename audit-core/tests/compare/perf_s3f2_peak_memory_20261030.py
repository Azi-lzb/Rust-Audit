#!/usr/bin/env python3
"""Memory-only Python/Rust process-tree profile for the S3-F2 scaled fixture."""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import os
import platform
import statistics
import subprocess
import sys
import tempfile
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import psutil


WORKSPACE = Path(__file__).resolve().parents[3]
AUDIT_ROOT = WORKSPACE / "audit-core"
PYTHON_SOURCE = Path(r"C:\Users\AZI\Desktop\VibeCode\3.open-data-audit")
DEFAULT_BINARY = Path(r"F:\RustBuild\audit\release\open-data-audit.exe")
DEFAULT_REPORT = AUDIT_ROOT / "probe" / "reports" / "s3f2_peak_memory_20261030.json"
SCALE_DRIVER = AUDIT_ROOT / "tests" / "compare" / "perf_s3f2_scale_20261030.py"
PARITY_HELPER = AUDIT_ROOT / "tests" / "compare" / "compare_s3_explanation_export.py"


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


def process_tree_memory(process: subprocess.Popen[bytes], sample_interval_s: float = 0.005) -> dict[str, Any]:
    root = psutil.Process(process.pid)
    peak_sampled_tree_rss = 0
    peak_process_count = 0
    per_process_peak_wset: dict[int, int] = {}
    sample_count = 0

    while True:
        alive: list[psutil.Process] = []
        try:
            alive = [root, *root.children(recursive=True)]
        except psutil.Error:
            try:
                alive = [root]
            except psutil.Error:
                alive = []

        sampled_tree_rss = 0
        live_count = 0
        for child in alive:
            try:
                memory = child.memory_info()
            except psutil.Error:
                continue
            live_count += 1
            sampled_tree_rss += int(memory.rss)
            peak_wset = getattr(memory, "peak_wset", memory.rss)
            per_process_peak_wset[child.pid] = max(
                per_process_peak_wset.get(child.pid, 0), int(peak_wset)
            )
        sample_count += 1
        peak_sampled_tree_rss = max(peak_sampled_tree_rss, sampled_tree_rss)
        peak_process_count = max(peak_process_count, live_count)
        if process.poll() is not None:
            break
        time.sleep(sample_interval_s)

    stdout, stderr = process.communicate()
    return {
        "return_code": process.returncode,
        "stdout": stdout.decode("utf-8", errors="replace"),
        "stderr": stderr.decode("utf-8", errors="replace"),
        "peak_sampled_process_tree_working_set_bytes": peak_sampled_tree_rss,
        "peak_sampled_process_tree_working_set_mb": round(peak_sampled_tree_rss / (1024 * 1024), 3),
        "sum_per_process_peak_working_set_bytes": sum(per_process_peak_wset.values()),
        "peak_live_process_count": peak_process_count,
        "sample_interval_ms": round(sample_interval_s * 1000, 3),
        "memory_samples": sample_count,
    }


def run_worker(command: list[str]) -> tuple[dict[str, Any], dict[str, Any]]:
    process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    memory = process_tree_memory(process)
    if memory["return_code"] != 0:
        raise RuntimeError(
            f"worker failed ({memory['return_code']}):\n{memory['stderr'][-4000:]}"
        )
    try:
        result = json.loads(memory["stdout"])
    except json.JSONDecodeError as error:
        raise RuntimeError(f"worker returned invalid JSON: {memory['stdout'][-2000:]}") from error
    return result, memory


def run_memory_profile(
    binary: Path,
    source_root: Path,
    report_path: Path,
    rows: int,
    rounds: int,
) -> dict[str, Any]:
    if not binary.is_file():
        raise FileNotFoundError(f"Rust Audit binary not found: {binary}")
    if not source_root.is_dir():
        raise FileNotFoundError(f"Python reference project not found: {source_root}")
    fixture = load_module("s3f2_peak_memory_fixture", SCALE_DRIVER)
    parity = load_module("s3f2_peak_memory_parity", PARITY_HELPER)

    with tempfile.TemporaryDirectory(prefix="audit-s3f2-peak-memory-") as temp_name:
        root = Path(temp_name)
        comparison = root / "comparison.xlsx"
        config = root / "mapping.xlsx"
        selected_rows = fixture.build_inputs(comparison, config, rows)
        input_hashes_before = {str(path): sha256(path) for path in (comparison, config)}
        python_script = Path(__file__).resolve()
        rounds_data = []

        for index in range(rounds):
            python_output = root / f"python-{index + 1}.xlsx"
            rust_output = root / f"rust-{index + 1}.xlsx"

            python_command = [
                sys.executable,
                str(python_script),
                "--worker",
                "python",
                str(comparison),
                str(config),
                str(python_output),
                str(source_root),
            ]
            rust_command = [
                str(binary),
                "s3-export-explanations",
                "--comparison-file",
                str(comparison),
                "--output-workbook",
                str(rust_output),
                "--config-workbook",
                str(config),
                "--developer-view",
            ]

            if index % 2 == 0:
                python_result, python_memory = run_worker(python_command)
                rust_result, rust_memory = run_worker(rust_command)
                order = "Python-first"
            else:
                rust_result, rust_memory = run_worker(rust_command)
                python_result, python_memory = run_worker(python_command)
                order = "Rust-first"

            if python_result.get("selected_rows") != selected_rows:
                raise AssertionError(
                    f"Python selected {python_result.get('selected_rows')}, expected {selected_rows}"
                )
            if rust_result.get("status") != "completed" or rust_result.get("selected_rows") != selected_rows:
                raise AssertionError(f"Rust did not complete the expected output: {rust_result!r}")
            differences, output_rows = parity.compare_workbooks(python_output, rust_output)
            if differences or output_rows != selected_rows:
                raise AssertionError(
                    f"round {index + 1} output mismatch: rows={output_rows}, differences={differences[:20]!r}"
                )

            rounds_data.append(
                {
                    "round": index + 1,
                    "order": order,
                    "output_rows_compared": output_rows,
                    "cell_differences": len(differences),
                    "python_output_bytes": python_output.stat().st_size,
                    "rust_output_bytes": rust_output.stat().st_size,
                    "python_process_tree": python_memory,
                    "rust_process_tree": rust_memory,
                }
            )

        input_hashes_after = {str(path): sha256(path) for path in (comparison, config)}
        if input_hashes_before != input_hashes_after:
            raise AssertionError("S3-F2 modified one or more input workbooks")

        python_peaks = [
            row["python_process_tree"]["peak_sampled_process_tree_working_set_bytes"]
            for row in rounds_data
        ]
        rust_peaks = [
            row["rust_process_tree"]["peak_sampled_process_tree_working_set_bytes"]
            for row in rounds_data
        ]
        report = {
            "schema_version": 1,
            "generated_at_utc": datetime.now(timezone.utc).isoformat(),
            "git_revision": git_revision(),
            "feature": "S3-F2",
            "measurement": "process-tree peak working set only; elapsed-time comparison intentionally omitted",
            "platform": platform.platform(),
            "python_version": platform.python_version(),
            "python_source": str(source_root),
            "rust_binary": str(binary),
            "rust_binary_sha256": sha256(binary),
            "fixture": {
                "kind": "synthetic_scaled_from_perf_s3f2_scale_20261030.py",
                "comparison_workbooks": 1,
                "comparison_sheets": 1,
                "data_rows": rows,
                "selected_rows": selected_rows,
                "comparison_columns": len(fixture.HEADERS),
                "mapping_workbooks": 1,
                "mapping_sheets": 1,
                "mapping_rows_including_header": 401,
                "comparison_bytes": comparison.stat().st_size,
                "mapping_bytes": config.stat().st_size,
                "input_sha256_before": input_hashes_before,
                "input_sha256_after": input_hashes_after,
            },
            "method": {
                "rounds": rounds,
                "round_order": "alternates Python-first and Rust-first",
                "sample_interval_ms": 5,
                "python_process": "fresh interpreter; process tree RSS sampled with psutil",
                "rust_process": "fresh Release CLI; process tree RSS sampled with psutil",
                "peak_metric": "maximum sampled sum of resident working sets across live process tree",
                "measurement_limit": "sampling may miss sub-interval peaks; includes process startup; excludes fixture generation and output comparison",
            },
            "summary": {
                "python_peak_working_set_median_mb": round(statistics.median(python_peaks) / (1024 * 1024), 3),
                "python_peak_working_set_max_mb": round(max(python_peaks) / (1024 * 1024), 3),
                "rust_peak_working_set_median_mb": round(statistics.median(rust_peaks) / (1024 * 1024), 3),
                "rust_peak_working_set_max_mb": round(max(rust_peaks) / (1024 * 1024), 3),
                "python_to_rust_peak_median_ratio": round(statistics.median(python_peaks) / statistics.median(rust_peaks), 3),
                "every_round_output_equal": True,
                "inputs_unchanged": True,
            },
            "rounds": rounds_data,
        }

    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return report


def python_worker(comparison: Path, config: Path, output: Path, source_root: Path) -> int:
    sys.dont_write_bytecode = True
    parity = load_module("s3f2_peak_memory_python_parity", PARITY_HELPER)
    python_export = parity.load_baseline(source_root)
    selected_rows = python_export(comparison, output, config_path=config)
    print(json.dumps({"status": "completed", "selected_rows": selected_rows}))
    return 0


def main() -> int:
    if len(sys.argv) > 1 and sys.argv[1] == "--worker":
        if len(sys.argv) != 7 or sys.argv[2] != "python":
            raise SystemExit("usage: --worker python COMPARISON CONFIG OUTPUT PYTHON_SOURCE_ROOT")
        return python_worker(
            Path(sys.argv[3]),
            Path(sys.argv[4]),
            Path(sys.argv[5]),
            Path(sys.argv[6]),
        )

    parser = argparse.ArgumentParser()
    parser.add_argument("--audit-binary", type=Path, default=DEFAULT_BINARY)
    parser.add_argument("--source-root", type=Path, default=PYTHON_SOURCE)
    parser.add_argument("--report", type=Path, default=DEFAULT_REPORT)
    parser.add_argument("--rows", type=int, default=7830)
    parser.add_argument("--rounds", type=int, default=3)
    args = parser.parse_args()
    if args.rows < 3 or args.rounds < 1:
        parser.error("--rows must be at least 3 and --rounds must be positive")
    report = run_memory_profile(args.audit_binary, args.source_root, args.report, args.rows, args.rounds)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
