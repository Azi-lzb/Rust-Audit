#!/usr/bin/env python3
"""S3-F4 10k-row A/B profile with complete output-workbook comparison."""

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
DEFAULT_REPORT = AUDIT_ROOT / "probe" / "reports" / "s3f4_parity_profile_20261030.json"


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


def load_perf_helper():
    return load_module(
        "audit_s3f5_compare_helpers",
        AUDIT_ROOT / "tests" / "compare" / "perf_s3f5_parity_20261030.py",
    )


def load_baseline_builder():
    return load_module(
        "audit_perf_compare_20261030_for_s3f4",
        AUDIT_ROOT / "tests" / "compare" / "perf_compare_20261030.py",
    )


def python_run(source_root: Path, current_csv: Path, common: Path, cross: Path, output: Path) -> dict[str, Any]:
    sys.path.insert(0, str(source_root / "core" / "src"))
    from base_audit.systems.s3_central_statistics.service import run_cross_period_check

    events: list[dict[str, Any]] = []
    started = time.perf_counter()
    result = run_cross_period_check(
        current_csv=current_csv,
        output_path=output,
        config_path=[common, cross],
        write_flow_logs=False,
        on_step=lambda message: events.append(
            {"message": message, "elapsed_ms": round((time.perf_counter() - started) * 1000, 2)}
        ),
    )
    elapsed = (time.perf_counter() - started) * 1000
    return {
        "elapsed_ms": round(elapsed, 2),
        "result_elapsed_ms": round(float(result.seconds) * 1000, 2),
        "output": Path(result.output_path),
        "rows": len(result.rows),
        "logs": len(result.log_rows),
        "events": events,
    }


def rust_command(binary: Path, current_csv: Path, common: Path, cross: Path, output: Path) -> list[str]:
    return [
        str(binary),
        "s3-check-current-values",
        "--current-csv",
        str(current_csv),
        "--config-workbook",
        str(common),
        "--config-workbook",
        str(cross),
        "--output-workbook",
        str(output),
        "--developer-view",
    ]


def rust_run(binary: Path, current_csv: Path, common: Path, cross: Path, output: Path) -> dict[str, Any]:
    started = time.perf_counter()
    completed = subprocess.run(
        rust_command(binary, current_csv, common, cross, output),
        check=True,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    wall = (time.perf_counter() - started) * 1000
    report = json.loads(completed.stdout)
    events = [json.loads(line) for line in completed.stderr.splitlines() if line.strip()]
    completed_event = next(
        (event for event in reversed(events) if event.get("status") == "completed"),
        {},
    )
    report["wall_ms_including_cli_startup"] = round(wall, 2)
    report["stage_timings_ms"] = {
        key: completed_event.get(key)
        for key in (
            "config_read_time_ms",
            "input_read_time_ms",
            "calculation_time_ms",
            "workbook_write_time_ms",
            "elapsed_ms",
        )
    }
    report["progress"] = completed_event
    return report


def csv_info(path: Path) -> dict[str, Any]:
    import csv

    with path.open("r", encoding="utf-8-sig", newline="") as stream:
        reader = csv.reader(stream)
        header = next(reader)
        rows = sum(1 for row in reader if any(value.strip() for value in row))
    return {"bytes": path.stat().st_size, "header_columns": len(header), "data_rows": rows, "sha256": sha256(path)}


def workbook_info(path: Path) -> dict[str, Any]:
    from openpyxl import load_workbook

    workbook = load_workbook(path, read_only=True, data_only=False)
    try:
        return {
            "bytes": path.stat().st_size,
            "sheets": [
                {"name": sheet.title, "rows": sheet.max_row, "columns": sheet.max_column}
                for sheet in workbook.worksheets
            ],
        }
    finally:
        workbook.close()


def monitored_tree(command: list[str], cwd: Path | None = None) -> dict[str, Any]:
    try:
        import psutil
    except ImportError as error:  # pragma: no cover - workstation dependency
        raise RuntimeError("psutil is required for process-tree peak-working-set measurement") from error

    with tempfile.TemporaryFile() as stdout_file, tempfile.TemporaryFile() as stderr_file:
        process = subprocess.Popen(command, cwd=cwd, stdout=stdout_file, stderr=stderr_file)
        try:
            root = psutil.Process(process.pid)
        except psutil.Error:
            root = None
        peak_bytes = 0
        peak_count = 0
        while process.poll() is None:
            try:
                children = root.children(recursive=True) if root is not None else []
                current = [root, *children] if root is not None else []
            except psutil.Error:
                current = [root] if root is not None else []
            current_bytes = 0
            count = 0
            for item in current:
                try:
                    info = item.memory_info()
                    current_bytes += int(getattr(info, "wset", info.rss))
                    count += 1
                except psutil.Error:
                    pass
            peak_bytes = max(peak_bytes, current_bytes)
            peak_count = max(peak_count, count)
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
        "peak_process_tree_working_set_bytes": peak_bytes,
        "peak_live_process_count": peak_count,
    }


def stage_median(reports: list[dict[str, Any]]) -> dict[str, float]:
    names = sorted({name for report in reports for name in report.get("stage_timings_ms", {})})
    return {
        name: round(statistics.median(float(report["stage_timings_ms"].get(name, 0) or 0) for report in reports), 2)
        for name in names
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--audit-binary", type=Path)
    parser.add_argument("--source-root", type=Path, default=DEFAULT_PYTHON_SOURCE)
    parser.add_argument("--report", type=Path, default=DEFAULT_REPORT)
    parser.add_argument("--rows", type=int, default=10000)
    parser.add_argument("--rounds", type=int, default=3)
    parser.add_argument("--memory-worker", choices=("python",))
    parser.add_argument("--current-csv", type=Path)
    parser.add_argument("--common-config", type=Path)
    parser.add_argument("--cross-config", type=Path)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()

    if args.memory_worker:
        if not all((args.current_csv, args.common_config, args.cross_config, args.output)):
            parser.error("memory worker requires current/config/output paths")
        result = python_run(
            args.source_root.resolve(),
            args.current_csv.resolve(),
            args.common_config.resolve(),
            args.cross_config.resolve(),
            args.output.resolve(),
        )
        print(json.dumps({"status": "completed", "rows": result["rows"]}, ensure_ascii=True))
        return 0

    if args.audit_binary is None or not args.audit_binary.is_file():
        parser.error("--audit-binary must point to an existing Release executable")
    if not args.source_root.is_dir():
        parser.error(f"Python source directory does not exist: {args.source_root}")
    if args.rows < 1 or args.rounds < 1:
        parser.error("--rows and --rounds must be positive")

    baseline = load_baseline_builder()
    helper = load_perf_helper()
    comparator = helper.compare_workbooks
    source_root = args.source_root.resolve()
    source_files = [
        "core/src/base_audit/systems/s3_central_statistics/service.py",
        "core/src/base_audit/systems/s3_central_statistics/cross_period_engine.py",
        "core/src/base_audit/systems/s3_central_statistics/csv_importer.py",
    ]
    source_hashes_before = {name: sha256(source_root / name) for name in source_files}
    python_runs: list[dict[str, Any]] = []
    rust_runs: list[dict[str, Any]] = []
    correctness_rounds: list[dict[str, Any]] = []

    with tempfile.TemporaryDirectory(prefix="audit-s3f4-parity-profile-") as temporary:
        root = Path(temporary)
        fixture = root / "fixture"
        fixture.mkdir()
        common, _feature, cross = baseline.build_s3_configs(fixture, with_cross=True)
        current = fixture / "current.csv"
        baseline.build_s3_csv(current, args.rows, date_span=1)
        inputs = {"current_csv": current, "common_config": common, "cross_config": cross}
        hashes_before = {name: sha256(path) for name, path in inputs.items()}
        sizes = {
            "current_csv": csv_info(current),
            "common_config": workbook_info(common),
            "cross_config": workbook_info(cross),
        }

        warm_python = python_run(source_root, current, common, cross, root / "warm-python.xlsx")
        warm_rust = rust_run(args.audit_binary.resolve(), current, common, cross, root / "warm-rust.xlsx")
        if warm_python["rows"] <= 0 or int(warm_rust.get("rows", 0)) != warm_python["rows"]:
            raise AssertionError(
                f"S3-F4 warm-up row count mismatch: Python={warm_python['rows']}, Rust={warm_rust.get('rows')}"
            )
        warm_comparison = comparator(warm_python["output"], Path(warm_rust["output_workbook"]), 1e-12)
        if warm_comparison["difference_count"]:
            raise AssertionError(f"S3-F4 warm-up output mismatch: {warm_comparison['differences'][:10]!r}")

        python_times: list[float] = []
        for index in range(args.rounds):
            python_output = root / f"python-{index + 1}.xlsx"
            rust_output = root / f"rust-{index + 1}.xlsx"
            if index % 2 == 0:
                python_result = python_run(source_root, current, common, cross, python_output)
                rust_result = rust_run(args.audit_binary.resolve(), current, common, cross, rust_output)
            else:
                rust_result = rust_run(args.audit_binary.resolve(), current, common, cross, rust_output)
                python_result = python_run(source_root, current, common, cross, python_output)
            python_runs.append(python_result)
            rust_runs.append(rust_result)
            python_times.append(python_result["elapsed_ms"])

            if int(rust_result.get("rows", 0)) != python_result["rows"]:
                raise AssertionError(
                    f"S3-F4 round {index + 1} output row count differs: "
                    f"Python={python_result['rows']}, Rust={rust_result.get('rows')}"
                )
            comparison = comparator(python_result["output"], Path(rust_result["output_workbook"]), 1e-12)
            if comparison["difference_count"]:
                raise AssertionError(
                    f"S3-F4 round {index + 1} output mismatch: {comparison['differences'][:10]!r}"
                )
            correctness_rounds.append(
                {
                    "round": index + 1,
                    "python_result_rows": python_result["rows"],
                    "rust_result_rows": rust_result.get("rows"),
                    "comparison": comparison,
                    "python_output": workbook_info(python_result["output"]),
                    "rust_output": workbook_info(Path(rust_result["output_workbook"])),
                    "rust_stages_ms": rust_result["stage_timings_ms"],
                }
            )

        hashes_after = {name: sha256(path) for name, path in inputs.items()}
        if hashes_after != hashes_before:
            raise AssertionError("S3-F4 modified its CSV or configuration inputs")

        memory_python_output = root / "memory-python.xlsx"
        python_memory_command = [
            sys.executable,
            str(Path(__file__).resolve()),
            "--memory-worker",
            "python",
            "--source-root",
            str(source_root),
            "--current-csv",
            str(current),
            "--common-config",
            str(common),
            "--cross-config",
            str(cross),
            "--output",
            str(memory_python_output),
        ]
        python_memory = monitored_tree(python_memory_command, cwd=WORKSPACE)
        if python_memory["return_code"] != 0:
            raise RuntimeError(f"S3-F4 Python memory run failed: {python_memory['stderr'][-4000:]}")
        rust_memory = monitored_tree(
            rust_command(args.audit_binary.resolve(), current, common, cross, root / "memory-rust.xlsx")
        )
        if rust_memory["return_code"] != 0:
            raise RuntimeError(f"S3-F4 Rust memory run failed: {rust_memory['stderr'][-4000:]}")

    source_hashes_after = {name: sha256(source_root / name) for name in source_files}
    if source_hashes_after != source_hashes_before:
        raise AssertionError("read-only Python source hashes changed during the S3-F4 profile")

    python_median = statistics.median(python_times)
    rust_wall = [float(run["wall_ms_including_cli_startup"]) for run in rust_runs]
    rust_internal = [float(run["total_time_ms"]) for run in rust_runs]
    report = {
        "schema_version": 1,
        "date": datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds"),
        "feature": "S3-F4",
        "scenario": f"current-only-{args.rows}-rows-two-cross-rules",
        "platform": platform.platform(),
        "python_version": platform.python_version(),
        "python_source": str(source_root),
        "python_source_sha256_before": source_hashes_before,
        "python_source_sha256_after": source_hashes_after,
        "rust_binary": str(args.audit_binary.resolve()),
        "rust_binary_sha256": sha256(args.audit_binary),
        "inputs": sizes,
        "input_sha256_before": hashes_before,
        "input_sha256_after": hashes_after,
        "timing_scope": {
            "python": "run_cross_period_check service wall time, flow-log output disabled in both engines",
            "rust_wall": "CLI subprocess wall time including process startup",
            "rust_internal": "S3-F04 report total_time_ms",
            "warmup_runs_excluded": 1,
            "round_order": "alternating Python-first and Rust-first",
        },
        "timings": {
            "python_service_ms": python_times,
            "python_median_ms": round(python_median, 2),
            "rust_cli_wall_ms": rust_wall,
            "rust_cli_wall_median_ms": round(statistics.median(rust_wall), 2),
            "rust_internal_ms": rust_internal,
            "rust_internal_median_ms": round(statistics.median(rust_internal), 2),
            "rust_stage_median_ms": stage_median(rust_runs),
            "python_over_rust_cli_wall": round(python_median / statistics.median(rust_wall), 3),
            "python_over_rust_internal": round(python_median / statistics.median(rust_internal), 3),
        },
        "warmup": {
            "python_rows": warm_python["rows"],
            "rust_rows": warm_rust.get("rows"),
            "difference_count": warm_comparison["difference_count"],
        },
        "correctness": {
            "rounds": correctness_rounds,
            "all_output_workbooks_value_formula_and_key_structure_match": True,
            "input_hashes_unchanged": hashes_after == hashes_before,
            "python_source_hashes_unchanged": source_hashes_after == source_hashes_before,
        },
        "process_tree_peak_memory": {
            "measurement": "sampled aggregate working set of parent and live child processes in a separate run; process startup included",
            "sampling_interval_ms": 5,
            "python_working_set_bytes": python_memory["peak_process_tree_working_set_bytes"],
            "python_peak_live_process_count": python_memory["peak_live_process_count"],
            "rust_cli_working_set_bytes": rust_memory["peak_process_tree_working_set_bytes"],
            "rust_peak_live_process_count": rust_memory["peak_live_process_count"],
        },
        "status": "passed",
    }
    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(
        json.dumps(
            {
                "status": report["status"],
                "report": str(args.report),
                "timings": report["timings"],
                "correctness": {
                    "rounds": len(correctness_rounds),
                    "rows": warm_python["rows"],
                    "zero_differences": all(item["comparison"]["difference_count"] == 0 for item in correctness_rounds),
                },
                "memory": report["process_tree_peak_memory"],
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
