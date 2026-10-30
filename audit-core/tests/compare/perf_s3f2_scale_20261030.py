#!/usr/bin/env python3
"""Focused S3-F2 scale profile with Python/Rust output parity checks."""

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

from openpyxl import Workbook


WORKSPACE = Path(__file__).resolve().parents[3]
DEFAULT_PYTHON_SOURCE = Path(r"C:\Users\AZI\Desktop\VibeCode\3.open-data-audit")

HEADERS = [
    "机构类代码", "机构类名称", "地区代码", "地区名称", "指标代码", "指标名称",
    "数据值", "上期值", "输出说明", "说明内容", "增减额（先出现）", "增减额（后出现）",
]


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def load_parity_helper():
    helper_path = Path(__file__).with_name("compare_s3_explanation_export.py")
    spec = importlib.util.spec_from_file_location("s3f2_parity_helper", helper_path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load S3-F2 parity helper: {helper_path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def build_inputs(comparison_path: Path, config_path: Path, rows: int) -> int:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "对比结果"
    sheet.append(HEADERS)
    selected_count = 0
    for index in range(rows):
        selected = "是" if index % 3 == 0 else "否"
        selected_count += selected == "是"
        sheet.append(
            [
                f"{7000001 + index % 400:07d}",
                f"机构{index % 400:04d}",
                f"{4401 + index % 5}",
                "地区",
                f"I-{index % 50:03d}",
                "指标",
                10.0 + index % 7,
                9.0 + index % 5,
                selected,
                f"说明{index}",
                1.0,
                1,
            ]
        )
    workbook.save(comparison_path)
    workbook.close()

    mapping = Workbook()
    reference = mapping.active
    reference.title = "机构地区参照"
    reference.append(["机构类代码", "地区代码", "机构名称"])
    for index in range(400):
        reference.append(
            [f"{7000001 + index:07d}", f"{4401 + index % 5}", f"精确机构{index:04d}"]
        )
    mapping.save(config_path)
    mapping.close()
    return selected_count


def run_rust(binary: Path, comparison: Path, config: Path, output: Path) -> dict:
    command = [
        str(binary),
        "s3-export-explanations",
        "--comparison-file",
        str(comparison),
        "--output-workbook",
        str(output),
        "--config-workbook",
        str(config),
        "--developer-view",
    ]
    started = time.perf_counter()
    completed = subprocess.run(
        command,
        check=True,
        capture_output=True,
        text=True,
        encoding="utf-8",
    )
    wall_ms = (time.perf_counter() - started) * 1000
    report = json.loads(completed.stdout)
    events = [json.loads(line) for line in completed.stderr.splitlines() if line.strip()]
    if [event.get("status") for event in events] != ["started", "completed"]:
        raise AssertionError(f"unexpected Rust progress events: {events!r}")
    report["wall_ms_including_cli_startup"] = round(wall_ms, 2)
    report["stage_timings_ms"] = {
        key: events[-1].get(key)
        for key in (
            "comparison_read_time_ms",
            "config_read_time_ms",
            "workbook_write_time_ms",
            "elapsed_ms",
        )
    }
    if not all(isinstance(value, int) for value in report["stage_timings_ms"].values()):
        raise AssertionError(f"missing Rust stage timing payload: {events!r}")
    return report


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--audit-binary", required=True, type=Path)
    parser.add_argument("--source-root", type=Path, default=DEFAULT_PYTHON_SOURCE)
    parser.add_argument(
        "--report",
        type=Path,
        default=WORKSPACE / "audit-core" / "probe" / "reports" / "s3f2_scale_profile_20261030.json",
    )
    parser.add_argument("--rows", type=int, default=7830)
    parser.add_argument("--rounds", type=int, default=3)
    args = parser.parse_args()
    if args.rows < 3 or args.rounds < 1:
        parser.error("--rows must be at least 3 and --rounds must be positive")
    if not args.audit_binary.is_file():
        parser.error(f"Audit binary does not exist: {args.audit_binary}")
    if not args.source_root.is_dir():
        parser.error(f"read-only Python project does not exist: {args.source_root}")

    parity = load_parity_helper()
    python_export = parity.load_baseline(args.source_root)
    result: dict = {
        "schema_version": 1,
        "date": datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds"),
        "feature": "S3-F2",
        "platform": platform.platform(),
        "python_version": platform.python_version(),
        "python_source": str(args.source_root),
        "rust_binary": str(args.audit_binary),
        "rust_binary_sha256": sha256(args.audit_binary),
        "fixture": {
            "kind": "synthetic_scaled_from_perf_compare_20261030",
            "comparison_workbooks": 1,
            "comparison_sheets": 1,
            "data_rows": args.rows,
            "selected_rows": sum(1 for index in range(args.rows) if index % 3 == 0),
            "comparison_columns": len(HEADERS),
            "mapping_workbooks": 1,
            "mapping_sheets": 1,
            "mapping_rows_including_header": 401,
            "rounds": args.rounds,
        },
        "status": "failed",
    }

    try:
        with tempfile.TemporaryDirectory(prefix="audit-s3f2-scale-profile-") as temporary:
            root = Path(temporary)
            comparison = root / "comparison.xlsx"
            config = root / "mapping.xlsx"
            selected_count = build_inputs(comparison, config, args.rows)
            input_hashes_before = {str(comparison): sha256(comparison), str(config): sha256(config)}
            result["fixture"].update(
                {
                    "selected_rows": selected_count,
                    "comparison_bytes": comparison.stat().st_size,
                    "mapping_bytes": config.stat().st_size,
                    "input_sha256": input_hashes_before,
                }
            )

            warm_python = root / "warm-python.xlsx"
            warm_rust = root / "warm-rust.xlsx"
            python_export(comparison, warm_python, config_path=config)
            run_rust(args.audit_binary, comparison, config, warm_rust)
            warm_python.unlink()
            warm_rust.unlink()

            python_times: list[float] = []
            rust_wall_times: list[float] = []
            rust_stage_reports: list[dict] = []
            parity_rounds: list[dict] = []
            for index in range(args.rounds):
                python_output = root / f"python-{index}.xlsx"
                rust_output = root / f"rust-{index}.xlsx"

                def run_python() -> int:
                    started = time.perf_counter()
                    count = python_export(comparison, python_output, config_path=config)
                    python_times.append(round((time.perf_counter() - started) * 1000, 2))
                    return count

                def run_rust_round() -> dict:
                    report = run_rust(args.audit_binary, comparison, config, rust_output)
                    rust_wall_times.append(report["wall_ms_including_cli_startup"])
                    rust_stage_reports.append(report["stage_timings_ms"])
                    return report

                if index % 2 == 0:
                    python_count = run_python()
                    rust_report = run_rust_round()
                else:
                    rust_report = run_rust_round()
                    python_count = run_python()

                if python_count != selected_count:
                    raise AssertionError(f"Python selected {python_count}, expected {selected_count}")
                if rust_report.get("status") != "completed" or rust_report.get("selected_rows") != selected_count:
                    raise AssertionError(f"Rust result does not match selected-row count: {rust_report!r}")
                differences, output_rows = parity.compare_workbooks(python_output, rust_output)
                if differences or output_rows != selected_count:
                    raise AssertionError(
                        f"round {index + 1} output mismatch: rows={output_rows}, differences={differences[:20]!r}"
                    )
                parity_rounds.append(
                    {
                        "round": index + 1,
                        "output_rows_compared": output_rows,
                        "differences": len(differences),
                        "python_output_bytes": python_output.stat().st_size,
                        "rust_output_bytes": rust_output.stat().st_size,
                        "rust_stage_timings_ms": rust_report["stage_timings_ms"],
                    }
                )
                python_output.unlink()
                rust_output.unlink()

            input_hashes_after = {str(comparison): sha256(comparison), str(config): sha256(config)}
            if input_hashes_after != input_hashes_before:
                raise AssertionError("S3-F2 modified an input workbook")

            python_median = statistics.median(python_times)
            rust_wall_median = statistics.median(rust_wall_times)
            rust_internal_median = statistics.median(
                report["elapsed_ms"] for report in rust_stage_reports
            )
            result.update(
                {
                    "status": "passed",
                    "timing_scope": {
                        "python": "in-process function call; excludes interpreter startup and fixture/output comparison",
                        "rust_wall": "Release CLI process wall time; includes process startup and excludes fixture/output comparison",
                        "rust_internal": "Rust developer-view elapsed_ms; excludes CLI process startup",
                        "warmup_runs_excluded": 1,
                        "round_order": "alternates Python/Rust first",
                    },
                    "python_ms": python_times,
                    "rust_wall_ms": rust_wall_times,
                    "rust_internal_elapsed_ms": [report["elapsed_ms"] for report in rust_stage_reports],
                    "python_median_ms": round(python_median, 2),
                    "rust_wall_median_ms": round(rust_wall_median, 2),
                    "rust_internal_median_ms": round(rust_internal_median, 2),
                    "python_over_rust_wall_speedup": round(python_median / rust_wall_median, 3),
                    "python_over_rust_internal_speedup": round(python_median / rust_internal_median, 3),
                    "rust_stage_timings_ms": rust_stage_reports,
                    "parity_rounds": parity_rounds,
                    "all_output_rounds_match": True,
                    "all_inputs_unchanged": True,
                }
            )
    except Exception as error:
        result["error"] = str(error)
        args.report.parent.mkdir(parents=True, exist_ok=True)
        args.report.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        raise

    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
