#!/usr/bin/env python3
"""Capture S1-F2 Rust phase timings on the established synthetic fixture.

This is a Rust-only diagnostic so phase timing can be collected without
starting Excel COM. Python/Rust total-time and output-parity evidence remains
in compare_s1_region_summary.py and its existing reports.
"""

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

import openpyxl
import psutil

sys.dont_write_bytecode = True

AUDIT_ROOT = Path(__file__).resolve().parents[2]
PARITY_DRIVER = AUDIT_ROOT / "tests" / "compare" / "compare_s1_region_summary.py"
TEMPLATE = AUDIT_ROOT / "tests" / "fixtures" / "s1_region_summary_template.xlsx"
DEFAULT_BINARY = Path(r"F:\RustBuild\audit\release\open-data-audit.exe")
DEFAULT_REPORT = AUDIT_ROOT / "probe" / "reports" / (
    "s1f2_stage_profile_"
    + datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S_%f_utc")
    + ".json"
)
STAGE_FIELDS = (
    "discovery_time_ms",
    "template_read_time_ms",
    "structure_preflight_time_ms",
    "source_read_time_ms",
    "calculation_time_ms",
    "history_elapsed_ms",
    "write_time_ms",
    "validation_time_ms",
    "elapsed_ms",
)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def load_parity_helpers() -> Any:
    spec = importlib.util.spec_from_file_location("s1f2_stage_parity_helpers", PARITY_DRIVER)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load S1-F2 fixture helpers: {PARITY_DRIVER}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def workbook_snapshot(path: Path) -> tuple[Any, ...]:
    workbook = openpyxl.load_workbook(path, data_only=False, read_only=True)
    try:
        return tuple(
            (
                sheet.title,
                sheet.max_row,
                sheet.max_column,
                tuple(tuple(cell.value for cell in row) for row in sheet.iter_rows()),
            )
            for sheet in workbook.worksheets
        )
    finally:
        workbook.close()


def background_process_snapshot() -> list[dict[str, Any]]:
    found: list[dict[str, Any]] = []
    for process in psutil.process_iter(["pid", "name", "cmdline", "create_time"]):
        try:
            name = process.info.get("name") or ""
            command_line = " ".join(process.info.get("cmdline") or [])
            if name.lower() == "excel.exe" or "xls-to-xlsx-com" in command_line.lower():
                found.append(
                    {
                        "pid": process.info.get("pid"),
                        "name": name,
                        "command_line": command_line[:400],
                        "create_time": process.info.get("create_time"),
                    }
                )
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    return found


def run_profile(binary: Path, report_path: Path, rounds: int) -> dict[str, Any]:
    if report_path.exists():
        raise FileExistsError(f"refusing to overwrite benchmark report: {report_path}")
    if not binary.is_file():
        raise FileNotFoundError(f"Rust Audit binary not found: {binary}")
    if not TEMPLATE.is_file():
        raise FileNotFoundError(f"S1-F2 template fixture not found: {TEMPLATE}")

    background = background_process_snapshot()
    helpers = load_parity_helpers()
    rounds_data: list[dict[str, Any]] = []
    snapshots: list[tuple[Any, ...]] = []

    with tempfile.TemporaryDirectory(prefix="audit-s1f2-stage-profile-") as temporary:
        root = Path(temporary)
        input_dir = root / "input"
        input_dir.mkdir()
        alpha = input_dir / "!Alpha_20260731.xlsx"
        beta = input_dir / "Beta_20260731.xlsx"
        helpers.write_source(
            alpha,
            {4: ("甲", 10), 5: (None, 0), 7: ("乙", 20)},
            "甲组",
            "批次甲",
        )
        helpers.write_source(beta, {4: ("丙", 30), 5: ("丁", 40)}, "乙组")
        inputs = [TEMPLATE, alpha, beta]
        hashes_before = {str(path): sha256(path) for path in inputs}
        sizes = {str(path): path.stat().st_size for path in inputs}

        baseline_snapshot: tuple[Any, ...] | None = None
        for index in range(rounds):
            output = root / f"rust-{index + 1}.xlsx"
            command = [
                str(binary),
                "s1-region-summary",
                "--template-workbook",
                str(TEMPLATE),
                "--input-dir",
                str(input_dir),
                "--output-workbook",
                str(output),
                "--feature-name",
                "任意行汇总",
                "--feature-name",
                "固定行汇总",
                "--developer-view",
            ]
            started = time.perf_counter()
            completed = subprocess.run(
                command,
                check=False,
                capture_output=True,
                text=True,
                encoding="utf-8",
            )
            cli_wall_ms = (time.perf_counter() - started) * 1000
            if completed.returncode:
                raise RuntimeError(
                    "Rust S1-F2 stage profile failed: "
                    + completed.stderr[-5000:]
                )

            payload = json.loads(completed.stdout)
            summary = payload.get("summary", {})
            missing = [field for field in STAGE_FIELDS if field not in summary]
            if missing:
                raise RuntimeError(f"Rust S1-F2 report missing stage fields: {missing}")
            if (
                payload.get("status") != "completed"
                or summary.get("source_workbooks") != 2
                or summary.get("source_sheets") != 3
                or summary.get("output_sheets") != 2
                or summary.get("data_rows") != 9
                or not output.is_file()
            ):
                raise AssertionError(f"unexpected S1-F2 output summary: {summary!r}")

            snapshot = workbook_snapshot(output)
            if baseline_snapshot is None:
                baseline_snapshot = snapshot
            elif snapshot != baseline_snapshot:
                raise AssertionError("S1-F2 Rust output values or structure changed between rounds")
            snapshots.append(snapshot)
            rounds_data.append(
                {
                    "round": index + 1,
                    "rust_cli_wall_ms": round(cli_wall_ms, 3),
                    "rust_internal_stage_timings_ms": {
                        field: summary[field] for field in STAGE_FIELDS
                    },
                    "reader": summary.get("reader"),
                    "writer": summary.get("writer"),
                    "output_bytes": output.stat().st_size,
                    "output_sha256": sha256(output),
                    "output_structure": [
                        {"sheet": name, "rows": rows, "columns": columns}
                        for name, rows, columns, _ in snapshot
                    ],
                }
            )

        hashes_after = {str(path): sha256(path) for path in inputs}

    eligible = not background
    report = {
        "schema_version": 1,
        "generated_at_utc": datetime.now(timezone.utc).isoformat(),
        "feature": "S1-F2",
        "profile_kind": "Rust-only phase diagnostic; no new Python/Rust ratio",
        "platform": platform.platform(),
        "python_version": platform.python_version(),
        "rust_binary": str(binary),
        "rust_binary_sha256": sha256(binary),
        "fixture": {
            "generator": str(PARITY_DRIVER),
            "template_and_mapping_config_bytes": sizes[str(TEMPLATE)],
            "source_workbooks": 2,
            "source_sheets": 3,
            "source_data_rows": 5,
            "source_bytes": {path.name: sizes[str(path)] for path in (alpha, beta)},
            "output_sheets": 2,
            "output_data_rows": 9,
            "sha256_before": hashes_before,
            "sha256_after": hashes_after,
        },
        "measurement_environment": {
            "performance_eligible_for_baseline": eligible,
            "reason": None
            if eligible
            else "Excel COM processes or the Rust-Tools XLS-to-XLSX runner were active during measurement; use only as stage identification.",
            "observed_background_processes": background,
        },
        "method": {
            "rounds": rounds,
            "timing_resolution": "Rust internal stages are integer milliseconds; CLI wall time has microsecond-derived milliseconds.",
            "rust_cli_wall_scope": "fresh Release CLI process; excludes fixture generation and output snapshot comparison",
            "rust_internal_scope": "discovery, template read, structure preflight, source read, calculation, history, writing, validation and total as emitted by S1-F2 summary",
            "correctness_scope": "expected counts and output workbook readability/shape; output cell snapshot identical across Rust rounds; prior Python/Rust 3-round full cell/structure parity remains the cross-runtime evidence",
            "source_policy": "synthetic temporary inputs only; Python source project is imported read-only for the established fixture generator",
        },
        "summary": {
            "rust_cli_wall_median_ms": round(statistics.median(row["rust_cli_wall_ms"] for row in rounds_data), 3),
            "rust_stage_medians_ms": {
                field: round(
                    statistics.median(row["rust_internal_stage_timings_ms"][field] for row in rounds_data),
                    3,
                )
                for field in STAGE_FIELDS
            },
            "output_rounds_identical": len(snapshots) == rounds,
            "all_inputs_unchanged": hashes_before == hashes_after,
        },
        "rounds": rounds_data,
    }

    report_path.parent.mkdir(parents=True, exist_ok=True)
    with report_path.open("x", encoding="utf-8", newline="\n") as stream:
        stream.write(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    return report


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--audit-binary", type=Path, default=DEFAULT_BINARY)
    parser.add_argument("--report", type=Path, default=DEFAULT_REPORT)
    parser.add_argument("--rounds", type=int, default=3)
    args = parser.parse_args()
    if args.rounds < 1:
        parser.error("--rounds must be positive")
    report = run_profile(args.audit_binary, args.report, args.rounds)
    print(json.dumps(report["summary"], ensure_ascii=False, indent=2))
    return 0 if report["summary"]["all_inputs_unchanged"] else 2


if __name__ == "__main__":
    raise SystemExit(main())
