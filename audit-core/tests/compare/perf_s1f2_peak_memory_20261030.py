#!/usr/bin/env python3
"""Memory-only S1-F2 Python/Rust process-tree profile."""

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

import psutil


WORKSPACE = Path(__file__).resolve().parents[3]
AUDIT_ROOT = WORKSPACE / "audit-core"
PARITY_DRIVER = AUDIT_ROOT / "tests" / "compare" / "compare_s1_region_summary.py"
DEFAULT_BINARY = Path(r"F:\RustBuild\audit\release\open-data-audit.exe")
DEFAULT_REPORT = AUDIT_ROOT / "probe" / "reports" / "s1f2_peak_memory_20261030.json"


def load_module(name: str, path: Path):
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load module: {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


def digest(path: Path) -> str:
    sha = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            sha.update(chunk)
    return sha.hexdigest()


def git_revision() -> str:
    completed = subprocess.run(
        ["git", "rev-parse", "--short", "HEAD"],
        cwd=WORKSPACE,
        check=True,
        capture_output=True,
        text=True,
    )
    return completed.stdout.strip()


def excel_process_ids() -> set[int]:
    pids = set()
    for process in psutil.process_iter(attrs=["pid", "name"]):
        try:
            name = (process.info.get("name") or "").casefold()
            if name == "excel.exe":
                pids.add(int(process.info["pid"]))
        except (psutil.Error, TypeError, ValueError):
            continue
    return pids


def sample_process_tree(
    process: subprocess.Popen[bytes],
    baseline_excel_pids: set[int],
    sample_interval_s: float = 0.005,
) -> dict[str, Any]:
    root = psutil.Process(process.pid)
    peak_tree_rss = 0
    peak_process_count = 0
    sample_count = 0
    discovered_excel_pids: set[int] = set()

    while True:
        try:
            candidates = [root, *root.children(recursive=True)]
        except psutil.Error:
            candidates = []
        if sample_count % 10 == 0:
            try:
                discovered_excel_pids.update(excel_process_ids() - baseline_excel_pids)
            except psutil.Error:
                pass
        for pid in sorted(discovered_excel_pids):
            try:
                candidates.append(psutil.Process(pid))
            except psutil.Error:
                continue

        by_pid = {candidate.pid: candidate for candidate in candidates}
        sampled_rss = 0
        live_count = 0
        for candidate in by_pid.values():
            try:
                sampled_rss += int(candidate.memory_info().rss)
                live_count += 1
            except psutil.Error:
                continue
        sample_count += 1
        peak_tree_rss = max(peak_tree_rss, sampled_rss)
        peak_process_count = max(peak_process_count, live_count)
        if process.poll() is not None:
            break
        time.sleep(sample_interval_s)

    stdout, stderr = process.communicate()
    remaining_excel_pids = set()
    cleanup_deadline = time.monotonic() + 5.0
    while discovered_excel_pids and time.monotonic() < cleanup_deadline:
        remaining_excel_pids = excel_process_ids() & discovered_excel_pids
        if not remaining_excel_pids:
            break
        time.sleep(0.05)
    return {
        "return_code": process.returncode,
        "stdout": stdout.decode("utf-8", errors="replace"),
        "stderr": stderr.decode("utf-8", errors="replace"),
        "peak_sampled_process_tree_working_set_bytes": peak_tree_rss,
        "peak_sampled_process_tree_working_set_mb": round(peak_tree_rss / (1024 * 1024), 3),
        "peak_live_process_count": peak_process_count,
        "new_excel_process_ids": sorted(discovered_excel_pids),
        "new_excel_process_ids_remaining_after_worker": sorted(remaining_excel_pids),
        "sample_interval_ms": round(sample_interval_s * 1000, 3),
        "memory_samples": sample_count,
    }


def run_measured(command: list[str]) -> dict[str, Any]:
    baseline_excel_pids = excel_process_ids()
    process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    sampled = sample_process_tree(process, baseline_excel_pids)
    if sampled["return_code"] != 0:
        raise RuntimeError(
            f"worker failed ({sampled['return_code']}):\n{sampled['stderr'][-5000:]}"
        )
    try:
        result = json.loads(sampled["stdout"])
    except json.JSONDecodeError as error:
        raise RuntimeError(f"worker returned invalid JSON: {sampled['stdout'][-2000:]}") from error
    sampled["result"] = result
    return sampled


def write_fixture(module: Any, root: Path) -> dict[str, Any]:
    input_dir = root / "input"
    input_dir.mkdir()
    alpha = input_dir / "!Alpha_20260731.xlsx"
    beta = input_dir / "Beta_20260731.xlsx"
    module.write_source(alpha, {4: ("甲", 10), 5: (None, 0), 7: ("乙", 20)}, "甲组", "批次甲")
    module.write_source(beta, {4: ("丙", 30), 5: ("丁", 40)}, "乙组")
    template = AUDIT_ROOT / "tests" / "fixtures" / "s1_region_summary_template.xlsx"
    paths = [template, alpha, beta]
    hashes = {str(path): digest(path) for path in paths}
    return {
        "input_dir": input_dir,
        "template": template,
        "source_files": [alpha, beta],
        "hashes": hashes,
        "sizes": {str(path): path.stat().st_size for path in paths},
    }


def run_python_worker(input_dir: Path, output_dir: Path, template: Path) -> dict[str, Any]:
    command = [
        sys.executable,
        str(Path(__file__).resolve()),
        "--worker",
        "python",
        str(input_dir),
        str(output_dir),
        str(template),
    ]
    return run_measured(command)


def python_worker(input_dir: Path, output_dir: Path, template: Path) -> int:
    module = load_module("s1f2_peak_memory_legacy_driver", PARITY_DRIVER)
    from base_audit.name_config import (
        FIXED_ROW_SUMMARY_FUNCTION,
        USED_RANGE_SUMMARY_FUNCTION,
        FeatureMapping,
    )
    from base_audit.region_summary import run_region_summaries

    features = (
        FeatureMapping(
            "任意行汇总",
            USED_RANGE_SUMMARY_FUNCTION,
            ("任意行汇总区域",),
            False,
            "",
            "",
        ),
        FeatureMapping(
            "固定行汇总",
            FIXED_ROW_SUMMARY_FUNCTION,
            ("固定行汇总区域",),
            False,
            "",
            "",
        ),
    )
    progress: list[str] = []
    result = run_region_summaries(
        template_path=template,
        input_dir=input_dir,
        output_dir=output_dir,
        config_path=template,
        summary_features=features,
        on_step=progress.append,
        engine_preference="自动",
    )
    output_path = Path(result.output_path)
    engine = next(
        (line.split("：", 1)[1] for line in progress if line.startswith("已连接表格引擎：")),
        "not_reported",
    )
    print(json.dumps({"status": "completed", "output_path": str(output_path), "engine": engine}))
    return 0


def run_memory_profile(binary: Path, report_path: Path, rounds: int) -> dict[str, Any]:
    if not binary.is_file():
        raise FileNotFoundError(f"Rust Audit binary not found: {binary}")
    fixture_module = load_module("s1f2_peak_memory_fixture", PARITY_DRIVER)
    with tempfile.TemporaryDirectory(prefix="audit-s1f2-peak-memory-") as temp_name:
        root = Path(temp_name)
        fixture = write_fixture(fixture_module, root)
        python_module = fixture_module
        rounds_data = []
        for index in range(rounds):
            python_output_dir = root / f"python-{index + 1}"
            python_output_dir.mkdir()
            rust_output = root / f"rust-{index + 1}.xlsx"
            rust_command = [
                str(binary),
                "s1-region-summary",
                "--template-workbook",
                str(fixture["template"]),
                "--input-dir",
                str(fixture["input_dir"]),
                "--output-workbook",
                str(rust_output),
                "--feature-name",
                "任意行汇总",
                "--feature-name",
                "固定行汇总",
                "--developer-view",
            ]
            if index % 2 == 0:
                python_sample = run_python_worker(
                    fixture["input_dir"], python_output_dir, fixture["template"]
                )
                rust_sample = run_measured(rust_command)
                order = "Python-first"
            else:
                rust_sample = run_measured(rust_command)
                python_sample = run_python_worker(
                    fixture["input_dir"], python_output_dir, fixture["template"]
                )
                order = "Rust-first"

            python_result = python_sample["result"]
            rust_result = rust_sample["result"]
            if python_result.get("status") != "completed":
                raise AssertionError(f"Python S1-F2 did not complete: {python_result!r}")
            if rust_result.get("status") != "completed":
                raise AssertionError(f"Rust S1-F2 did not complete: {rust_result!r}")
            python_output = Path(python_result["output_path"])
            comparison = python_module.compare_workbooks(python_output, rust_output)
            if comparison.get("status") != "passed" or comparison.get("value_mismatches") != 0:
                raise AssertionError(f"S1-F2 output comparison failed: {comparison!r}")
            rounds_data.append(
                {
                    "round": index + 1,
                    "order": order,
                    "python_engine": python_result.get("engine"),
                    "python_process_tree": {
                        key: value
                        for key, value in python_sample.items()
                        if key not in {"stdout", "stderr", "result"}
                    },
                    "rust_process_tree": {
                        key: value
                        for key, value in rust_sample.items()
                        if key not in {"stdout", "stderr", "result"}
                    },
                    "output_comparison": comparison,
                }
            )

        hashes_after = {str(path): digest(path) for path in [fixture["template"], *fixture["source_files"]]}
        if fixture["hashes"] != hashes_after:
            raise AssertionError("S1-F2 changed the template or a source workbook")

        python_peaks = [
            item["python_process_tree"]["peak_sampled_process_tree_working_set_bytes"]
            for item in rounds_data
        ]
        rust_peaks = [
            item["rust_process_tree"]["peak_sampled_process_tree_working_set_bytes"]
            for item in rounds_data
        ]
        report = {
            "schema_version": 1,
            "generated_at_utc": datetime.now(timezone.utc).isoformat(),
            "git_revision": git_revision(),
            "feature": "S1-F2",
            "measurement": "process-tree peak working set only; elapsed-time comparison intentionally omitted",
            "platform": platform.platform(),
            "python_version": platform.python_version(),
            "python_source": str(fixture_module.SOURCE_CORE.parent),
            "rust_binary": str(binary),
            "rust_binary_sha256": digest(binary),
            "fixture": {
                "kind": "same fixture generator used by compare_s1_region_summary.py",
                "source_workbooks": len(fixture["source_files"]),
                "source_sheets": 3,
                "source_data_rows": 5,
                "output_sheets": 2,
                "output_data_rows": 9,
                "template_bytes": fixture["sizes"][str(fixture["template"])],
                "source_bytes": {
                    str(path.name): fixture["sizes"][str(path)] for path in fixture["source_files"]
                },
                "sha256_before": fixture["hashes"],
                "sha256_after": hashes_after,
            },
            "method": {
                "rounds": rounds,
                "round_order": "alternates Python-first and Rust-first",
                "sample_interval_ms": 5,
                "python_engine": "Microsoft Excel COM; existing baseline uses ExcelSession DispatchEx and quits its owned instance",
                "rust_engine": "Calamine + Rust writer; fresh Release CLI process",
                "peak_metric": "maximum sampled RSS sum for worker process tree plus newly launched EXCEL.EXE processes; preexisting Excel process IDs excluded",
                "measurement_limit": "sampling can miss sub-interval peaks; includes process startup and Excel activation; excludes fixture creation and output comparison",
                "preexisting_excel_process_ids_excluded_from_sampling": True,
            },
            "summary": {
                "python_peak_working_set_median_mb": round(statistics.median(python_peaks) / (1024 * 1024), 3),
                "python_peak_working_set_max_mb": round(max(python_peaks) / (1024 * 1024), 3),
                "rust_peak_working_set_median_mb": round(statistics.median(rust_peaks) / (1024 * 1024), 3),
                "rust_peak_working_set_max_mb": round(max(rust_peaks) / (1024 * 1024), 3),
                "python_to_rust_peak_median_ratio": round(statistics.median(python_peaks) / statistics.median(rust_peaks), 3),
                "all_round_output_comparisons_passed": True,
                "all_input_hashes_unchanged": True,
                "all_new_excel_processes_exited": all(
                    not row["python_process_tree"]["new_excel_process_ids_remaining_after_worker"]
                    for row in rounds_data
                ),
            },
            "rounds": rounds_data,
        }

    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return report


def main() -> int:
    if len(sys.argv) > 1 and sys.argv[1] == "--worker":
        if len(sys.argv) != 6 or sys.argv[2] != "python":
            raise SystemExit("usage: --worker python INPUT_DIR OUTPUT_DIR TEMPLATE")
        return python_worker(Path(sys.argv[3]), Path(sys.argv[4]), Path(sys.argv[5]))

    parser = argparse.ArgumentParser()
    parser.add_argument("--audit-binary", type=Path, default=DEFAULT_BINARY)
    parser.add_argument("--report", type=Path, default=DEFAULT_REPORT)
    parser.add_argument("--rounds", type=int, default=3)
    args = parser.parse_args()
    if args.rounds < 1:
        parser.error("--rounds must be positive")
    report = run_memory_profile(args.audit_binary, args.report, args.rounds)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
