#!/usr/bin/env python3
"""Profile S2-F1 with the external-rule manual-acceptance fixture."""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import platform
import shutil
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
DEFAULT_FIXTURE = AUDIT_ROOT / "manual_acceptance" / "20261027" / "s2_f01_external_rule_cases"
DEFAULT_PYTHON_SOURCE = Path(r"C:\Users\AZI\Desktop\VibeCode\3.open-data-audit")
CURRENT_NAME = "9144#2026-06-30#x#20202#甲银行.xlsx"
PREVIOUS_NAME = "9144#2026-03-31#x#20202#甲银行.xlsx"
CONFIG_NAME = "s2_config_external_rule.xlsx"


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def import_helper(name: str, path: Path):
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load helper: {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


def copy_fixture(fixture: Path, destination: Path, external_name: str) -> dict[str, Path]:
    input_root = destination / "input"
    (input_root / "current").mkdir(parents=True)
    (input_root / "previous").mkdir(parents=True)
    shutil.copy2(fixture / "input" / "current" / CURRENT_NAME, input_root / "current" / CURRENT_NAME)
    shutil.copy2(fixture / "input" / "previous" / PREVIOUS_NAME, input_root / "previous" / PREVIOUS_NAME)
    config_path = input_root / CONFIG_NAME
    external_path = input_root / external_name
    shutil.copy2(fixture / "config" / CONFIG_NAME, config_path)
    shutil.copy2(fixture / external_name, external_path)
    return {
        "input_root": input_root,
        "current_dir": input_root / "current",
        "previous_dir": input_root / "previous",
        "config": config_path,
        "external": external_path,
    }


def input_hashes(paths: dict[str, Path]) -> dict[str, str]:
    files = [
        paths["current_dir"] / CURRENT_NAME,
        paths["previous_dir"] / PREVIOUS_NAME,
        paths["config"],
        paths["external"],
    ]
    return {path.relative_to(paths["input_root"]).as_posix(): sha256(path) for path in files}


def plain(value: Any) -> Any:
    if hasattr(value, "__dataclass_fields__"):
        return {name: plain(getattr(value, name)) for name in value.__dataclass_fields__}
    if isinstance(value, dict):
        return {str(key): plain(item) for key, item in value.items()}
    if isinstance(value, (tuple, list)):
        return [plain(item) for item in value]
    if isinstance(value, Path):
        return str(value)
    if hasattr(value, "value") and type(value).__module__ == "enum":
        return value.value
    return value


def python_service(source_root: Path, paths: dict[str, Path], output: Path) -> dict[str, Any]:
    sys.path.insert(0, str(source_root / "core" / "src"))
    from base_audit.systems.s2_report_collection.service import PeriodAuditV2Service

    events: list[dict[str, Any]] = []
    stage_started: dict[str, float] = {}
    stage_timings: dict[str, float] = {}
    active_stage: str | None = None
    started = time.perf_counter()

    def on_progress(event: dict[str, Any]) -> None:
        nonlocal active_stage
        now = time.perf_counter()
        stage = str(event.get("stage", "unknown"))
        if active_stage != stage:
            if active_stage is not None:
                stage_timings[active_stage] = round((now - stage_started[active_stage]) * 1000, 2)
            active_stage = stage
            stage_started[stage] = now
        events.append({**event, "_elapsed_ms": round((now - started) * 1000, 2)})

    result = PeriodAuditV2Service().run(
        current_dir=paths["current_dir"],
        previous_dir=paths["previous_dir"],
        config_path=paths["config"],
        output_dir=output,
        central_path=paths["external"],
        external_reader="AUTO",
        on_progress=on_progress,
    )
    ended = time.perf_counter()
    if active_stage is not None:
        stage_timings[active_stage] = round((ended - stage_started[active_stage]) * 1000, 2)
    return {
        "findings": plain(result.findings),
        "summary": plain(result.summary),
        "workbook": Path(result.paths.result),
        "events": events,
        "stage_timings_ms": stage_timings,
        "service_elapsed_ms": round((ended - started) * 1000, 2),
    }


def rust_command(binary: Path, paths: dict[str, Path], output: Path) -> list[str]:
    return [
        str(binary),
        "s2",
        "--current-dir",
        str(paths["current_dir"]),
        "--previous-dir",
        str(paths["previous_dir"]),
        "--config",
        str(paths["config"]),
        "--external",
        str(paths["external"]),
        "--output-dir",
        str(output),
    ]


def parse_rust_process(stdout: str, stderr: str, wall_ms: float) -> dict[str, Any]:
    result = json.loads(stdout)
    events = []
    for line in stderr.splitlines():
        if line.strip():
            item = json.loads(line)
            if item.get("type") == "progress":
                events.append(item)
    first_by_stage: dict[str, float] = {}
    for item in events:
        event = item.get("event", {})
        first_by_stage.setdefault(str(event.get("stage", "unknown")), float(item.get("elapsed_ms", 0)))
    stage_timings: dict[str, float] = {}
    ordered = list(first_by_stage.items())
    for index, (stage, start_ms) in enumerate(ordered):
        if index + 1 < len(ordered):
            stage_timings[stage] = round(ordered[index + 1][1] - start_ms, 2)
        elif stage != "complete":
            service_ms = max((float(item.get("elapsed_ms", 0)) for item in events), default=0.0)
            stage_timings[stage] = round(service_ms - start_ms, 2)
        else:
            stage_timings[stage] = 0.0
    service_elapsed_ms = max((float(item.get("elapsed_ms", 0)) for item in events), default=0.0)
    result.update(
        {
            "events": events,
            "stage_timings_ms": stage_timings,
            "service_elapsed_ms": service_elapsed_ms,
            "wall_ms_including_cli_startup": round(wall_ms, 2),
        }
    )
    return result


def run_rust(binary: Path, paths: dict[str, Path], output: Path) -> dict[str, Any]:
    started = time.perf_counter()
    completed = subprocess.run(
        rust_command(binary, paths, output),
        check=True,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    wall_ms = (time.perf_counter() - started) * 1000
    return parse_rust_process(completed.stdout, completed.stderr, wall_ms)


def run_monitored(command: list[str], cwd: Path | None = None) -> tuple[str, str, int, int]:
    try:
        import psutil
    except ImportError as error:  # pragma: no cover - workstation dependency
        raise RuntimeError("psutil is required for process peak-working-set measurement") from error

    # Redirect to temporary files so progress output cannot fill an unread pipe
    # while the parent samples memory.
    with tempfile.TemporaryFile() as stdout_file, tempfile.TemporaryFile() as stderr_file:
        process = subprocess.Popen(command, cwd=cwd, stdout=stdout_file, stderr=stderr_file)
        try:
            monitored = psutil.Process(process.pid)
        except psutil.Error:
            monitored = None
        peak_bytes = 0
        while process.poll() is None:
            if monitored is not None:
                try:
                    info = monitored.memory_info()
                    peak_bytes = max(peak_bytes, int(getattr(info, "peak_wset", info.rss)))
                except psutil.Error:
                    pass
            time.sleep(0.005)
        process.wait()
        stdout_file.seek(0)
        stderr_file.seek(0)
        stdout = stdout_file.read().decode("utf-8", errors="replace")
        stderr = stderr_file.read().decode("utf-8", errors="replace")
    return stdout, stderr, process.returncode or 0, peak_bytes


def stage_median(reports: list[dict[str, Any]]) -> dict[str, float]:
    stages = sorted({stage for report in reports for stage in report["stage_timings_ms"]})
    return {
        stage: round(statistics.median(report["stage_timings_ms"].get(stage, 0.0) for report in reports), 2)
        for stage in stages
    }


def normalize_workbook_path_value(value: str, helper: Any) -> str:
    # Rust's Windows path serializer may add the extended-length `\\?\` prefix.
    return helper.normalize_path_text(value.replace("\\\\?\\", ""))


def workbook_info(path: Path) -> dict[str, Any]:
    from openpyxl import load_workbook

    workbook = load_workbook(path, read_only=True, data_only=False)
    try:
        sheets = []
        for sheet in workbook.worksheets:
            sheets.append(
                {
                    "name": sheet.title,
                    "rows": sheet.max_row,
                    "columns": sheet.max_column,
                    "data_rows_excluding_header": max(0, sheet.max_row - 1),
                }
            )
        return {"bytes": path.stat().st_size, "sheets": sheets}
    finally:
        workbook.close()


def check_external_value(workbook_path: Path, expected_external: float, expected_difference: float) -> dict[str, Any]:
    from openpyxl import load_workbook

    workbook = load_workbook(workbook_path, read_only=True, data_only=True)
    try:
        sheet = workbook["外部核对规则"]
        if sheet.max_row != 2:
            raise AssertionError(f"expected one external-rule data row; got {sheet.max_row - 1}")
        values = [sheet.cell(2, column).value for column in range(1, 17)]
        if values[7] != "20202001":
            raise AssertionError(f"unexpected external-rule indicator: {values[7]!r}")
        if not isinstance(values[12], (int, float)) or abs(float(values[12]) - 1234567.89) > 1e-9:
            raise AssertionError(f"unexpected report value: {values[12]!r}")
        if not isinstance(values[13], (int, float)) or abs(float(values[13]) - expected_external) > 1e-9:
            raise AssertionError(f"unexpected external value: {values[13]!r}")
        if not isinstance(values[14], (int, float)) or abs(float(values[14]) - expected_difference) > 1e-9:
            raise AssertionError(f"unexpected external difference: {values[14]!r}")
        expected_level = None if expected_difference == 0 else "硬性"
        if values[10] != expected_level:
            raise AssertionError(f"expected external-rule level {expected_level!r}, got {values[10]!r}")
        return {
            "indicator_code": values[7],
            "report_value": values[12],
            "external_value": values[13],
            "absolute_difference": values[14],
            "level": values[10],
        }
    finally:
        workbook.close()


def python_memory_worker(source_root: Path, paths: dict[str, Path], output: Path) -> int:
    run = python_service(source_root, paths, output)
    print(json.dumps({"status": "completed", "summary": run["summary"]}, ensure_ascii=True, default=str))
    return 0


def memory_measure(binary: Path, source_root: Path, paths: dict[str, Path], root: Path) -> dict[str, Any]:
    py_output = root / "memory-python-output"
    py_command = [
        sys.executable,
        str(Path(__file__).resolve()),
        "--memory-worker",
        "python",
        "--source-root",
        str(source_root),
        "--input-root",
        str(paths["input_root"]),
        "--external-name",
        paths["external"].name,
        "--output-dir",
        str(py_output),
    ]
    _, _, py_return, py_peak = run_monitored(py_command, cwd=WORKSPACE)
    if py_return != 0:
        raise RuntimeError("Python memory-sample worker failed")

    rust_output = root / "memory-rust-output"
    rust_stdout, rust_stderr, rust_return, rust_peak = run_monitored(rust_command(binary, paths, rust_output))
    if rust_return != 0:
        raise RuntimeError(f"Rust memory-sample run failed: {rust_stderr[-4000:]}")
    rust_result = json.loads(rust_stdout)
    return {
        "measurement": "OS peak working set high-water mark, sampled during a separate one-run process; process startup included",
        "sampler_interval_ms": 5,
        "python_process_peak_working_set_bytes": py_peak,
        "rust_cli_process_peak_working_set_bytes": rust_peak,
        "rust_cli_status": rust_result.get("status"),
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--audit-binary", type=Path)
    parser.add_argument("--source-root", type=Path, default=DEFAULT_PYTHON_SOURCE)
    parser.add_argument("--fixture-root", type=Path, default=DEFAULT_FIXTURE)
    parser.add_argument("--report", type=Path, default=AUDIT_ROOT / "probe" / "reports" / "s2f1_external_profile_20261030.json")
    parser.add_argument("--rounds", type=int, default=3)
    parser.add_argument("--memory-worker", choices=("python",))
    parser.add_argument("--input-root", type=Path)
    parser.add_argument("--external-name")
    parser.add_argument("--output-dir", type=Path)
    args = parser.parse_args()

    if args.memory_worker:
        if not all((args.input_root, args.external_name, args.output_dir)):
            parser.error("memory worker requires --input-root, --external-name and --output-dir")
        paths = {
            "input_root": args.input_root.resolve(),
            "current_dir": args.input_root.resolve() / "current",
            "previous_dir": args.input_root.resolve() / "previous",
            "config": args.input_root.resolve() / CONFIG_NAME,
            "external": args.input_root.resolve() / args.external_name,
        }
        return python_memory_worker(args.source_root.resolve(), paths, args.output_dir.resolve())

    if args.audit_binary is None or not args.audit_binary.is_file():
        parser.error("--audit-binary must point to an existing Release executable")
    if not args.source_root.is_dir() or not args.fixture_root.is_dir():
        parser.error("the Python source and manual-acceptance fixture directories must exist")
    if args.rounds < 1:
        parser.error("--rounds must be positive")

    helper = import_helper("s2_service_parity_helpers", AUDIT_ROOT / "tests" / "compare" / "run_s2_service_parity.py")
    workbook_compare = helper.load_workbook_comparator(AUDIT_ROOT)
    compare_tree = helper.compare_tree
    source_hashes_before = {
        relative: sha256(args.source_root / relative)
        for relative in (
            "core/src/base_audit/systems/s2_report_collection/service.py",
            "core/src/base_audit/systems/s2_report_collection/external_engine.py",
            "core/src/base_audit/systems/s2_report_collection/exporter.py",
        )
    }
    report: dict[str, Any] = {
        "schema_version": 1,
        "date": datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds"),
        "feature": "S2-F1",
        "scenario": "external-rule-manual-acceptance-normal-and-exception",
        "platform": platform.platform(),
        "python_version": platform.python_version(),
        "python_source": str(args.source_root.resolve()),
        "rust_binary": str(args.audit_binary.resolve()),
        "rust_binary_sha256": sha256(args.audit_binary),
        "fixture_root": str(args.fixture_root.resolve()),
        "rounds": args.rounds,
        "status": "failed",
        "scenarios": {},
    }

    with tempfile.TemporaryDirectory(prefix="audit-s2f1-external-profile-") as temporary:
        root = Path(temporary)
        for scenario, external_name, expected_external, expected_difference in (
            ("normal", "external_normal.xlsx", 1234567.89, 0.0),
            ("exception", "external_exception.xlsx", 1234.5, 1233333.39),
        ):
            case_root = root / scenario
            case_root.mkdir()
            paths = copy_fixture(args.fixture_root, case_root, external_name)
            hashes_before = input_hashes(paths)
            fixture_files = [
                paths["current_dir"] / CURRENT_NAME,
                paths["previous_dir"] / PREVIOUS_NAME,
                paths["config"],
                paths["external"],
            ]
            file_info = {}
            for path in fixture_files:
                info = workbook_info(path)
                file_info[path.relative_to(paths["input_root"]).as_posix()] = info

            # Exclude one warm-up run from timing summaries, but retain it as a correctness check.
            warm_python = python_service(args.source_root.resolve(), paths, case_root / "warm-python")
            warm_rust = run_rust(args.audit_binary.resolve(), paths, case_root / "warm-rust")
            warm_py_result = check_external_value(warm_python["workbook"], expected_external, expected_difference)
            warm_rs_result = check_external_value(Path(warm_rust["report_path"]), expected_external, expected_difference)
            if warm_py_result != warm_rs_result:
                raise AssertionError(f"warm-up external result mismatch for {scenario}")

            python_runs: list[dict[str, Any]] = []
            rust_runs: list[dict[str, Any]] = []
            parity_rounds: list[dict[str, Any]] = []
            for index in range(args.rounds):
                py_dir = case_root / f"python-{index + 1}"
                rs_dir = case_root / f"rust-{index + 1}"
                if index % 2 == 0:
                    python_run = python_service(args.source_root.resolve(), paths, py_dir)
                    rust_run = run_rust(args.audit_binary.resolve(), paths, rs_dir)
                else:
                    rust_run = run_rust(args.audit_binary.resolve(), paths, rs_dir)
                    python_run = python_service(args.source_root.resolve(), paths, py_dir)
                python_runs.append(python_run)
                rust_runs.append(rust_run)

                business_differences: list[dict[str, Any]] = []
                compare_tree(python_run["summary"], rust_run.get("summary"), "summary", business_differences)
                workbook_stats, raw_workbook_differences = workbook_compare(
                    python_run["workbook"], Path(rust_run["report_path"]), 1e-15
                )
                path_representation_differences = [
                    difference
                    for difference in raw_workbook_differences
                    if difference.get("kind") == "value"
                    and isinstance(difference.get("python"), str)
                    and isinstance(difference.get("rust"), str)
                    and normalize_workbook_path_value(difference["python"], helper)
                    == normalize_workbook_path_value(difference["rust"], helper)
                ]
                workbook_differences = [
                    difference for difference in raw_workbook_differences
                    if difference not in path_representation_differences
                ]
                expected_result = check_external_value(
                    python_run["workbook"], expected_external, expected_difference
                )
                actual_result = check_external_value(
                    Path(rust_run["report_path"]), expected_external, expected_difference
                )
                if business_differences or workbook_differences or expected_result != actual_result:
                    raise AssertionError(
                        f"{scenario} round {index + 1} mismatch: summary={business_differences[:10]!r}; "
                        f"workbook={workbook_differences[:10]!r}; expected={expected_result!r}; actual={actual_result!r}"
                    )
                parity_rounds.append(
                    {
                        "round": index + 1,
                        "summary_differences": len(business_differences),
                        "workbook_differences": len(workbook_differences),
                        "windows_extended_path_prefix_differences": len(path_representation_differences),
                        "workbook_stats": workbook_stats,
                        "expected_rule_result": actual_result,
                        "python_output_bytes": Path(python_run["workbook"]).stat().st_size,
                        "rust_output_bytes": Path(rust_run["report_path"]).stat().st_size,
                    }
                )

            hashes_after = input_hashes(paths)
            if hashes_after != hashes_before:
                raise AssertionError(f"{scenario}: one of the input files was modified")
            memory = memory_measure(args.audit_binary.resolve(), args.source_root.resolve(), paths, case_root)

            python_times = [run["service_elapsed_ms"] for run in python_runs]
            rust_wall_times = [run["wall_ms_including_cli_startup"] for run in rust_runs]
            rust_internal_times = [run["service_elapsed_ms"] for run in rust_runs]
            python_stage = stage_median(python_runs)
            rust_stage = stage_median(rust_runs)
            report["scenarios"][scenario] = {
                "external_file": external_name,
                "inputs": file_info,
                "input_sha256_before": hashes_before,
                "input_sha256_after": hashes_after,
                "python_service_ms": {
                    "rounds": python_times,
                    "median": round(statistics.median(python_times), 2),
                    "stage_median_ms": python_stage,
                },
                "rust_cli_ms": {
                    "wall_rounds_including_cli_startup": rust_wall_times,
                    "wall_median_including_cli_startup": round(statistics.median(rust_wall_times), 2),
                    "service_rounds_from_cli_progress": rust_internal_times,
                    "service_median_from_cli_progress": round(statistics.median(rust_internal_times), 2),
                    "stage_median_ms": rust_stage,
                },
                "python_over_rust_wall": round(statistics.median(python_times) / statistics.median(rust_wall_times), 3),
                "python_over_rust_internal": round(statistics.median(python_times) / statistics.median(rust_internal_times), 3),
                "correctness": {
                    "warmup_expected_rule_result": warm_rs_result,
                    "parity_rounds": parity_rounds,
                    "all_rounds_zero_differences": True,
                    "input_hashes_unchanged": True,
                },
                "process_peak_memory": memory,
            }

    source_hashes_after = {
        relative: sha256(args.source_root / relative)
        for relative in source_hashes_before
    }
    if source_hashes_after != source_hashes_before:
        raise AssertionError("read-only Python source hashes changed during the profile")
    report["python_source_sha256_before"] = source_hashes_before
    report["python_source_sha256_after"] = source_hashes_after
    report["status"] = "passed"
    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"status": report["status"], "report": str(args.report), "scenarios": report["scenarios"]}, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
