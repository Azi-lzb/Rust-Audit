#!/usr/bin/env python3
"""Compare Python and Rust S2-F1 on a synthetic 21+21 workbook batch."""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import shutil
import statistics
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


WORKSPACE = Path(__file__).resolve().parents[3]
AUDIT_ROOT = WORKSPACE / "audit-core"
DEFAULT_SOURCE = Path(r"C:\Users\AZI\Desktop\VibeCode\3.open-data-audit")
DEFAULT_FIXTURE = (
    AUDIT_ROOT / "manual_acceptance" / "20261027" / "s2_f01_external_rule_cases"
)
CONFIG_PROFILE = AUDIT_ROOT / "tests" / "compare" / "perf_s2f1_config_read_reuse_20261030.py"
EXTERNAL_PROFILE = AUDIT_ROOT / "tests" / "compare" / "perf_s2f1_external_20261030.py"
INSTITUTIONS = [
    ("9144", "甲银行"),
    ("9145", "合成机构01"),
    ("9146", "合成机构02"),
    ("9147", "合成机构03"),
    ("9148", "合成机构04"),
    ("9149", "合成机构05"),
    ("9150", "合成机构06"),
]
FORMS = ("20202", "20203", "20204")


def import_helper(name: str, path: Path):
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load helper: {path}")
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


def input_hashes(paths: dict[str, Path]) -> dict[str, str]:
    result = {}
    for key in ("current_dir", "previous_dir"):
        for path in sorted(paths[key].glob("*.xlsx")):
            result[path.relative_to(paths["input_root"]).as_posix()] = sha256(path)
    for key in ("config", "external"):
        result[paths[key].relative_to(paths["input_root"]).as_posix()] = sha256(
            paths[key]
        )
    return result


def make_config(config_path: Path, config_helper: Any) -> dict[str, int]:
    from openpyxl import load_workbook

    counts = config_helper.scale_config(config_path, 728, len(INSTITUTIONS))
    workbook = load_workbook(config_path)
    try:
        sheet = workbook["机构参照"]
        for row_number, (institution_id, institution_name) in enumerate(
            INSTITUTIONS, start=2
        ):
            sheet.cell(row_number, 1).value = institution_name
            sheet.cell(row_number, 2).value = institution_id
        workbook.save(config_path)
    finally:
        workbook.close()
    return counts


def make_report(path: Path, form: str, rows: list[tuple[str, str, float]]) -> int:
    from openpyxl import Workbook

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = form
    sheet.append(["合成报表"])
    sheet.append([None])
    sheet.append(["指标代码", "指标名称", "本期数值"])
    for code, name, value in rows:
        sheet.append([code, name, value])
    workbook.save(path)
    workbook.close()
    return len(rows)


def make_batch(input_root: Path, config_helper: Any) -> dict[str, Any]:
    from openpyxl import load_workbook

    current_dir = input_root / "current"
    previous_dir = input_root / "previous"
    current_dir.mkdir(parents=True)
    previous_dir.mkdir(parents=True)
    config_path = input_root / "s2_config_external_rule.xlsx"
    external_path = input_root / "external_normal.xlsx"
    config_counts = make_config(config_path, config_helper)

    config = load_workbook(config_path, read_only=True, data_only=True)
    try:
        indicator_sheet = config["指标参照"]
        configured = {}
        for row in indicator_sheet.iter_rows(min_row=2, values_only=True):
            if not row or row[0] in (None, ""):
                continue
            configured[str(row[0])] = (str(row[1]), str(row[2]))
    finally:
        config.close()

    by_form = {form: [] for form in FORMS}
    for code, (name, form) in configured.items():
        if code.isdigit() and int(code) >= 90_000_000:
            by_form.setdefault(form, []).append((code, name))
    for values in by_form.values():
        values.sort(key=lambda item: int(item[0]))

    workbook_rows = 0
    for institution_index, (institution_id, institution_name) in enumerate(
        INSTITUTIONS
    ):
        for form_index, form in enumerate(FORMS):
            selected = list(by_form[form])
            if institution_index == 0 and form == "20202":
                selected.insert(0, ("20202001", "贷款余额"))
            for period_name, date, previous in (
                ("current", "2026-06-30", False),
                ("previous", "2026-03-31", True),
            ):
                rows = []
                for indicator_index, (code, name) in enumerate(selected):
                    base = 1_000_000.0 + institution_index * 100_000.0
                    base += int(code) % 90_000_000 * 0.25
                    if code == "20202001":
                        value = 1_234_567.89
                    else:
                        value = base
                        if not previous and int(code) % 11 == 0:
                            value *= 1.02
                    rows.append((code, name, value))
                filename = (
                    f"{institution_id}#{date}#x#{form}#{institution_name}.xlsx"
                )
                path = (current_dir if period_name == "current" else previous_dir) / filename
                row_count = make_report(path, form, rows)
                if period_name == "current":
                    workbook_rows += row_count

    if len(list(current_dir.glob("*.xlsx"))) != 21 or len(
        list(previous_dir.glob("*.xlsx"))
    ) != 21:
        raise AssertionError("synthetic period directories must contain 21 workbooks each")

    sample = next(iter(sorted(current_dir.glob("*.xlsx"))))
    sample_workbook = load_workbook(sample, read_only=True, data_only=True)
    try:
        sample_sheet = sample_workbook.active
        sample_info = {
            "bytes": sample.stat().st_size,
            "sheet_count": len(sample_workbook.worksheets),
            "sheet_name": sample_sheet.title,
            "rows": sample_sheet.max_row,
            "columns": sample_sheet.max_column,
        }
    finally:
        sample_workbook.close()
    return {
        "config": config_counts,
        "institution_count": len(INSTITUTIONS),
        "forms": list(FORMS),
        "current_workbook_count": 21,
        "previous_workbook_count": 21,
        "current_import_record_count": workbook_rows,
        "input_workbook_sample": sample_info,
        "input_file_bytes": sum(
            path.stat().st_size
            for directory in (current_dir, previous_dir)
            for path in directory.glob("*.xlsx")
        ),
    }


def normalize_differences(differences: list[dict[str, Any]], parity: Any):
    return [
        difference
        for difference in differences
        if not (
            difference.get("kind") == "value"
            and isinstance(difference.get("python"), str)
            and isinstance(difference.get("rust"), str)
            and parity.normalize_path_text(difference["python"].replace("\\\\?\\", ""))
            == parity.normalize_path_text(difference["rust"].replace("\\\\?\\", ""))
        )
    ]


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--audit-binary", type=Path, required=True)
    parser.add_argument("--source-root", type=Path, default=DEFAULT_SOURCE)
    parser.add_argument("--fixture-root", type=Path, default=DEFAULT_FIXTURE)
    parser.add_argument("--report", type=Path)
    parser.add_argument("--rounds", type=int, default=3)
    parser.add_argument("--memory", action="store_true")
    args = parser.parse_args()
    if args.rounds < 1:
        parser.error("rounds must be positive")
    if not args.audit_binary.is_file() or not args.source_root.is_dir() or not args.fixture_root.is_dir():
        parser.error("the Rust executable, Python source, and manual-acceptance fixture must exist")

    config_helper = import_helper("s2f1_config_read_reuse", CONFIG_PROFILE)
    profile = import_helper("s2f1_external_profile", EXTERNAL_PROFILE)
    parity = import_helper(
        "s2_service_parity_helpers",
        AUDIT_ROOT / "tests" / "compare" / "run_s2_service_parity.py",
    )
    compare_workbooks = parity.load_workbook_comparator(AUDIT_ROOT)
    source_files = (
        "core/src/base_audit/systems/s2_report_collection/service.py",
        "core/src/base_audit/systems/s2_report_collection/config_check.py",
        "core/src/base_audit/systems/s2_report_collection/config.py",
        "core/src/base_audit/systems/s2_report_collection/importer.py",
        "core/src/base_audit/systems/s2_report_collection/period_engine.py",
        "core/src/base_audit/systems/s2_report_collection/exporter.py",
    )
    source_hashes_before = {
        relative: sha256(args.source_root / relative) for relative in source_files
    }
    report: dict[str, Any] = {
        "schema_version": 1,
        "date": datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds"),
        "feature": "S2-F1",
        "scenario": "synthetic-21-plus-21-batch",
        "scope": (
            "Synthetic business-shaped input generated from the local S2 manual-acceptance "
            "fixture schema. It does not use or represent the unavailable historical real batch."
        ),
        "platform": profile.platform.platform(),
        "python_version": profile.platform.python_version(),
        "source_root": str(args.source_root.resolve()),
        "rust_binary": str(args.audit_binary.resolve()),
        "rust_binary_sha256": sha256(args.audit_binary),
        "fixture_root": str(args.fixture_root.resolve()),
        "rounds": args.rounds,
        "status": "failed",
    }

    with tempfile.TemporaryDirectory(prefix="audit-s2f1-21x21-scale-") as temporary:
        root = Path(temporary)
        input_root = root / "input"
        input_root.mkdir()
        shutil.copy2(
            args.fixture_root / "config" / "s2_config_external_rule.xlsx",
            input_root / "s2_config_external_rule.xlsx",
        )
        shutil.copy2(
            args.fixture_root / "external_normal.xlsx",
            input_root / "external_normal.xlsx",
        )
        paths = {
            "input_root": input_root,
            "current_dir": input_root / "current",
            "previous_dir": input_root / "previous",
            "config": input_root / "s2_config_external_rule.xlsx",
            "external": input_root / "external_normal.xlsx",
        }
        batch_info = make_batch(input_root, config_helper)
        hashes_before = input_hashes(paths)

        warm_python = profile.python_service(
            args.source_root.resolve(), paths, root / "warm-python"
        )
        warm_rust = profile.run_rust(
            args.audit_binary.resolve(), paths, root / "warm-rust"
        )
        warm_summary_differences: list[dict[str, Any]] = []
        parity.compare_tree(
            warm_python["summary"],
            warm_rust.get("summary"),
            "summary",
            warm_summary_differences,
        )
        _, warm_workbook_differences = compare_workbooks(
            warm_python["workbook"], Path(warm_rust["report_path"]), 1e-15
        )
        warm_workbook_differences = normalize_differences(
            warm_workbook_differences, parity
        )
        if warm_summary_differences or warm_workbook_differences:
            raise AssertionError(
                "warm-up output mismatch: "
                f"summary={warm_summary_differences[:10]!r}; "
                f"workbook={warm_workbook_differences[:10]!r}"
            )

        python_runs: list[dict[str, Any]] = []
        rust_runs: list[dict[str, Any]] = []
        parity_rounds: list[dict[str, Any]] = []
        for index in range(args.rounds):
            py_dir = root / f"python-{index + 1}"
            rs_dir = root / f"rust-{index + 1}"
            if index % 2 == 0:
                python_run = profile.python_service(
                    args.source_root.resolve(), paths, py_dir
                )
                rust_run = profile.run_rust(
                    args.audit_binary.resolve(), paths, rs_dir
                )
            else:
                rust_run = profile.run_rust(
                    args.audit_binary.resolve(), paths, rs_dir
                )
                python_run = profile.python_service(
                    args.source_root.resolve(), paths, py_dir
                )
            python_runs.append(python_run)
            rust_runs.append(rust_run)
            summary_differences: list[dict[str, Any]] = []
            parity.compare_tree(
                python_run["summary"],
                rust_run.get("summary"),
                "summary",
                summary_differences,
            )
            workbook_stats, workbook_differences = compare_workbooks(
                python_run["workbook"], Path(rust_run["report_path"]), 1e-15
            )
            workbook_differences = normalize_differences(workbook_differences, parity)
            if summary_differences or workbook_differences:
                raise AssertionError(
                    f"round {index + 1} mismatch: "
                    f"summary={summary_differences[:10]!r}; "
                    f"workbook={workbook_differences[:10]!r}"
                )
            parity_rounds.append(
                {
                    "round": index + 1,
                    "summary_difference_count": len(summary_differences),
                    "workbook_difference_count": len(workbook_differences),
                    "workbook_stats": workbook_stats,
                    "findings": len(python_run["findings"]),
                    "python_output_bytes": Path(python_run["workbook"]).stat().st_size,
                    "rust_output_bytes": Path(rust_run["report_path"]).stat().st_size,
                }
            )

        hashes_after = input_hashes(paths)
        python_service = [run["service_elapsed_ms"] for run in python_runs]
        rust_service = [run["service_elapsed_ms"] for run in rust_runs]
        rust_wall = [run["wall_ms_including_cli_startup"] for run in rust_runs]
        report.update(
            {
                "input_profile": batch_info,
                "input_sha256_before": hashes_before,
                "input_sha256_after": hashes_after,
                "inputs_unchanged": hashes_before == hashes_after,
                "python_service_ms": {
                    "rounds": [round(value, 2) for value in python_service],
                    "median": round(statistics.median(python_service), 2),
                    "stage_medians": profile.stage_median(python_runs),
                },
                "rust_service_ms": {
                    "rounds": [round(value, 2) for value in rust_service],
                    "median": round(statistics.median(rust_service), 2),
                    "wall_rounds_including_cli_startup": [
                        round(value, 2) for value in rust_wall
                    ],
                    "wall_median_including_cli_startup": round(
                        statistics.median(rust_wall), 2
                    ),
                    "stage_medians": profile.stage_median(rust_runs),
                },
                "parity_rounds": parity_rounds,
                "source_sha256_before": source_hashes_before,
            }
        )
        if args.memory:
            report["process_peak_memory"] = profile.memory_measure(
                args.audit_binary.resolve(), args.source_root.resolve(), paths, root
            )
        report["source_sha256_after"] = {
            relative: sha256(args.source_root / relative) for relative in source_files
        }
        report["source_unchanged"] = (
            report["source_sha256_before"] == report["source_sha256_after"]
        )
        report["status"] = (
            "passed"
            if report["inputs_unchanged"]
            and report["source_unchanged"]
            and len(parity_rounds) == args.rounds
            else "failed"
        )

    output = args.report or (
        AUDIT_ROOT / "probe" / "reports" / "s2f1_batch_scale_20261030.json"
    )
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(
        json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    print(
        json.dumps(
            {
                key: report[key]
                for key in (
                    "status",
                    "input_profile",
                    "python_service_ms",
                    "rust_service_ms",
                    "parity_rounds",
                    "process_peak_memory",
                    "inputs_unchanged",
                    "source_unchanged",
                )
                if key in report
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    print(f"\nSaved report: {output}")
    return 0 if report["status"] == "passed" else 1


if __name__ == "__main__":
    raise SystemExit(main())
