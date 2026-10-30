#!/usr/bin/env python3
"""Profile S2-F1 config validation/loading with a production-sized synthetic config."""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
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
PROFILE_PATH = AUDIT_ROOT / "tests" / "compare" / "perf_s2f1_external_20261030.py"


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


def scale_config(config_path: Path, indicator_count: int, institution_count: int) -> dict[str, int]:
    from openpyxl import load_workbook

    workbook = load_workbook(config_path)
    try:
        indicator_sheet = workbook["指标参照"]
        current_indicators = [
            row
            for row in range(2, indicator_sheet.max_row + 1)
            if indicator_sheet.cell(row, 1).value not in (None, "")
        ]
        if not current_indicators:
            raise RuntimeError("the source config has no indicator row to use as a schema")
        template = [
            indicator_sheet.cell(current_indicators[-1], column).value
            for column in range(1, indicator_sheet.max_column + 1)
        ]
        next_row = indicator_sheet.max_row + 1
        for index in range(max(0, indicator_count - len(current_indicators))):
            row = list(template)
            row[0] = str(90_000_000 + index)
            row[1] = f"合成指标{index + 1:04d}"
            row[2] = ("20202", "20203", "20204")[index % 3]
            row[3] = "余额"
            row[4] = "是"
            row[5] = "G1"
            row[6] = 0
            row[7] = "否"
            row[8] = "性能基准生成"
            for column, value in enumerate(row, start=1):
                indicator_sheet.cell(next_row + index, column).value = value

        institution_sheet = workbook["机构参照"]
        current_institutions = [
            row
            for row in range(2, institution_sheet.max_row + 1)
            if institution_sheet.cell(row, 1).value not in (None, "")
        ]
        institution_template = [
            institution_sheet.cell(current_institutions[-1], column).value
            for column in range(1, institution_sheet.max_column + 1)
        ]
        next_institution_row = institution_sheet.max_row + 1
        for index in range(max(0, institution_count - len(current_institutions))):
            row = list(institution_template)
            row[0] = f"合成机构{index + 1:02d}"
            row[1] = str(90_100 + index)
            for column, value in enumerate(row, start=1):
                institution_sheet.cell(next_institution_row + index, column).value = value
        workbook.save(config_path)
        return {
            "indicator_rows": len(current_indicators)
            + max(0, indicator_count - len(current_indicators)),
            "institution_rows": len(current_institutions)
            + max(0, institution_count - len(current_institutions)),
        }
    finally:
        workbook.close()


def normalize_differences(differences: list[dict[str, Any]], helper: Any) -> list[dict[str, Any]]:
    return [
        difference
        for difference in differences
        if not (
            difference.get("kind") == "value"
            and isinstance(difference.get("python"), str)
            and isinstance(difference.get("rust"), str)
            and helper.normalize_path_text(difference["python"].replace("\\\\?\\", ""))
            == helper.normalize_path_text(difference["rust"].replace("\\\\?\\", ""))
        )
    ]


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--audit-binary", type=Path, required=True)
    parser.add_argument("--source-root", type=Path, default=DEFAULT_SOURCE)
    parser.add_argument("--fixture-root", type=Path, default=DEFAULT_FIXTURE)
    parser.add_argument("--report", type=Path)
    parser.add_argument("--rounds", type=int, default=5)
    parser.add_argument("--indicators", type=int, default=728)
    parser.add_argument("--institutions", type=int, default=7)
    args = parser.parse_args()
    if args.rounds < 1 or args.indicators < 2 or args.institutions < 1:
        parser.error("rounds must be positive; indicators must be >= 2; institutions must be positive")
    if not args.audit_binary.is_file() or not args.source_root.is_dir() or not args.fixture_root.is_dir():
        parser.error("the Rust executable, Python source, and manual-acceptance fixture must exist")

    profile = import_helper("s2f1_external_profile", PROFILE_PATH)
    parity = import_helper(
        "s2_service_parity_helpers",
        AUDIT_ROOT / "tests" / "compare" / "run_s2_service_parity.py",
    )
    compare_workbooks = parity.load_workbook_comparator(AUDIT_ROOT)
    source_files = (
        "core/src/base_audit/systems/s2_report_collection/service.py",
        "core/src/base_audit/systems/s2_report_collection/config_check.py",
        "core/src/base_audit/systems/s2_report_collection/config.py",
    )
    source_hashes_before = {
        relative: sha256(args.source_root / relative) for relative in source_files
    }
    report: dict[str, Any] = {
        "schema_version": 1,
        "date": datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds"),
        "feature": "S2-F1",
        "scenario": "synthetic-production-sized-config-read-reuse",
        "scope": (
            "1 current + 1 previous workbook; 728 configured indicators and 7 institution "
            "profiles. This isolates config scaling and is not the historic 21+21 real-data batch."
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

    with tempfile.TemporaryDirectory(prefix="audit-s2f1-config-scale-") as temporary:
        root = Path(temporary)
        paths = profile.copy_fixture(args.fixture_root, root, "external_normal.xlsx")
        synthetic_counts = scale_config(
            paths["config"], args.indicators, args.institutions
        )
        hashes_before = profile.input_hashes(paths)

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
                }
            )

        hashes_after = profile.input_hashes(paths)
        python_service = [
            run["service_elapsed_ms"] for run in python_runs
        ]
        rust_service = [
            run["service_elapsed_ms"] for run in rust_runs
        ]
        rust_wall = [
            run["wall_ms_including_cli_startup"] for run in rust_runs
        ]
        report.update(
            {
                "synthetic_config": synthetic_counts,
                "input_sha256_before": hashes_before,
                "input_sha256_after": hashes_after,
                "inputs_unchanged": hashes_before == hashes_after,
                "python_service_ms": {
                    "rounds": [round(value, 2) for value in python_service],
                    "median": round(statistics.median(python_service), 2),
                    "config_stage_median": round(
                        statistics.median(
                            run["stage_timings_ms"].get("config", 0.0)
                            for run in python_runs
                        ),
                        2,
                    ),
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
                    "config_stage_median": round(
                        statistics.median(
                            run["stage_timings_ms"].get("config", 0.0)
                            for run in rust_runs
                        ),
                        2,
                    ),
                    "stage_medians": profile.stage_median(rust_runs),
                },
                "parity_rounds": parity_rounds,
                "source_sha256_before": source_hashes_before,
            }
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
        AUDIT_ROOT
        / "probe"
        / "reports"
        / "s2f1_config_read_reuse_20261030_before.json"
    )
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(
        json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    print(json.dumps(report, ensure_ascii=False, indent=2))
    print(f"\nSaved report: {output}")
    return 0 if report["status"] == "passed" else 1


if __name__ == "__main__":
    raise SystemExit(main())
