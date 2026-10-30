#!/usr/bin/env python3
"""Profile S3-F3 execution with conflicting feedback and compare full outputs."""

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
from pathlib import Path
from typing import Any

from openpyxl import load_workbook


WORKSPACE = Path(__file__).resolve().parents[3]
AUDIT_ROOT = WORKSPACE / "audit-core"
DEFAULT_PYTHON_SOURCE = Path(r"C:\Users\AZI\Desktop\VibeCode\3.open-data-audit")
DEFAULT_BINARY = Path(r"F:\RustBuild\audit\release\open-data-audit.exe")
DEFAULT_REPORT = AUDIT_ROOT / "probe" / "reports" / "s3f3_conflict_profile_20261030.json"
BASE_PROFILE_PATH = Path(__file__).with_name("perf_s3f3_parity_20261030.py")


def load_profile_helpers():
    spec = importlib.util.spec_from_file_location("s3f3_perf_profile_helpers", BASE_PROFILE_PATH)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load S3-F3 profile helpers from {BASE_PROFILE_PATH}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


HELPERS = load_profile_helpers()


def feedback_summary(result: Any) -> dict[str, int]:
    return {
        "imported": int(getattr(result, "imported", -1)),
        "unmatched": int(getattr(result, "unmatched", -1)),
        "conflicts": int(getattr(result, "conflicts", -1)),
    }


def build_conflicting_feedback(root: Path, rows: int) -> tuple[Path, Path, Path]:
    target, feedback_a = HELPERS.build_inputs(root, rows)
    feedback_b = root / "feedback_conflicting_reasons.xlsx"
    feedback_b.write_bytes(feedback_a.read_bytes())
    workbook = load_workbook(feedback_b)
    try:
        sheet = workbook["机构反馈"]
        for index in range(rows):
            sheet.cell(row=index + 2, column=8).value = f"待复核冲突原因{index:05d}"
        workbook.save(feedback_b)
    finally:
        workbook.close()
    return target, feedback_a, feedback_b


def conflict_command(binary: Path, target: Path, feedbacks: tuple[Path, Path]) -> list[str]:
    command = [
        str(binary),
        "s3-import-explanation-feedback",
        "--target-workbook",
        str(target),
    ]
    for feedback in feedbacks:
        command.extend(["--feedback-workbook", str(feedback)])
    command.append("--developer-view")
    return command


def rust_call(binary: Path, target: Path, feedbacks: tuple[Path, Path]) -> tuple[float, dict[str, Any]]:
    started = time.perf_counter_ns()
    completed = subprocess.run(
        conflict_command(binary, target, feedbacks),
        capture_output=True,
        text=True,
        encoding="utf-8",
        timeout=180,
    )
    wall_ms = (time.perf_counter_ns() - started) / 1_000_000
    if completed.returncode != 0:
        raise RuntimeError(f"Rust conflict import failed ({completed.returncode}): {completed.stderr[-4000:]}")
    try:
        report = json.loads(completed.stdout)
    except json.JSONDecodeError as error:
        raise RuntimeError(f"Rust S3-F3 returned invalid JSON: {completed.stdout[-1000:]}") from error
    return wall_ms, report


def run_python(source_root: Path, target: Path, feedbacks: tuple[Path, Path]) -> tuple[float, Any]:
    function = HELPERS.load_python_api(source_root)
    started = time.perf_counter_ns()
    result = function(target, list(feedbacks))
    return (time.perf_counter_ns() - started) / 1_000_000, result


def validate_conflict_report(path: Path, rows: int) -> dict[str, Any]:
    if not path.is_file():
        raise AssertionError(f"conflict workbook was not created: {path}")
    workbook = load_workbook(path, read_only=True, data_only=False)
    try:
        if workbook.sheetnames != ["冲突明细"]:
            raise AssertionError(f"unexpected conflict report sheets: {workbook.sheetnames}")
        sheet = workbook["冲突明细"]
        if sheet.max_row != rows + 1:
            raise AssertionError(f"expected {rows} conflict rows, got {sheet.max_row - 1}")
        if sheet.max_column < 10:
            raise AssertionError(f"conflict report has too few columns: {sheet.max_column}")
        first_reason = str(sheet.cell(row=2, column=9).value or "")
        first_cause = str(sheet.cell(row=2, column=10).value or "")
        if "已核实原因00000" not in first_reason or "待复核冲突原因00000" not in first_reason:
            raise AssertionError(f"conflict report lost one of the competing reasons: {first_reason}")
        if "多个反馈文件给出不同变动原因" not in first_cause:
            raise AssertionError(f"conflict report has an unexpected cause: {first_cause}")
        return {
            "sheet_count": len(workbook.sheetnames),
            "rows_including_header": sheet.max_row,
            "columns": sheet.max_column,
            "file_bytes": path.stat().st_size,
        }
    finally:
        workbook.close()


def verify_rust_backup(target: Path, original_hash: str) -> dict[str, Any]:
    backups = sorted(target.parent.glob(f"{target.stem}_导入前备份_*.xlsx"))
    if len(backups) != 1:
        raise AssertionError(f"expected exactly one Rust target backup, got {len(backups)}")
    backup_hash = HELPERS.sha256(backups[0])
    if backup_hash != original_hash:
        raise AssertionError("Rust backup does not match original target bytes")
    if HELPERS.sha256(target) != original_hash:
        raise AssertionError("conflict-only Rust execution changed the target bytes")
    return {"backup_sha256_matches_original": True, "target_sha256_unchanged": True}


def prepare_run_target(root: Path, label: str, original: Path) -> Path:
    directory = root / label
    directory.mkdir(parents=True)
    target = directory / "target.xlsx"
    target.write_bytes(original.read_bytes())
    return target


def execute_python_worker(source_root: Path, target: Path, feedbacks: tuple[Path, Path]) -> int:
    result = HELPERS.load_python_api(source_root)(target, list(feedbacks))
    print(json.dumps(feedback_summary(result)))
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--python-source", type=Path, default=DEFAULT_PYTHON_SOURCE)
    parser.add_argument("--binary", type=Path, default=DEFAULT_BINARY)
    parser.add_argument("--rows", type=int, default=1000)
    parser.add_argument("--rounds", type=int, default=3)
    parser.add_argument("--report", type=Path, default=DEFAULT_REPORT)
    parser.add_argument("--worker", action="store_true")
    parser.add_argument("--target", type=Path)
    parser.add_argument("--feedback-a", type=Path)
    parser.add_argument("--feedback-b", type=Path)
    args = parser.parse_args()

    if args.worker:
        if args.target is None or args.feedback_a is None or args.feedback_b is None:
            parser.error("--worker requires --target, --feedback-a, and --feedback-b")
        return execute_python_worker(args.python_source, args.target, (args.feedback_a, args.feedback_b))

    if args.rows <= 0 or args.rounds <= 0:
        parser.error("--rows and --rounds must be positive")
    if not args.python_source.is_dir():
        parser.error(f"Python source project does not exist: {args.python_source}")
    if not args.binary.is_file():
        parser.error(f"Rust Release binary does not exist: {args.binary}")

    python_module_path = (
        args.python_source
        / "core"
        / "src"
        / "base_audit"
        / "systems"
        / "s3_central_statistics"
        / "explanation_exporter.py"
    )
    python_hash_before = HELPERS.sha256(python_module_path)
    binary_hash = HELPERS.sha256(args.binary)
    rounds: list[dict[str, Any]] = []

    with tempfile.TemporaryDirectory(prefix="audit-s3f3-conflict-profile-") as temporary:
        root = Path(temporary)
        target_base, feedback_a, feedback_b = build_conflicting_feedback(root, args.rows)
        feedbacks = (feedback_a, feedback_b)
        original_target_hash = HELPERS.sha256(target_base)
        feedback_hashes = [HELPERS.sha256(path) for path in feedbacks]
        input_info = {
            "target": {"bytes": target_base.stat().st_size, "workbook": HELPERS.workbook_snapshot(target_base)},
            "feedback_a": {"bytes": feedback_a.stat().st_size, "workbook": HELPERS.workbook_snapshot(feedback_a)},
            "feedback_b": {"bytes": feedback_b.stat().st_size, "workbook": HELPERS.workbook_snapshot(feedback_b)},
        }
        input_info = {
            name: {
                "bytes": entry["bytes"],
                "sheets": [
                    {"name": sheet["name"], "rows": sheet["max_row"], "columns": sheet["max_column"]}
                    for sheet in entry["workbook"]["sheets"]
                ],
            }
            for name, entry in input_info.items()
        }

        for round_index in range(-1, args.rounds):
            label = "warmup" if round_index < 0 else f"round-{round_index + 1}"
            python_target = prepare_run_target(root, f"{label}-python", target_base)
            rust_target = prepare_run_target(root, f"{label}-rust", target_base)

            if round_index < 0 or round_index % 2 == 0:
                python_ms, python_result = run_python(args.python_source, python_target, feedbacks)
                rust_ms, rust_result = rust_call(args.binary, rust_target, feedbacks)
            else:
                rust_ms, rust_result = rust_call(args.binary, rust_target, feedbacks)
                python_ms, python_result = run_python(args.python_source, python_target, feedbacks)

            python_summary = feedback_summary(python_result)
            rust_summary = {
                "imported": int(rust_result.get("imported", -1)),
                "unmatched": int(rust_result.get("unmatched", -1)),
                "conflicts": int(rust_result.get("conflicts", -1)),
            }
            expected = {"imported": 0, "unmatched": 0, "conflicts": args.rows}
            if python_summary != expected or rust_summary != expected:
                raise AssertionError(f"unexpected conflict summaries: Python={python_summary}, Rust={rust_summary}")

            python_conflict_path = Path(python_result.conflict_report_path)
            rust_conflict_path = Path(str(rust_result["conflict_report_path"]))
            python_conflict_info = validate_conflict_report(python_conflict_path, args.rows)
            rust_conflict_info = validate_conflict_report(rust_conflict_path, args.rows)
            conflict_comparison = HELPERS.compare_workbooks(python_conflict_path, rust_conflict_path)
            target_comparison = HELPERS.compare_workbooks(python_target, rust_target)
            if conflict_comparison["difference_count"] or target_comparison["difference_count"]:
                raise AssertionError(
                    f"S3-F3 conflict output mismatch: conflict={conflict_comparison}, target={target_comparison}"
                )
            rust_backup = verify_rust_backup(rust_target, original_target_hash)
            if [HELPERS.sha256(path) for path in feedbacks] != feedback_hashes:
                raise AssertionError("S3-F3 conflict execution changed a feedback input")

            if round_index >= 0:
                rounds.append({
                    "round": round_index + 1,
                    "order": "python-first" if round_index % 2 == 0 else "rust-first",
                    "python_service_ms": python_ms,
                    "rust_cli_wall_ms": rust_ms,
                    "rust_internal_total_ms": rust_result.get("total_time_ms"),
                    "rust_stages_ms": {
                        "target_read": rust_result.get("target_read_time_ms"),
                        "feedback_read": rust_result.get("feedback_read_time_ms"),
                        "conflict_report_write": rust_result.get("conflict_write_time_ms"),
                        "writeback": rust_result.get("writeback_time_ms"),
                        "total": rust_result.get("total_time_ms"),
                    },
                    "summary": python_summary,
                    "conflict_report": {
                        "python": python_conflict_info,
                        "rust": rust_conflict_info,
                        "comparison": conflict_comparison,
                    },
                    "target_comparison": target_comparison,
                    "rust_backup": rust_backup,
                })

        memory_python_target = prepare_run_target(root, "memory-python", target_base)
        memory_rust_target = prepare_run_target(root, "memory-rust", target_base)
        python_memory = HELPERS.monitored_run([
            sys.executable,
            str(Path(__file__).resolve()),
            "--python-source",
            str(args.python_source),
            "--worker",
            "--target",
            str(memory_python_target),
            "--feedback-a",
            str(feedback_a),
            "--feedback-b",
            str(feedback_b),
        ])
        rust_memory = HELPERS.monitored_run(conflict_command(args.binary, memory_rust_target, feedbacks))
        python_memory_result = json.loads(python_memory["stdout"])
        rust_memory_result = json.loads(rust_memory["stdout"])
        if any(
            python_memory_result.get(key) != value
            for key, value in {"imported": 0, "unmatched": 0, "conflicts": args.rows}.items()
        ):
            raise AssertionError(f"unexpected Python memory-probe result: {python_memory_result}")
        if rust_memory_result.get("conflicts") != args.rows or rust_memory_result.get("imported") != 0:
            raise AssertionError(f"unexpected Rust memory-probe result: {rust_memory_result}")
        python_memory_outputs = list(memory_python_target.parent.glob("*导入冲突清单_*.xlsx"))
        rust_memory_outputs = list(memory_rust_target.parent.glob("*导入冲突清单_*.xlsx"))
        if len(python_memory_outputs) != 1 or len(rust_memory_outputs) != 1:
            raise AssertionError(
                "expected one memory-probe conflict report from each engine, "
                f"got Python={len(python_memory_outputs)}, Rust={len(rust_memory_outputs)}"
            )
        python_memory_conflicts = python_memory_outputs[0]
        rust_memory_conflicts = rust_memory_outputs[0]
        memory_output_comparison = HELPERS.compare_workbooks(python_memory_conflicts, rust_memory_conflicts)
        if memory_output_comparison["difference_count"]:
            raise AssertionError(f"memory-probe conflict reports differ: {memory_output_comparison}")

    python_hash_after = HELPERS.sha256(python_module_path)
    if python_hash_after != python_hash_before:
        raise AssertionError("the read-only Python source changed during S3-F3 conflict profiling")

    python_median = statistics.median(item["python_service_ms"] for item in rounds)
    rust_wall_median = statistics.median(item["rust_cli_wall_ms"] for item in rounds)
    rust_internal = [item["rust_internal_total_ms"] for item in rounds]
    rust_internal_median = statistics.median(rust_internal) if all(value is not None for value in rust_internal) else None
    report = {
        "status": "passed",
        "feature": "S3-F3",
        "scenario": "1000 matching keys with conflicting reasons from two feedback workbooks",
        "platform": platform.platform(),
        "python_version": sys.version.split()[0],
        "python_source": str(args.python_source),
        "python_source_sha256_unchanged": python_hash_before == python_hash_after,
        "python_source_sha256": python_hash_before,
        "rust_binary": str(args.binary),
        "rust_binary_sha256": binary_hash,
        "fixture": {
            "unique_matching_keys": args.rows,
            "feedback_workbooks": 2,
            "conflicting_keys": args.rows,
            "expected_imported": 0,
            "expected_unmatched": 0,
            "expected_conflicts": args.rows,
            "inputs": input_info,
            "feedback_inputs_sha256_unchanged": True,
        },
        "methodology": {
            "warmup_rounds": 1,
            "measured_rounds": args.rounds,
            "order": "alternating Python/Rust",
            "python_timing": "in-process service call; excludes interpreter/module startup and output comparison",
            "rust_timing": "Release CLI wall clock including process startup; service stages also recorded",
            "memory_sampling": "separate process-tree working-set probe sampled every 5 ms",
        },
        "timing_medians_ms": {
            "python_service": round(python_median, 2),
            "rust_cli_wall": round(rust_wall_median, 2),
            "rust_internal_total": rust_internal_median,
            "python_over_rust_cli": round(python_median / rust_wall_median, 3),
            "python_over_rust_internal": round(python_median / rust_internal_median, 3)
            if rust_internal_median
            else None,
        },
        "rounds": rounds,
        "memory": {
            "python_working_set_bytes": python_memory["sampled_peak_process_tree_working_set_bytes"],
            "rust_working_set_bytes": rust_memory["sampled_peak_process_tree_working_set_bytes"],
            "python_peak_live_process_count": python_memory["peak_live_process_count"],
            "rust_peak_live_process_count": rust_memory["peak_live_process_count"],
            "conflict_output_difference_count": memory_output_comparison["difference_count"],
        },
        "correctness": {
            "summary_differences": 0,
            "conflict_workbook_value_structure_differences": sum(
                item["conflict_report"]["comparison"]["difference_count"] for item in rounds
            ),
            "updated_target_differences": sum(item["target_comparison"]["difference_count"] for item in rounds),
            "rust_backup_matches_original": all(item["rust_backup"]["backup_sha256_matches_original"] for item in rounds),
            "rust_target_sha256_unchanged": all(item["rust_backup"]["target_sha256_unchanged"] for item in rounds),
        },
    }
    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    markdown_path = args.report.with_suffix(".md")
    markdown_path.write_text(render_markdown(report), encoding="utf-8")
    print(json.dumps({
        "status": report["status"],
        "python_median_ms": report["timing_medians_ms"]["python_service"],
        "rust_cli_wall_median_ms": report["timing_medians_ms"]["rust_cli_wall"],
        "rust_internal_median_ms": report["timing_medians_ms"]["rust_internal_total"],
        "python_over_rust_cli": report["timing_medians_ms"]["python_over_rust_cli"],
        "conflict_output_differences": report["correctness"]["conflict_workbook_value_structure_differences"],
        "peak_working_set_bytes": report["memory"],
        "report": str(args.report),
        "markdown": str(markdown_path),
    }, ensure_ascii=False, indent=2))
    return 0


def render_markdown(report: dict[str, Any]) -> str:
    timing = report["timing_medians_ms"]
    memory = report["memory"]
    python_mb = memory["python_working_set_bytes"] / (1024 * 1024)
    rust_mb = memory["rust_working_set_bytes"] / (1024 * 1024)
    lines = [
        "# S3-F3 冲突反馈性能与输出对拍（2026-10-30）",
        "",
        "## 负载",
        "",
        f"合成目标含 {report['fixture']['unique_matching_keys']:,} 个唯一匹配键；两份反馈工作簿对每个键给出不同原因，预期导入 0、未匹配 0、冲突 {report['fixture']['conflicting_keys']:,}。仅使用临时文件，旧 Python 项目只读。",
        "",
        "## 耗时",
        "",
        "| Python 服务 | Rust CLI 墙钟 | Rust 服务内部 | Python / Rust CLI |",
        "|---:|---:|---:|---:|",
        f"| {timing['python_service']:.2f} ms | {timing['rust_cli_wall']:.2f} ms | {timing['rust_internal_total']} ms | {timing['python_over_rust_cli']:.2f}× |",
        "",
        "Python 计时不含解释器和模块启动；Rust CLI 墙钟含进程启动。均在预热后交错运行。",
        "",
        "## 输出与内存",
        "",
        f"三轮摘要、目标工作簿和冲突清单值/结构均逐项对拍无差异；冲突清单每轮包含 {report['fixture']['conflicting_keys']:,} 条明细。Rust 目标在冲突场景保持原字节，备份与原目标 SHA-256 一致，反馈输入未变。",
        f"独立进程树采样峰值工作集：Python {python_mb:.2f} MB、Rust {rust_mb:.2f} MB；采样周期 5 ms。",
        "",
        "| 轮次 | 顺序 | Python | Rust CLI | Rust 内部 | 冲突写出 | 冲突清单字节 | 值/结构差异 |",
        "|---:|---|---:|---:|---:|---:|---:|---:|",
    ]
    for item in report["rounds"]:
        lines.append(
            f"| {item['round']} | {item['order']} | {item['python_service_ms']:.2f} ms | "
            f"{item['rust_cli_wall_ms']:.2f} ms | {item['rust_internal_total_ms']} ms | "
            f"{item['rust_stages_ms']['conflict_report_write']} ms | {item['conflict_report']['rust']['file_bytes']:,} | "
            f"{item['conflict_report']['comparison']['difference_count']} |"
        )
    lines += [
        "",
        "## 限制",
        "",
        "这是 100% 冲突的合成压力场景，不代表真实反馈冲突率。此专项验证冲突输出链路和高冲突规模的成本；既有无冲突 1,000 条性能基线仍单独保留。工作簿值与结构对拍，不做普通字体、边框或样式逐项比较。",
        "",
        f"复跑驱动：`tests/compare/perf_s3f3_conflict_profile_20261030.py`。Rust Release SHA-256：`{report['rust_binary_sha256']}`。",
        "",
    ]
    return "\n".join(lines)


if __name__ == "__main__":
    raise SystemExit(main())
