#!/usr/bin/env python3
"""Profile S1 structure preflight plus conditional audit, including each source read."""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import os
import statistics
import subprocess
import sys
import tempfile
import time
from datetime import datetime, timezone
from decimal import Decimal
from pathlib import Path
from typing import Any

import psutil


WORKSPACE = Path(__file__).resolve().parents[3]
AUDIT_ROOT = WORKSPACE / "audit-core"
PYTHON_SOURCE = Path(r"C:\Users\AZI\Desktop\VibeCode\3.open-data-audit")
DEFAULT_BINARY = Path(r"F:\RustBuild\audit\release\open-data-audit.exe")
DEFAULT_REPORT = AUDIT_ROOT / "probe" / "reports" / "s1_preflight_reads_20261030_shared_read_after.json"
PROFILE_DRIVER = AUDIT_ROOT / "tests" / "compare" / "perf_s1_conditional_profile_20261030.py"
SHEET = "S1Conditional"
HEADERS = ["id", "positive", "signed", "formula_target", "left", "right", "mapped", "text"]
PERIOD = "2026-09"
BATCH_ID = "s1-preflight-profile"
AUDIT_TIME = "2026-10-30T12:00:00+08:00"


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


def json_safe(value: Any) -> Any:
    if isinstance(value, Decimal):
        return int(value) if value == value.to_integral_value() else float(value)
    if isinstance(value, dict):
        return {str(key): json_safe(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [json_safe(item) for item in value]
    return value


def digest_issue_rows(issues: list[Any], normalize) -> list[str]:
    digests = []
    for issue in issues:
        row = json_safe(normalize(issue))
        payload = json.dumps(row, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
        digests.append(hashlib.sha256(payload.encode("utf-8")).hexdigest())
    return sorted(digests)


def sample_child_rss(process: subprocess.Popen[bytes]) -> float:
    try:
        child = psutil.Process(process.pid)
    except psutil.Error:
        return 0.0
    peak = 0
    while process.poll() is None:
        try:
            peak = max(peak, child.memory_info().rss)
        except psutil.Error:
            break
        time.sleep(0.005)
    try:
        peak = max(peak, child.memory_info().rss)
    except psutil.Error:
        pass
    return peak / (1024 * 1024)


def python_worker(source: Path, template: Path) -> dict[str, Any]:
    helper = load_module("s1_preflight_compare_helper", AUDIT_ROOT / "tests" / "compare" / "compare_s1_conditional_audit.py")
    sys.dont_write_bytecode = True
    sys.path.insert(0, str(PYTHON_SOURCE / "core" / "src"))
    from base_audit.conditional_evaluators import ConditionalFormatService
    from base_audit.models import CopyRange, TemplateDefinition
    from base_audit.preflight_xlsx import validate_source_xlsx
    from openpyxl import load_workbook

    structure_range = CopyRange(SHEET, "A1:H1")
    template_started = time.perf_counter()
    template_book = load_workbook(template, read_only=True, data_only=False, keep_links=False)
    try:
        expected_headers = next(template_book[SHEET].iter_rows(
            min_row=1, max_row=1, min_col=1, max_col=8, values_only=True
        ))
    finally:
        template_book.close()
    template_elapsed_ms = (time.perf_counter() - template_started) * 1000
    definition = TemplateDefinition(
        rules=[], copy_ranges=[structure_range], structured=False,
        structure_ranges=[structure_range],
    )
    expected = [(structure_range, [list(expected_headers)])]
    preflight_started = time.perf_counter()
    match = validate_source_xlsx(source, definition, expected)
    preflight_elapsed_ms = (time.perf_counter() - preflight_started) * 1000
    if not match.matched:
        raise RuntimeError(f"Python structure preflight rejected generated source: {match.details}")

    service = ConditionalFormatService(mode="PYTHON")
    issues, stats = service.extract_issues(
        workbook_path=source,
        ranges=None,
        structure_ranges=[structure_range],
        period=PERIOD,
        batch_id=BATCH_ID,
        audit_time=AUDIT_TIME,
        source_file=source,
        mapping=None,
    )
    return {
        "issue_digests": digest_issue_rows(issues, helper.normalize_issue),
        "issue_count": len(issues),
        "template_read_ms": round(template_elapsed_ms, 3),
        "preflight": {
            "matched": match.matched,
            "matched_labels": match.matched_labels,
            "checked_labels": match.checked_labels,
            "elapsed_ms": round(preflight_elapsed_ms, 3),
        },
        "conditional": {
            "elapsed_ms": round(float(stats["total_time"]) * 1000, 3),
            "normalize_ms": round(float(stats.get("rule_normalize_time", 0)) * 1000, 3),
            "evaluation_ms": round(float(stats.get("evaluation_time", 0)) * 1000, 3),
            "issue_build_ms": round(float(stats.get("result_build_time", 0)) * 1000, 3),
            "issues": len(issues),
        },
    }


def run_worker(command: list[str], out_path: Path, err_path: Path, env: dict[str, str] | None = None) -> tuple[Any, float, float, str]:
    started = time.perf_counter()
    with out_path.open("wb") as stdout, err_path.open("wb") as stderr:
        process = subprocess.Popen(command, stdout=stdout, stderr=stderr, env=env)
        peak_mb = sample_child_rss(process)
        code = process.wait()
    external_ms = (time.perf_counter() - started) * 1000
    err = err_path.read_text(encoding="utf-8", errors="replace")
    if code != 0:
        raise RuntimeError(f"worker failed ({code}): {err[-4000:]}")
    result = json.loads(out_path.read_text(encoding="utf-8"))
    return result, external_ms, peak_mb, err


def run_python(source: Path, template: Path, out_dir: Path, index: int) -> dict[str, Any]:
    command = [sys.executable, str(Path(__file__).resolve()), "--python-worker", str(source), str(template)]
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    result, external, peak, _ = run_worker(command, out_dir / f"py-{index}.json", out_dir / f"py-{index}.log", env)
    result["external_ms"] = round(external, 3)
    result["peak_rss_mb"] = round(peak, 3)
    return result


def run_rust(binary: Path, source: Path, template: Path, rows: int, out_dir: Path, index: int, normalize_issue) -> dict[str, Any]:
    command = [
        str(binary), "conditional-audit", "--workbook", str(source), "--source-file", str(source),
        "--template-workbook", str(template), "--structure-area", f"{SHEET}!A1:H1",
        "--period", PERIOD, "--batch-id", BATCH_ID, "--audit-time", AUDIT_TIME,
        "--developer-view",
    ]
    for address in (f"B2:B{rows + 1}", f"C2:C{rows + 1}", f"D2:D{rows + 1}", f"G2:G{rows + 1}"):
        command.extend(["--scan-area", f"{SHEET}!{address}"])
    report, external, peak, log = run_worker(
        command, out_dir / f"rs-{index}.json", out_dir / f"rs-{index}.log"
    )
    events = [json.loads(line) for line in log.splitlines() if line.strip()]
    file_preflight = next(event for event in events if event.get("step") == "structure_preflight_file")
    preflight = next(event for event in events if event.get("step") == "structure_preflight" and event.get("status") == "completed")
    conditional_global = next(
        event for event in events
        if event.get("step") == "conditional_format_issue_extraction"
        and event.get("status") == "completed"
    )
    conditional_files = [
        event["event"] for event in events
        if event.get("step") == "conditional_format_issue_extraction"
        and isinstance(event.get("event"), dict)
        and event["event"].get("status") == "completed"
    ]
    conditional_file = conditional_files[0] if conditional_files else conditional_global
    return {
        "issue_digests": digest_issue_rows(report.get("issues", []), normalize_issue),
        "issue_count": len(report.get("issues", [])),
        "preflight": {
            "matched": file_preflight.get("status") == "matched",
            "matched_labels": file_preflight.get("matched_labels"),
            "checked_labels": file_preflight.get("checked_labels"),
            "source_read_ms": file_preflight.get("read_time_ms"),
            "elapsed_ms": file_preflight.get("elapsed_ms"),
            "summary_elapsed_ms": preflight.get("elapsed_ms"),
            "template_read_ms": preflight.get("template_read_time_ms"),
        },
        "conditional": {
            "elapsed_ms": conditional_file.get("total_time_ms", conditional_file.get("elapsed_ms")),
            "command_elapsed_ms": conditional_global.get("elapsed_ms"),
            "combined_command_elapsed_ms": conditional_global.get("elapsed_ms"),
            "source_read_ms": conditional_file.get("read_time_ms"),
            "shared_source_read_ms": conditional_global.get("shared_source_read_ms", file_preflight.get("read_time_ms")),
            "evaluation_ms": conditional_file.get("evaluation_time_ms"),
            "issue_build_ms": conditional_file.get("issue_build_time_ms"),
            "issues": conditional_file.get("issue_count", conditional_file.get("issues")),
        },
        "external_ms": round(external, 3),
        "peak_rss_mb": round(peak, 3),
    }


def issue_digest_parity(py_digests: list[str], rs_digests: list[str]) -> dict[str, Any]:
    py_counts: dict[str, int] = {}
    rs_counts: dict[str, int] = {}
    for digest in py_digests:
        py_counts[digest] = py_counts.get(digest, 0) + 1
    for digest in rs_digests:
        rs_counts[digest] = rs_counts.get(digest, 0) + 1
    missing_in_rust = sum(max(0, count - rs_counts.get(digest, 0)) for digest, count in py_counts.items())
    missing_in_python = sum(max(0, count - py_counts.get(digest, 0)) for digest, count in rs_counts.items())
    return {
        "python_issues": len(py_digests),
        "rust_issues": len(rs_digests),
        "missing_in_rust": missing_in_rust,
        "missing_in_python": missing_in_python,
        "canonical_issue_sha256_multiset_matches": py_digests == rs_digests,
    }


def median(items: list[float]) -> float:
    return round(statistics.median(items), 3)


def main() -> int:
    if len(sys.argv) == 4 and sys.argv[1] == "--python-worker":
        print(json.dumps(json_safe(python_worker(Path(sys.argv[2]), Path(sys.argv[3]))), ensure_ascii=False))
        return 0

    parser = argparse.ArgumentParser()
    parser.add_argument("--binary", type=Path, default=DEFAULT_BINARY)
    parser.add_argument("--rows", type=int, default=4000)
    parser.add_argument("--rounds", type=int, default=5)
    parser.add_argument("--warmup", type=int, default=1)
    parser.add_argument("--report", type=Path, default=DEFAULT_REPORT)
    args = parser.parse_args()
    if args.rows < 1 or args.rounds < 1 or args.warmup < 0:
        parser.error("rows/rounds must be positive and warmup nonnegative")

    if not args.binary.is_file():
        raise FileNotFoundError(args.binary)
    source_core = PYTHON_SOURCE / "core"
    if not source_core.is_dir():
        raise FileNotFoundError(source_core)
    helper = load_module("s1_preflight_compare_parent", AUDIT_ROOT / "tests" / "compare" / "compare_s1_conditional_audit.py")
    fixture_module = load_module("s1_preflight_fixture", PROFILE_DRIVER)
    records: list[dict[str, Any]] = []
    parity_records: list[dict[str, Any]] = []
    py_peaks: list[float] = []
    rs_peaks: list[float] = []
    with tempfile.TemporaryDirectory(prefix="audit_s1_preflight_profile_") as root_name:
        root = Path(root_name)
        source = root / "S1-Source.xlsx"
        template = root / "S1-Template.xlsx"
        fixture = fixture_module.generate_fixture(source, args.rows)
        template.write_bytes(source.read_bytes())
        source_hash, template_hash = sha256(source), sha256(template)
        py_index = rs_index = 0

        def run_py(label: str) -> dict[str, Any]:
            nonlocal py_index
            py_index += 1
            path = root / f"py-{label}-{py_index}"
            path.mkdir()
            result = run_python(source, template, path, py_index)
            if label == "measure":
                py_peaks.append(result["peak_rss_mb"])
            return result

        def run_rs(label: str, round_index: int) -> dict[str, Any]:
            nonlocal rs_index
            rs_index += 1
            path = root / f"rs-{label}-{rs_index}"
            path.mkdir()
            result = run_rust(
                args.binary.resolve(), source, template, args.rows, path, rs_index,
                helper.normalize_issue,
            )
            if label == "measure":
                rs_peaks.append(result["peak_rss_mb"])
            return result

        for index in range(args.warmup):
            py_warm = run_py("warm")
            rs_warm = run_rs("warm", index)
            check = issue_digest_parity(py_warm["issue_digests"], rs_warm["issue_digests"])
            if check["missing_in_rust"] or check["missing_in_python"] or not check["canonical_issue_sha256_multiset_matches"]:
                raise AssertionError(f"warmup issue mismatch: {check}")

        for index in range(args.rounds):
            if index % 2 == 0:
                rs_result = run_rs("measure", index)
                py_result = run_py("measure")
            else:
                py_result = run_py("measure")
                rs_result = run_rs("measure", index)
            check = issue_digest_parity(py_result["issue_digests"], rs_result["issue_digests"])
            if check["missing_in_rust"] or check["missing_in_python"] or not check["canonical_issue_sha256_multiset_matches"]:
                raise AssertionError(f"round {index + 1} issue mismatch: {check}")
            parity_records.append(check)
            records.append({
                "round": index + 1,
                "order": "rust_then_python" if index % 2 == 0 else "python_then_rust",
                "python": {key: py_result[key] for key in ("template_read_ms", "preflight", "conditional", "peak_rss_mb")},
                "rust": {key: rs_result[key] for key in ("preflight", "conditional", "peak_rss_mb")},
                "parity": check,
            })

        final_hashes = {"source": sha256(source), "template": sha256(template)}
        source_bytes = source.stat().st_size
        template_bytes = template.stat().st_size
        if final_hashes != {"source": source_hash, "template": template_hash}:
            raise AssertionError("source or template bytes changed during profiling")

    py_preflight = [row["python"]["preflight"]["elapsed_ms"] for row in records]
    rs_preflight = [row["rust"]["preflight"]["elapsed_ms"] for row in records]
    py_cf = [row["python"]["conditional"]["elapsed_ms"] for row in records]
    rs_cf = [row["rust"]["conditional"]["elapsed_ms"] for row in records]
    py_combined = [row["python"]["template_read_ms"] + row["python"]["preflight"]["elapsed_ms"] + row["python"]["conditional"]["elapsed_ms"] for row in records]
    rs_combined = [row["rust"]["conditional"]["combined_command_elapsed_ms"] for row in records]
    report = {
        "schema_version": 1,
        "suite": "audit.s1_structure_preflight_and_conditional_audit_profile",
        "status": "passed",
        "created_at": datetime.now(timezone.utc).isoformat(),
        "environment": {
            "platform": sys.platform,
            "python_version": sys.version.split()[0],
            "rust_binary": str(args.binary.resolve()),
            "rust_binary_sha256": sha256(args.binary.resolve()),
            "timer_scope": "Python template read + structure preflight + conditional service internal totals; Rust conditional-audit developer event elapsed includes one shared source read, template read, structure preflight, conditional evaluation/issue build, and CLI work after argument dispatch; it excludes process startup.",
            "memory_sampling_interval_ms": 5,
        },
        "input": {
            **fixture,
            "source_bytes": source_bytes,
            "template_bytes": template_bytes,
            "source_sha256_before": source_hash,
            "source_sha256_after": final_hashes["source"],
            "template_sha256_before": template_hash,
            "template_sha256_after": final_hashes["template"],
        },
        "runs": {"warmup": args.warmup, "measured": args.rounds},
        "medians_ms": {
            "python_template_read_once": median([row["python"]["template_read_ms"] for row in records]),
            "python_structure_preflight": median(py_preflight),
            "rust_structure_preflight_file_elapsed": median(rs_preflight),
            "rust_structure_source_read": median([row["rust"]["preflight"]["source_read_ms"] for row in records]),
            "rust_template_read_once": median([row["rust"]["preflight"]["template_read_ms"] for row in records]),
            "python_conditional_service": median(py_cf),
            "rust_conditional_read": median([row["rust"]["conditional"]["source_read_ms"] for row in records]),
            "rust_source_read_shared_with_preflight": median([row["rust"]["conditional"]["shared_source_read_ms"] for row in records]),
            "rust_conditional_evaluation": median([row["rust"]["conditional"]["evaluation_ms"] for row in records]),
            "rust_conditional_issue_build": median([row["rust"]["conditional"]["issue_build_ms"] for row in records]),
            "python_preflight_plus_conditional": median(py_combined),
            "rust_preflight_plus_conditional": median(rs_combined),
            "python_over_rust_combined_ratio": round(statistics.median(py_combined) / statistics.median(rs_combined), 4),
        },
        "peak_rss_mb_median": {
            "python_process": median(py_peaks),
            "rust_process": median(rs_peaks),
        },
        "correctness": {
            "all_rounds_issue_parity": all(row["canonical_issue_sha256_multiset_matches"] and row["missing_in_rust"] == 0 and row["missing_in_python"] == 0 for row in parity_records),
            "rounds": parity_records,
            "source_unchanged": source_hash == final_hashes["source"],
            "template_unchanged": template_hash == final_hashes["template"],
        },
        "measured_rounds": records,
    }
    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
