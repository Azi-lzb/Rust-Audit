#!/usr/bin/env python3
"""End-to-end S1-F1 Python/Rust profile using the shared synthetic golden fixture.

The elapsed-time comparison covers the full feature path, not a standalone COM
cold-start probe. Peak RSS includes each worker tree and only Excel processes
created by that measured run; pre-existing Excel instances are excluded.
"""

from __future__ import annotations

import argparse
import gc
import hashlib
import importlib.util
import itertools
import json
import math
import os
import platform
import re
import shutil
import statistics
import subprocess
import sys
import tempfile
import time
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.comments import Comment
import psutil

sys.dont_write_bytecode = True

WORKSPACE = Path(__file__).resolve().parents[3]
PYTHON_SOURCE = Path(r"C:\Users\AZI\Desktop\VibeCode\3.open-data-audit")
GOLDEN_TESTS = PYTHON_SOURCE / "core" / "tests" / "golden"
DEFAULT_BINARY = Path(r"F:\RustBuild\audit\release\open-data-audit.exe")
DEFAULT_REPORT = WORKSPACE / "audit-core" / "probe" / "reports" / (
    f"s1f1_full_chain_{datetime.now(timezone.utc).strftime('%Y%m%d_%H%M%S_%f_utc')}.json"
)
PERIOD = "2026-06"
DATA_SHEET = "20202资产负债表"
ISSUE_SHEET = "本期审核结果"


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def git_revision() -> str:
    completed = subprocess.run(
        ["git", "rev-parse", "--short", "HEAD"], cwd=WORKSPACE,
        check=True, capture_output=True, text=True,
    )
    return completed.stdout.strip()


def load_golden_fixtures():
    sys.path.insert(0, str(PYTHON_SOURCE / "core" / "src"))
    sys.path.insert(0, str(GOLDEN_TESTS))
    spec = importlib.util.spec_from_file_location("s1f1_perf_golden_fixtures", GOLDEN_TESTS / "golden_fixtures.py")
    if spec is None or spec.loader is None:
        raise RuntimeError("cannot load the read-only Python golden fixture helper")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


class PreserveFailedTemporaryDirectory:
    """Clean successful benchmark data, but preserve failed runs for diagnosis."""

    def __init__(self, prefix: str):
        self.name = tempfile.mkdtemp(prefix=prefix)
        self.preserve_requested = False

    def __enter__(self) -> str:
        return self.name

    def preserve(self) -> None:
        self.preserve_requested = True

    def __exit__(self, exception_type, exception, traceback) -> bool:
        if exception_type is None and not self.preserve_requested:
            shutil.rmtree(self.name, ignore_errors=True)
        else:
            print(f"Preserved failed S1-F1 benchmark workspace: {self.name}", file=sys.stderr)
        return False


def make_fixture(root: Path) -> dict[str, Path]:
    fixtures = load_golden_fixtures()
    source_dir = root / PERIOD / "源数据"
    template_dir = root / "模板"
    source_dir.mkdir(parents=True)
    template_dir.mkdir()

    fixtures.build_source(
        source_dir / "甲银行.xlsx", deposit=2_000_000, loan=3_000_000, reserve=30_000,
    )
    fixtures.build_source(
        source_dir / "乙银行.xlsx", deposit=5_000_000, loan=4_000_000, reserve=10_000,
        comment=fixtures.CF_COMMENT,
    )
    fixtures.build_bad_source(source_dir / "丙银行.xlsx")
    template = template_dir / "模板.xlsx"
    fixtures.build_audit_template(template)
    template_book = openpyxl.load_workbook(template)
    template_book["审核规则"]["A1"].comment = Comment("规则表说明", "测试")
    template_book.save(template)
    template_book.close()
    external = root / "外部参照.xlsx"
    fixtures.build_external(external)
    config = fixtures.make_config(root)
    fixtures.write_history_row(
        config,
        workbook_name="甲银行.xlsx",
        rule_number="甲银行｜20202资产负债表｜D4｜贷款大于存款｜贷款高于存款",
        feedback="去年已说明",
        opinion="继续关注",
    )
    return {
        "source_dir": source_dir,
        "template_dir": template_dir,
        "template": template,
        "external": external,
        "config": config,
    }


def excel_inventory() -> dict[int, float]:
    result: dict[int, float] = {}
    for process in psutil.process_iter(["pid", "name", "create_time"]):
        try:
            if (process.info.get("name") or "").casefold() == "excel.exe":
                created = process.info.get("create_time")
                if created is not None:
                    result[int(process.info["pid"])] = float(created)
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            continue
    return result


def measure_process(
    command: list[str], *, environment: dict[str, str] | None = None,
    timeout_seconds: float = 240.0, sample_interval_s: float = 0.005,
) -> tuple[dict[str, Any], bytes, bytes]:
    baseline_excel = excel_inventory()
    new_excel: dict[int, float] = {}
    known_pids: set[int] = set()
    peak_rss = 0
    peak_process_count = 0
    sample_count = 0
    started = time.perf_counter()
    with tempfile.TemporaryFile() as stdout_file, tempfile.TemporaryFile() as stderr_file:
        process = subprocess.Popen(
            command, stdout=stdout_file, stderr=stderr_file, env=environment,
        )
        root = psutil.Process(process.pid)
        known_pids.add(process.pid)
        timed_out = False
        while True:
            try:
                current = [root, *root.children(recursive=True)]
            except psutil.Error:
                current = []
            for child in current:
                known_pids.add(child.pid)
            inventory = excel_inventory()
            for process_id, created in inventory.items():
                if process_id not in baseline_excel:
                    new_excel.setdefault(process_id, created)

            candidates = known_pids | set(new_excel)
            live = 0
            rss = 0
            for process_id in candidates:
                try:
                    candidate = psutil.Process(process_id)
                    if not candidate.is_running() or candidate.status() == psutil.STATUS_ZOMBIE:
                        continue
                    rss += int(candidate.memory_info().rss)
                    live += 1
                except psutil.Error:
                    continue
            peak_rss = max(peak_rss, rss)
            peak_process_count = max(peak_process_count, live)
            sample_count += 1

            worker_done = process.poll() is not None
            live_excel = []
            for process_id, created in new_excel.items():
                try:
                    candidate = psutil.Process(process_id)
                    if abs(candidate.create_time() - created) <= 0.05 and candidate.is_running():
                        live_excel.append(process_id)
                except psutil.Error:
                    continue
            live_workers = []
            for process_id in known_pids:
                if process_id in new_excel or process_id == process.pid:
                    continue
                try:
                    candidate = psutil.Process(process_id)
                    if candidate.is_running() and candidate.status() != psutil.STATUS_ZOMBIE:
                        live_workers.append(process_id)
                except psutil.Error:
                    continue
            if worker_done and not live_excel and not live_workers:
                break
            if time.perf_counter() - started > timeout_seconds:
                timed_out = True
                try:
                    process.terminate()
                    process.wait(timeout=5)
                except (psutil.Error, subprocess.TimeoutExpired):
                    pass
                break
            time.sleep(sample_interval_s)

        # Capture only Excel instances created by this run. Never target PIDs
        # that existed before the worker started.
        terminated_excel: list[int] = []
        for process_id, created in new_excel.items():
            try:
                candidate = psutil.Process(process_id)
                if abs(candidate.create_time() - created) <= 0.05 and candidate.is_running():
                    candidate.terminate()
                    candidate.wait(timeout=5)
                    terminated_excel.append(process_id)
            except psutil.NoSuchProcess:
                continue
            except (psutil.AccessDenied, psutil.TimeoutExpired, psutil.ZombieProcess):
                continue

        if process.poll() is None:
            try:
                process.wait(timeout=10)
            except subprocess.TimeoutExpired:
                process.kill()
                process.wait(timeout=5)
        stdout_file.seek(0)
        stderr_file.seek(0)
        stdout = stdout_file.read()
        stderr = stderr_file.read()

    remaining_excel = []
    for process_id, created in new_excel.items():
        try:
            candidate = psutil.Process(process_id)
            if abs(candidate.create_time() - created) <= 0.05 and candidate.is_running():
                remaining_excel.append(process_id)
        except psutil.Error:
            continue
    sample = {
        "return_code": process.returncode,
        "timed_out": timed_out,
        "peak_sampled_worker_plus_new_excel_working_set_bytes": peak_rss,
        "peak_sampled_worker_plus_new_excel_working_set_mb": round(peak_rss / (1024 * 1024), 3),
        "peak_live_process_count": peak_process_count,
        "sample_interval_ms": round(sample_interval_s * 1000, 3),
        "memory_samples": sample_count,
        "new_excel_process_ids": sorted(new_excel),
        "terminated_owned_excel_process_ids": sorted(terminated_excel),
        "remaining_new_excel_process_ids": sorted(remaining_excel),
        "elapsed_wall_ms": round((time.perf_counter() - started) * 1000, 2),
    }
    if timed_out or remaining_excel:
        raise RuntimeError(
            f"worker failed or did not clean up (rc={process.returncode}, timeout={timed_out}, "
            f"remaining Excel={remaining_excel}):\n{stderr.decode('utf-8', errors='replace')[-6000:]}"
        )
    return sample, stdout, stderr


def python_worker(source_dir: Path, template_dir: Path, external: Path, config: Path, output_dir: Path, python_source: Path) -> int:
    sys.dont_write_bytecode = True
    sys.path.insert(0, str(python_source / "core" / "src"))
    from base_audit.service import AuditService
    from base_audit.workflow.runner import run_dag_flow

    started = time.perf_counter()
    result = run_dag_flow(
        "dag:汇总核查表校验",
        service=AuditService(config_path=config, engine_preference="自动"),
        template_path=next(template_dir.glob("*.xlsx")),
        input_dir=source_dir,
        output_dir=output_dir,
        period=PERIOD,
        history_path=config,
        external_path=external,
        recursive=False,
        write_flow_logs=True,
    )
    elapsed_ms = round((time.perf_counter() - started) * 1000, 2)
    print(json.dumps({
        "status": "completed",
        "elapsed_ms": elapsed_ms,
        "summary_path": str(result.summary_path) if result.summary_path else None,
        "issue_count": len(result.current_issues),
        "file_count": len(result.files),
        "performance_lines": list(result.performance_lines),
        "output_dir": str(output_dir),
    }, ensure_ascii=True))
    return 0


def canonical(value: Any) -> Any:
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    if isinstance(value, str):
        value = re.sub(r"20\d{2}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}", "<DATETIME>", value)
        return re.sub(r"\d{8}_?\d{6}", "<TIMESTAMP>", value)
    return value


def compare_result_workbooks(python_path: Path, rust_path: Path) -> dict[str, Any]:
    python_book = openpyxl.load_workbook(python_path, read_only=True, data_only=True)
    rust_book = openpyxl.load_workbook(rust_path, read_only=True, data_only=True)
    try:
        python_names = list(python_book.sheetnames)
        rust_names = list(rust_book.sheetnames)
        differences: list[dict[str, Any]] = []
        compared_cells = 0
        dimensions_equal = python_names == rust_names
        for name in python_names:
            if name not in rust_book.sheetnames:
                continue
            left = python_book[name]
            right = rust_book[name]
            dimensions_equal &= left.max_row == right.max_row and left.max_column == right.max_column
            max_row = max(left.max_row, right.max_row)
            max_column = max(left.max_column, right.max_column)
            for row_index, (left_row, right_row) in enumerate(itertools.zip_longest(
                left.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_column, values_only=True),
                right.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_column, values_only=True),
                fillvalue=(),
            ), 1):
                for column_index, (left_value, right_value) in enumerate(itertools.zip_longest(left_row, right_row), 1):
                    compared_cells += 1
                    same = False
                    if isinstance(left_value, (int, float)) and not isinstance(left_value, bool) and isinstance(right_value, (int, float)) and not isinstance(right_value, bool):
                        same = math.isclose(float(left_value), float(right_value), rel_tol=0.0, abs_tol=1e-9)
                    else:
                        same = canonical(left_value) == canonical(right_value)
                    if not same and len(differences) < 30:
                        differences.append({
                            "sheet": name,
                            "row": row_index,
                            "column": column_index,
                            "python": repr(left_value),
                            "rust": repr(right_value),
                        })
        return {
            "passed": not differences and dimensions_equal,
            "python_sheet_names": python_names,
            "rust_sheet_names": rust_names,
            "dimensions_equal": dimensions_equal,
            "compared_cells": compared_cells,
            "cell_differences": len(differences),
            "difference_samples": differences,
            "python_result_rows": max(0, python_book[ISSUE_SHEET].max_row - 1) if ISSUE_SHEET in python_book.sheetnames else None,
            "rust_result_rows": max(0, rust_book[ISSUE_SHEET].max_row - 1) if ISSUE_SHEET in rust_book.sheetnames else None,
        }
    finally:
        python_book.close()
        rust_book.close()


def find_rust_summary(output_dir: Path) -> Path:
    matches = []
    for path in output_dir.rglob("*.xlsx"):
        try:
            book = openpyxl.load_workbook(path, read_only=True, data_only=True)
            is_summary = ISSUE_SHEET in book.sheetnames
            book.close()
            if is_summary:
                matches.append(path)
        except Exception:
            continue
    if len(matches) != 1:
        raise AssertionError(f"expected one Rust S1 result workbook, found {matches}")
    return matches[0]


def find_audit_copy(output_dir: Path, source_stem: str) -> Path:
    matches = []
    for path in output_dir.rglob("*.xlsx"):
        if source_stem not in path.stem:
            continue
        try:
            book = openpyxl.load_workbook(path, read_only=True, data_only=False)
            is_audit_copy = DATA_SHEET in book.sheetnames and str(book[DATA_SHEET]["D4"].value or "").startswith("=")
            book.close()
            if is_audit_copy:
                matches.append(path)
        except Exception:
            continue
    if len(matches) != 1:
        raise AssertionError(f"expected one audit copy for {source_stem}, found {matches}")
    return matches[0]


def formula_copy_check(python_output: Path, rust_output: Path) -> dict[str, Any]:
    books = {"python": {}, "rust": {}}
    for engine, directory in (("python", python_output), ("rust", rust_output)):
        for source_stem in ("甲银行", "乙银行"):
            path = find_audit_copy(directory, source_stem)
            formula_book = openpyxl.load_workbook(path, data_only=False)
            value_book = openpyxl.load_workbook(path, read_only=True, data_only=True)
            try:
                formulas = {cell: formula_book[DATA_SHEET][cell].value for cell in ("D4", "D5")}
                values = {cell: value_book[DATA_SHEET][cell].value for cell in ("D4", "D5")}
                sheets = list(formula_book.sheetnames)
                rule_sheet_cells = {
                    cell.coordinate: cell.value
                    for row in formula_book["审核规则"].iter_rows()
                    for cell in row
                    if cell.value is not None
                } if "审核规则" in sheets else None
                rule_sheet_comments = {
                    cell.coordinate: {"text": cell.comment.text, "author": cell.comment.author}
                    for row in formula_book["审核规则"].iter_rows()
                    for cell in row
                    if cell.comment is not None
                } if "审核规则" in sheets else None
                books[engine][source_stem] = {
                    "formulas": formulas,
                    "values": values,
                    "sheets": sheets,
                    "rule_sheet_cells": rule_sheet_cells,
                    "rule_sheet_comments": rule_sheet_comments,
                    "unused_external_sheet_absent": "未引用表" not in sheets,
                }
            finally:
                formula_book.close()
                value_book.close()

    expected = {
        "甲银行": {"D4": "软性|贷款大于存款|贷款高于存款", "D5": None},
        "乙银行": {"D4": None, "D5": "错误|拨备核对|拨备不足|拨备低于存款百分之一|||"},
    }
    formula_equivalence = True
    cache_equivalence = True
    expected_cache_results = True
    required_sheets = True
    rule_sheet_content_equivalence = True
    rule_sheet_comment_equivalence = True
    sheet_order_equivalence = True
    for source_stem in expected:
        py = books["python"][source_stem]
        rust = books["rust"][source_stem]
        formula_equivalence &= all(
            canonical_formula(py["formulas"][cell]) == canonical_formula(rust["formulas"][cell])
            for cell in ("D4", "D5")
        )
        cache_equivalence &= all(canonical(py["values"][cell]) == canonical(rust["values"][cell]) for cell in ("D4", "D5"))
        expected_cache_results &= all(canonical(py["values"][cell]) == expected[source_stem][cell] for cell in ("D4", "D5"))
        rule_sheet_content_equivalence &= canonical(py["rule_sheet_cells"]) == canonical(rust["rule_sheet_cells"])
        rule_sheet_comment_equivalence &= canonical(py["rule_sheet_comments"]) == canonical(rust["rule_sheet_comments"])
        sheet_order_equivalence &= py["sheets"] == rust["sheets"]
        required_sheets &= (
            "参照表" in py["sheets"]
            and "参照表" in rust["sheets"]
            and "审核规则" in py["sheets"]
            and "审核规则" in rust["sheets"]
            and py["unused_external_sheet_absent"]
            and rust["unused_external_sheet_absent"]
        )
    return {
        "formula_text_equal": formula_equivalence,
        "formula_cache_equal": cache_equivalence,
        "formula_cache_matches_independent_expected_values": expected_cache_results,
        "referenced_external_sheet_copied_and_unreferenced_sheet_omitted": required_sheets,
        "structured_rule_sheet_content_equal": rule_sheet_content_equivalence,
        "structured_rule_sheet_comments_equal": rule_sheet_comment_equivalence,
        "worksheet_names_and_order_equal": sheet_order_equivalence,
        "workbooks": books,
    }


def canonical_formula(value: Any) -> str | None:
    if not isinstance(value, str):
        return None
    return re.sub(r"\s+", "", value.replace("'参照表'!", "参照表!").replace("\"", "\""))


def verify_excel_can_open(paths: list[Path]) -> dict[str, Any]:
    """Open result copies in a fresh Excel COM instance, outside timed runs."""
    import win32com.client
    import win32process

    baseline_excel = excel_inventory()
    application = win32com.client.DispatchEx("Excel.Application")
    application.Visible = False
    application.DisplayAlerts = False
    application.AskToUpdateLinks = False
    application.EnableEvents = False
    process_id = int(win32process.GetWindowThreadProcessId(application.Hwnd)[1])
    if process_id in baseline_excel:
        application.Quit()
        raise RuntimeError("Excel COM validation unexpectedly attached to a pre-existing Excel process")

    workbooks: list[dict[str, Any]] = []
    try:
        for path in paths:
            workbook = application.Workbooks.Open(
                str(path.resolve()),
                UpdateLinks=0,
                ReadOnly=True,
                IgnoreReadOnlyRecommended=True,
                AddToMru=False,
            )
            try:
                sheet = workbook.Worksheets(DATA_SHEET)
                formulas = {
                    cell: str(sheet.Range(cell).Formula)
                    for cell in ("D4", "D5")
                    if str(sheet.Range(cell).Formula).startswith("=")
                }
                workbooks.append({
                    "path": path.name,
                    "worksheet_count": int(workbook.Worksheets.Count),
                    "formula_texts": formulas,
                })
            finally:
                workbook.Close(SaveChanges=False)
    finally:
        application.Quit()
        del application
        gc.collect()

    try:
        psutil.Process(process_id).wait(timeout=25)
        process_exited = True
    except psutil.NoSuchProcess:
        process_exited = True
    except psutil.TimeoutExpired:
        process_exited = False
    return {
        "opened_all": len(workbooks) == len(paths),
        "owned_excel_pid": process_id,
        "owned_excel_process_exited": process_exited,
        "workbooks": workbooks,
    }


def create_stats(inputs: dict[str, Path]) -> dict[str, Any]:
    paths = [*sorted(inputs["source_dir"].glob("*.xlsx")), inputs["template"], inputs["external"], inputs["config"]]
    stats: dict[str, Any] = {}
    for path in paths:
        book = openpyxl.load_workbook(path, read_only=True, data_only=False)
        try:
            sheets = []
            nonempty = 0
            for sheet in book.worksheets:
                sheet_nonempty = 0
                for row in sheet.iter_rows(values_only=True):
                    sheet_nonempty += sum(value is not None for value in row)
                nonempty += sheet_nonempty
                sheets.append({"name": sheet.title, "rows": sheet.max_row, "columns": sheet.max_column, "nonempty_cells": sheet_nonempty})
            stats[path.name] = {"bytes": path.stat().st_size, "sheets": sheets, "nonempty_cells": nonempty}
        finally:
            book.close()
    return stats


def parse_worker_json(stdout: bytes) -> dict[str, Any]:
    try:
        return json.loads(stdout.decode("utf-8", errors="strict").strip().splitlines()[-1])
    except (UnicodeDecodeError, json.JSONDecodeError, IndexError) as error:
        raise RuntimeError(f"worker returned invalid JSON: {stdout.decode('utf-8', errors='replace')[-2000:]}") from error


def median_rust_stage(rounds: list[dict[str, Any]], section: str, key: str) -> float | None:
    values = [
        row["rust_performance_timings"][section][key]
        for row in rounds
        if row.get("rust_performance_timings")
        and row["rust_performance_timings"].get(section)
        and key in row["rust_performance_timings"][section]
    ]
    return round(statistics.median(values) / 1_000, 3) if values else None


def run_profile(binary: Path, python_source: Path, report_path: Path, rounds: int) -> dict[str, Any]:
    if report_path.exists():
        raise FileExistsError(f"refusing to overwrite benchmark report: {report_path}")
    if not binary.is_file():
        raise FileNotFoundError(binary)
    if not python_source.is_dir():
        raise FileNotFoundError(python_source)
    temporary_workspace = PreserveFailedTemporaryDirectory(prefix="audit-s1f1-full-chain-")
    with temporary_workspace as temp_name:
        root = Path(temp_name)
        inputs = make_fixture(root)
        fixed_inputs = [*sorted(inputs["source_dir"].glob("*.xlsx")), inputs["template"], inputs["external"], inputs["config"]]
        hashes_before = {str(path.relative_to(root)): sha256(path) for path in fixed_inputs}
        fixture_stats = create_stats(inputs)
        rounds_data: list[dict[str, Any]] = []
        script = Path(__file__).resolve()

        for index in range(rounds):
            run_root = root / f"round-{index + 1}"
            run_root.mkdir()
            python_config = run_root / "python-history-config.xlsx"
            rust_config = run_root / "rust-history-config.xlsx"
            shutil.copy2(inputs["config"], python_config)
            shutil.copy2(inputs["config"], rust_config)
            python_output = run_root / "python-output"
            rust_output = run_root / "rust-output"
            manifest_path = run_root / "rust-s1-manifest.json"
            manifest_path.write_text(json.dumps({
                "s1": {
                    "sourceDirectory": str(inputs["source_dir"]),
                    "templateDirectory": str(inputs["template_dir"]),
                    "mappingConfig": str(rust_config),
                    "externalData": str(inputs["external"]),
                }
            }, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

            python_command = [
                sys.executable, str(script), "--worker", "python",
                str(inputs["source_dir"]), str(inputs["template_dir"]), str(inputs["external"]),
                str(python_config), str(python_output), str(python_source),
            ]
            rust_command = [
                str(binary), "acceptance-run", "--manifest", str(manifest_path),
                "--output-dir", str(rust_output),
            ]
            rust_env = os.environ.copy()
            rust_env["AUDIT_STATE_DIR"] = str(run_root / "isolated-audit-state")
            if index % 2 == 0:
                python_memory, python_stdout, python_stderr = measure_process(python_command)
                rust_memory, rust_stdout, rust_stderr = measure_process(rust_command, environment=rust_env)
                order = "Python-first"
            else:
                rust_memory, rust_stdout, rust_stderr = measure_process(rust_command, environment=rust_env)
                python_memory, python_stdout, python_stderr = measure_process(python_command)
                order = "Rust-first"

            if python_memory["return_code"] != 0:
                raise RuntimeError(
                    f"Python worker failed (rc={python_memory['return_code']}):\n"
                    f"{python_stderr.decode('utf-8', errors='replace')[-5000:]}"
                )
            python_result = parse_worker_json(python_stdout)
            if python_result.get("status") != "completed" or not python_result.get("summary_path"):
                raise RuntimeError(f"Python S1-F1 flow failed: {python_result!r}\n{python_stderr.decode('utf-8', errors='replace')[-5000:]}")
            rust_report_path = rust_output / "acceptance-report.json"
            if not rust_report_path.is_file():
                raise RuntimeError(
                    f"Rust acceptance report missing (rc={rust_memory['return_code']}): {rust_report_path}\n"
                    f"{rust_stderr.decode('utf-8', errors='replace')[-5000:]}"
                )
            rust_report = json.loads(rust_report_path.read_text(encoding="utf-8"))
            if rust_report.get("status") != "completed" or not rust_report.get("steps"):
                raise RuntimeError(
                    "Rust S1-F1 flow failed: "
                    + json.dumps({
                        "return_code": rust_memory["return_code"],
                        "status": rust_report.get("status"),
                        "steps": rust_report.get("steps"),
                        "input_files": rust_report.get("input_files"),
                    }, ensure_ascii=False, indent=2)
                )
            rust_step = rust_report["steps"][0]
            rust_api_result = json.loads(rust_step.get("stdout", "{}"))
            rust_preflight = rust_api_result.get("preflight", {})
            rust_timing_sections = rust_preflight.get("performanceTimings", {})
            rust_timed_us = sum(
                rust_timing_sections.get(section, {}).get(key, 0)
                for section, key in (
                    ("request", "requestPreparationTotalUs"),
                    ("preflight", "totalUs"),
                    ("workflow", "workflowAdapterTotalUs"),
                )
            )
            python_summary = Path(python_result["summary_path"])
            rust_summary = find_rust_summary(rust_output)
            result_comparison = compare_result_workbooks(python_summary, rust_summary)
            formula_comparison = formula_copy_check(python_output, rust_output)
            python_excel_open = verify_excel_can_open([
                find_audit_copy(python_output, source_stem)
                for source_stem in ("甲银行", "乙银行")
            ])
            rust_excel_open = verify_excel_can_open([
                find_audit_copy(rust_output, source_stem)
                for source_stem in ("甲银行", "乙银行")
            ])
            python_config_hash_after = sha256(python_config)
            rust_config_hash_after = sha256(rust_config)
            source_hashes_after = {str(path.relative_to(root)): sha256(path) for path in fixed_inputs if path != inputs["config"]}

            rounds_data.append({
                "round": index + 1,
                "order": order,
                "python_business_elapsed_ms": python_result["elapsed_ms"],
                "rust_business_elapsed_ms": rust_step["elapsed_ms"],
                "python_worker_plus_excel": python_memory,
                "rust_worker_plus_excel": rust_memory,
                "python_engine_performance_lines": python_result.get("performance_lines", []),
                "rust_step_status": rust_step.get("status"),
                "rust_step_key_counts": rust_step.get("key_counts", {}),
                "rust_performance_timings": rust_preflight.get("performanceTimings"),
                "rust_acceptance_unclassified_us": max(
                    0,
                    rust_step.get("elapsed_ms", 0) * 1_000 - rust_timed_us,
                ),
                "summary_workbook_comparison": result_comparison,
                "formula_copy_and_cache_comparison": formula_comparison,
                "excel_com_open_check": {
                    "python": python_excel_open,
                    "rust": rust_excel_open,
                },
                "source_template_external_inputs_unchanged": source_hashes_after == {key: value for key, value in hashes_before.items() if key != str(inputs["config"].relative_to(root))},
                "python_history_config_changed": python_config_hash_after != hashes_before[str(inputs["config"].relative_to(root))],
                "rust_history_config_changed": rust_config_hash_after != hashes_before[str(inputs["config"].relative_to(root))],
                "python_summary_path": str(python_summary.relative_to(root)),
                "rust_summary_path": str(rust_summary.relative_to(root)),
            })

        hashes_after = {str(path.relative_to(root)): sha256(path) for path in fixed_inputs}
        python_times = [row["python_business_elapsed_ms"] for row in rounds_data]
        rust_times = [row["rust_business_elapsed_ms"] for row in rounds_data]
        python_memory = [row["python_worker_plus_excel"]["peak_sampled_worker_plus_new_excel_working_set_bytes"] for row in rounds_data]
        rust_memory = [row["rust_worker_plus_excel"]["peak_sampled_worker_plus_new_excel_working_set_bytes"] for row in rounds_data]
        output_checks_passed = all(
            row["summary_workbook_comparison"]["passed"]
            and row["formula_copy_and_cache_comparison"]["formula_text_equal"]
            and row["formula_copy_and_cache_comparison"]["formula_cache_equal"]
            and row["formula_copy_and_cache_comparison"]["formula_cache_matches_independent_expected_values"]
            and row["formula_copy_and_cache_comparison"]["referenced_external_sheet_copied_and_unreferenced_sheet_omitted"]
            and row["formula_copy_and_cache_comparison"]["structured_rule_sheet_content_equal"]
            and row["formula_copy_and_cache_comparison"]["structured_rule_sheet_comments_equal"]
            and row["formula_copy_and_cache_comparison"]["worksheet_names_and_order_equal"]
            and row["excel_com_open_check"]["python"]["opened_all"]
            and row["excel_com_open_check"]["python"]["owned_excel_process_exited"]
            and row["excel_com_open_check"]["rust"]["opened_all"]
            and row["excel_com_open_check"]["rust"]["owned_excel_process_exited"]
            for row in rounds_data
        )
        if not output_checks_passed:
            temporary_workspace.preserve()
        cleanup_passed = all(
            not row[engine]["remaining_new_excel_process_ids"]
            for row in rounds_data for engine in ("python_worker_plus_excel", "rust_worker_plus_excel")
        )
        report = {
            "schema_version": 1,
            "generated_at_utc": datetime.now(timezone.utc).isoformat(),
            "git_revision": git_revision(),
            "feature": "S1-F1",
            "measurement": "full end-to-end workflow; COM lifecycle is included in total; Rust stage timers use microsecond resolution from the same GUI worker path; nested substage fields are labeled explicitly",
            "platform": platform.platform(),
            "python_version": platform.python_version(),
            "python_source": str(python_source),
            "rust_binary": str(binary),
            "rust_binary_sha256": sha256(binary),
            "fixture": {
                "kind": "shared synthetic fixture from the read-only Python golden audit flow",
                "period": PERIOD,
                "source_workbooks": 3,
                "template_workbooks": 1,
                "external_workbooks": 1,
                "configuration_workbooks": 1,
                "expected_accepted_sources": 2,
                "expected_rejected_sources": 1,
                "expected_formula_issues": 2,
                "expected_conditional_format_issues": 1,
                "expected_rule_sheet_comments_per_audit_copy": 1,
                "workbook_stats": fixture_stats,
                "input_sha256_before": hashes_before,
                "input_sha256_after": hashes_after,
            },
            "method": {
                "rounds": rounds,
                "round_order": "alternates Python-first and Rust-first",
                "elapsed_metric": "Python run_dag_flow internal elapsed; Rust acceptance-run S1-F01 step elapsed (includes the acceptance adapter's 100ms state polling granularity)",
                "wall_metric": "fresh Python worker / Rust CLI launch until the worker and all newly created Excel processes exit; includes runtime startup and process shutdown",
                "memory_metric": "5ms sampled RSS sum over worker process tree plus only newly created EXCEL.EXE; pre-existing Excel PIDs excluded",
                "python_entry": "AuditService + workflow.runner.run_dag_flow(dag:汇总核查表校验), Windows COM pipeline",
                "rust_entry": "open-data-audit acceptance-run -> web_bridge::run_s1_f1_for_acceptance -> GUI S1-F1 worker",
                "output_comparison": "full final result workbook cell values and dimensions; formula text and calculated caches checked in accepted audit copies",
                "limitations": "synthetic golden fixture is small; process sampling can miss peaks shorter than 5ms; no style-by-style comparison",
            },
            "summary": {
                "python_business_elapsed_median_ms": round(statistics.median(python_times), 2),
                "rust_business_elapsed_median_ms": round(statistics.median(rust_times), 2),
                "python_to_rust_business_elapsed_ratio": round(statistics.median(python_times) / statistics.median(rust_times), 3) if statistics.median(rust_times) else None,
                "python_worker_wall_median_ms": round(statistics.median(row["python_worker_plus_excel"]["elapsed_wall_ms"] for row in rounds_data), 2),
                "rust_worker_wall_median_ms": round(statistics.median(row["rust_worker_plus_excel"]["elapsed_wall_ms"] for row in rounds_data), 2),
                "python_peak_working_set_median_mb": round(statistics.median(python_memory) / (1024 * 1024), 3),
                "python_peak_working_set_max_mb": round(max(python_memory) / (1024 * 1024), 3),
                "rust_peak_working_set_median_mb": round(statistics.median(rust_memory) / (1024 * 1024), 3),
                "rust_peak_working_set_max_mb": round(max(rust_memory) / (1024 * 1024), 3),
                "rust_request_stage_median_ms": {
                    name: median_rust_stage(rounds_data, "request", key)
                    for name, key in (
                        ("source_directory_scan", "sourceDirectoryScanUs"),
                        ("request_preparation_other", "requestPreparationOtherUs"),
                        ("request_preparation_total", "requestPreparationTotalUs"),
                    )
                },
                "rust_acceptance_unclassified_median_ms": round(
                    statistics.median(row["rust_acceptance_unclassified_us"] for row in rounds_data) / 1_000,
                    3,
                ),
                "rust_preflight_stage_median_ms": {
                    name: median_rust_stage(rounds_data, "preflight", key)
                    for name, key in (
                        ("template_profile_load", "templateProfileLoadUs"),
                        ("template_matching", "templateMatchingUs"),
                        ("template_preflight", "templatePreflightUs"),
                        ("source_structure_compare", "sourceStructureCompareUs"),
                        ("unclassified_preflight", "unclassifiedPreflightUs"),
                        ("total", "totalUs"),
                    )
                },
                "rust_workflow_stage_median_ms": {
                    name: median_rust_stage(rounds_data, "workflow", key)
                    for name, key in (
                        ("period_detection", "periodDetectionUs"),
                        ("workflow_setup_and_initial_manifest", "workflowSetupAndInitialManifestUs"),
                        ("copy_and_recalculation_preparation", "copyAndRecalculationPreparationTotalUs"),
                        ("excel_core_edit_and_output_validation", "excelCoreEditAndOutputValidationUs"),
                        ("copy_preparation_and_validation_other", "copyPreparationAndValidationOtherUs"),
                        ("external_excel_recalculation", "externalExcelRecalculationUs"),
                        ("issue_extraction_total", "issueExtractionTotalUs"),
                        ("formula_and_validation_issue_extraction", "formulaAndValidationIssueExtractionUs"),
                        ("conditional_format_issue_extraction", "conditionalFormatIssueExtractionUs"),
                        ("issue_extraction_other", "issueExtractionOtherUs"),
                        ("result_merge_history_and_workbook_write", "resultMergeHistoryAndWorkbookWriteUs"),
                        ("other_workflow_and_recovery_persistence", "otherWorkflowAndRecoveryPersistenceUs"),
                        ("workflow_total", "workflowTotalUs"),
                        ("recalculation_session_initialization", "recalculationSessionInitializationUs"),
                        ("recalculation_session_shutdown", "recalculationSessionShutdownUs"),
                        ("workflow_adapter_total", "workflowAdapterTotalUs"),
                    )
                },
                "full_output_and_formula_cache_checks_passed": output_checks_passed,
                "all_master_inputs_unchanged": hashes_before == hashes_after,
                "all_new_excel_processes_exited": cleanup_passed,
            },
            "rounds": rounds_data,
        }
    report_path.parent.mkdir(parents=True, exist_ok=True)
    with report_path.open("x", encoding="utf-8", newline="\n") as stream:
        stream.write(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    return report


def main() -> int:
    if len(sys.argv) > 1 and sys.argv[1:3] == ["--worker", "python"]:
        if len(sys.argv) != 9:
            raise SystemExit("usage: --worker python SOURCE_DIR TEMPLATE_DIR EXTERNAL CONFIG OUTPUT PYTHON_SOURCE")
        return python_worker(*(Path(value) for value in sys.argv[3:9]))

    parser = argparse.ArgumentParser()
    parser.add_argument("--audit-binary", type=Path, default=DEFAULT_BINARY)
    parser.add_argument("--python-source", type=Path, default=PYTHON_SOURCE)
    parser.add_argument("--report", type=Path, default=DEFAULT_REPORT)
    parser.add_argument("--rounds", type=int, default=3)
    args = parser.parse_args()
    if args.rounds < 1:
        parser.error("--rounds must be positive")
    report = run_profile(args.audit_binary, args.python_source, args.report, args.rounds)
    print(json.dumps(report["summary"], ensure_ascii=False, indent=2))
    return 0 if report["summary"]["full_output_and_formula_cache_checks_passed"] else 2


if __name__ == "__main__":
    raise SystemExit(main())
