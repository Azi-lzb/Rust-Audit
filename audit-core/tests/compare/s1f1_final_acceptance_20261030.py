#!/usr/bin/env python3
"""Final S1-F1 native COM cold-start and batch-session acceptance.

Uses only generated workbooks. Python uses DispatchEx in the already-running
benchmark host; Rust uses one isolated Audit worker per batch. Each engine is
measured from COM initialization/worker spawn through output-cache verification
and confirmed shutdown. The 2-book case runs five interleaved rounds; 10- and
33-book cases run twice each. One 2-book round places a broken workbook between
two valid files to verify that a single failure does not poison the session.
"""

from __future__ import annotations

import argparse
import gc
import hashlib
import json
import os
import shutil
import statistics
import subprocess
import sys
import tempfile
import time
from pathlib import Path
from typing import Any

import psutil

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

WORKSPACE = Path(__file__).resolve().parents[3]
DEFAULT_BIN = Path(r"F:\RustBuild\audit\release\open-data-audit.exe")
ROWS = 1200
FORMULAS_PER_BOOK = 6000
CACHED_VALUES_PER_BOOK = 7200
FAILURE_TOKEN = b"synthetic invalid xlsx; intentional continue-after-failure case"
FILETIME_EPOCH_SECONDS = 11_644_473_600


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


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


def wait_for_process_exit(process_id: int, expected_created: float, timeout: float) -> bool:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        try:
            process = psutil.Process(process_id)
            if abs(process.create_time() - expected_created) > 0.05:
                return True  # PID was reused; the original process is gone.
            if not process.is_running() or process.status() == psutil.STATUS_ZOMBIE:
                return True
        except psutil.NoSuchProcess:
            return True
        except (psutil.AccessDenied, psutil.ZombieProcess):
            return False
        time.sleep(0.05)
    return False


def terminate_exact_excel(process_id: int, expected_created: float) -> bool:
    try:
        process = psutil.Process(process_id)
        if abs(process.create_time() - expected_created) > 0.05:
            return False
        process.terminate()
        process.wait(timeout=5)
        return True
    except psutil.NoSuchProcess:
        return True
    except (psutil.AccessDenied, psutil.TimeoutExpired, psutil.ZombieProcess):
        return False


def make_originals(root: Path) -> list[Path]:
    compare_dir = WORKSPACE / "audit-core" / "tests" / "compare"
    sys.path.insert(0, str(compare_dir))
    import perf_compare_20261030 as base
    from openpyxl import load_workbook

    originals: list[Path] = []
    for index in range(2):
        path = root / f"synthetic-original-{index}.xlsx"
        base.build_s1_formula_workbook(path, rows=ROWS)
        # The accepted sample scope is 5 formula columns x 1,200 rows plus
        # 1,200 key inputs = 6,000 formulas and 7,200 populated cached cells.
        book = load_workbook(path, data_only=False)
        try:
            sheet = book["检查表"]
            for row in range(1, ROWS + 1):
                sheet.cell(row=row, column=7).value = None
            book.save(path)
        finally:
            book.close()
        originals.append(path)
    return originals


def verify_output(path: Path) -> dict[str, Any]:
    from openpyxl import load_workbook

    formula_book = load_workbook(path, read_only=True, data_only=False)
    formula_lines: list[str] = []
    try:
        formula_sheet = formula_book["检查表"]
        formulas = 0
        for row_number, row in enumerate(
            formula_sheet.iter_rows(min_row=1, max_row=ROWS, min_col=1, max_col=6), 1
        ):
            if row[0].value != f"K{row_number:05d}":
                raise AssertionError(f"{path.name}: unexpected key at A{row_number}: {row[0].value!r}")
            for column_index, cell in enumerate(row[1:], 2):
                expected_lookup_column = 3 if column_index in (2, 4, 6) else 2
                expected_formula = (
                    f"=VLOOKUP($A{row_number},参照表!$A:$C,"
                    f"{expected_lookup_column},FALSE)+1"
                )
                normalized_formula = (
                    cell.value.replace("'参照表'!", "参照表!")
                    if isinstance(cell.value, str)
                    else cell.value
                )
                if normalized_formula != expected_formula:
                    raise AssertionError(
                        f"{path.name}: formula mismatch at {cell.coordinate}: {cell.value!r}"
                    )
                formula_lines.append(f"{cell.coordinate}={cell.value}")
                formulas += 1
        if formulas != FORMULAS_PER_BOOK:
            raise AssertionError(f"{path.name}: expected {FORMULAS_PER_BOOK} formulas, found {formulas}")
    finally:
        formula_book.close()

    value_book = load_workbook(path, read_only=True, data_only=True)
    cache_lines: list[str] = []
    try:
        value_sheet = value_book["检查表"]
        cached_values = 0
        verified_results = 0
        for row_number, row in enumerate(
            value_sheet.iter_rows(min_row=1, max_row=ROWS, min_col=1, max_col=6), 1
        ):
            if row[0].value != f"K{row_number:05d}":
                raise AssertionError(f"{path.name}: key cache missing at A{row_number}")
            cached_values += 1
            for column_index, cell in enumerate(row[1:], 2):
                lookup_column = 3 if column_index in (2, 4, 6) else 2
                expected = (200.0 if lookup_column == 3 else 100.0) + row_number + 1
                if not isinstance(cell.value, (int, float)) or abs(float(cell.value) - expected) > 1e-9:
                    raise AssertionError(
                        f"{path.name}: cache mismatch at {cell.coordinate}: "
                        f"expected {expected}, got {cell.value!r}"
                    )
                cached_values += 1
                verified_results += 1
                cache_lines.append(f"{cell.coordinate}={float(cell.value):.12g}")
        if cached_values != CACHED_VALUES_PER_BOOK:
            raise AssertionError(
                f"{path.name}: expected {CACHED_VALUES_PER_BOOK} cached cells, found {cached_values}"
            )
        if verified_results != FORMULAS_PER_BOOK:
            raise AssertionError(
                f"{path.name}: expected {FORMULAS_PER_BOOK} verified formula results, "
                f"found {verified_results}"
            )
    finally:
        value_book.close()

    return {
        "formula_cells": formulas,
        "cached_values": cached_values,
        "formula_results_verified": verified_results,
        "formula_hash": hashlib.sha256("\n".join(formula_lines).encode()).hexdigest(),
        "result_hash": hashlib.sha256("\n".join(cache_lines).encode()).hexdigest(),
    }


def copy_batch_inputs(originals: list[Path], directory: Path, count: int, inject_failure: bool):
    directory.mkdir(parents=True, exist_ok=True)
    valid_inputs: list[Path] = []
    for index in range(count):
        source = originals[index % len(originals)]
        target = directory / f"input-{index:02d}.xlsx"
        shutil.copy2(source, target)
        valid_inputs.append(target)

    entries: list[dict[str, Any]] = []
    if inject_failure:
        bad_path = directory / "intentional-invalid.xlsx"
        bad_path.write_bytes(FAILURE_TOKEN)
        entries.append({"path": valid_inputs[0], "valid_index": 0})
        entries.append({"path": bad_path, "valid_index": None})
        entries.extend({"path": path, "valid_index": i} for i, path in enumerate(valid_inputs[1:], 1))
    else:
        entries.extend({"path": path, "valid_index": i} for i, path in enumerate(valid_inputs))
    hashes = {str(entry["path"]): sha256(entry["path"]) for entry in entries}
    return entries, hashes


def python_batch(originals: list[Path], batch_dir: Path, count: int, inject_failure: bool):
    import pythoncom
    import win32com.client

    entries, input_hashes = copy_batch_inputs(originals, batch_dir, count, inject_failure)
    baseline = excel_inventory()
    total_started = time.perf_counter()
    timings: dict[str, Any] = {"com_initialize_ms": None, "application_properties_ms": {}}
    books: list[dict[str, Any]] = []
    app = None
    workbook = None
    worksheet = None
    process_id = None
    process_created = None
    forced_excel_termination = False
    com_initialized = False
    app_quit_succeeded = False
    activation_started_wall = None
    try:
        started = time.perf_counter()
        pythoncom.CoInitialize()
        com_initialized = True
        timings["com_initialize_ms"] = round((time.perf_counter() - started) * 1000, 2)

        activation_started_wall = time.time()
        started = time.perf_counter()
        app = win32com.client.DispatchEx("Excel.Application")
        activation_ms = (time.perf_counter() - started) * 1000
        timings["excel_application_com_activation_ms"] = round(activation_ms, 2)

        properties = [
            ("Visible", False),
            ("DisplayAlerts", False),
            ("ScreenUpdating", False),
            ("EnableEvents", False),
            ("AskToUpdateLinks", False),
        ]
        for name, value in properties:
            started = time.perf_counter()
            setattr(app, name, value)
            timings["application_properties_ms"][name] = round((time.perf_counter() - started) * 1000, 2)
        started = time.perf_counter()
        try:
            app.Calculation = -4135
            timings["calculation_manual_succeeded"] = True
        except Exception:
            timings["calculation_manual_succeeded"] = False
        timings["application_properties_ms"]["Calculation"] = round(
            (time.perf_counter() - started) * 1000, 2
        )
        timings["application_properties_total_ms"] = round(
            sum(timings["application_properties_ms"].values()), 2
        )

        pid_lookup_started = time.perf_counter()
        process_id = int(app.Hwnd)
        # Excel.Application.Hwnd is a window handle; use GetWindowThreadProcessId.
        import win32process

        process_id = int(win32process.GetWindowThreadProcessId(process_id)[1])
        process = psutil.Process(process_id)
        process_created = process.create_time()
        timings["excel_hwnd_pid_lookup_ms"] = round((time.perf_counter() - pid_lookup_started) * 1000, 2)
        launch_offset_ms = (process_created - activation_started_wall) * 1000
        timings["excel_process_launch_after_activation_start_ms_estimate"] = round(launch_offset_ms, 2)
        timings["excel_com_ready_after_process_creation_ms_estimate"] = round(
            activation_ms - launch_offset_ms, 2
        )
        identity_check_started = time.perf_counter()
        if process_id in baseline:
            raise AssertionError(f"Python DispatchEx attached to pre-existing Excel PID {process_id}")
        new_excel = set(excel_inventory()) - set(baseline)
        if process_id not in new_excel:
            raise AssertionError(f"DispatchEx PID {process_id} was not newly created")
        timings["ownership_assertion_ms"] = round(
            (time.perf_counter() - identity_check_started) * 1000, 2
        )
        timings["new_excel_pids_at_ready"] = sorted(new_excel)

        for entry in entries:
            source: Path = entry["path"]
            valid_index = entry["valid_index"]
            record: dict[str, Any] = {
                "input": source.name,
                "status": "failed",
                "open_ms": None,
                "calculate_ms": None,
                "save_ms": None,
                "close_ms": 0.0,
                "cache_verify_ms": 0.0,
                "engine_process_id": process_id,
            }
            try:
                started = time.perf_counter()
                workbook = app.Workbooks.Open(str(source), 0, False)
                record["open_ms"] = round((time.perf_counter() - started) * 1000, 2)
                if valid_index is None:
                    record["unexpectedly_opened_invalid_input"] = True
                else:
                    started = time.perf_counter()
                    for worksheet in workbook.Worksheets:
                        worksheet.Calculate()
                    record["calculate_ms"] = round((time.perf_counter() - started) * 1000, 2)
                    output = batch_dir / f"output-{valid_index:02d}.xlsx"
                    started = time.perf_counter()
                    workbook.SaveAs(str(output), 51)
                    record["save_ms"] = round((time.perf_counter() - started) * 1000, 2)
                    started = time.perf_counter()
                    workbook.Close(False)
                    workbook = None
                    worksheet = None
                    record["close_ms"] = round((time.perf_counter() - started) * 1000, 2)
                    started = time.perf_counter()
                    record["verification"] = verify_output(output)
                    record["cache_verify_ms"] = round((time.perf_counter() - started) * 1000, 2)
                    record["output"] = str(output)
                    record["status"] = "success"
            except Exception as error:
                record["error"] = f"{type(error).__name__}: {error}"
                if workbook is not None:
                    started = time.perf_counter()
                    try:
                        workbook.Close(False)
                    except Exception:
                        pass
                    workbook = None
                    record["close_ms"] = round((time.perf_counter() - started) * 1000, 2)
                    worksheet = None
            books.append(record)

    finally:
        shutdown_started = time.perf_counter()
        quit_ms = 0.0
        release_ms = 0.0
        uninitialize_ms = 0.0
        if workbook is not None:
            try:
                workbook.Close(False)
            except Exception:
                pass
            workbook = None
        worksheet = None
        if app is not None:
            started = time.perf_counter()
            try:
                app.Quit()
                app_quit_succeeded = True
            except Exception:
                app_quit_succeeded = False
            quit_ms = round((time.perf_counter() - started) * 1000, 2)
            started = time.perf_counter()
            app = None
            gc.collect()
            release_ms = round((time.perf_counter() - started) * 1000, 2)
        if com_initialized:
            started = time.perf_counter()
            pythoncom.CoUninitialize()
            uninitialize_ms = round((time.perf_counter() - started) * 1000, 2)
        process_wait_started = time.perf_counter()
        process_exited = (
            process_id is None
            or process_created is None
            or wait_for_process_exit(process_id, process_created, 5)
        )
        if not process_exited and process_id is not None and process_created is not None:
            forced_excel_termination = terminate_exact_excel(process_id, process_created)
            process_exited = wait_for_process_exit(process_id, process_created, 5)
        process_wait_ms = round((time.perf_counter() - process_wait_started) * 1000, 2)
        timings["shutdown_ms"] = {
            "excel_quit_call_ms": quit_ms,
            "com_proxy_release_ms": release_ms,
            "com_uninitialize_ms": uninitialize_ms,
            "excel_process_exit_wait_ms": process_wait_ms,
            "total_ms": round((time.perf_counter() - shutdown_started) * 1000, 2),
            "excel_quit_succeeded": app_quit_succeeded,
            "excel_process_exited": process_exited,
            "forced_excel_termination": forced_excel_termination,
        }

    final_check_started = time.perf_counter()
    for path_text, before_hash in input_hashes.items():
        if sha256(Path(path_text)) != before_hash:
            raise AssertionError(f"Python modified input copy: {path_text}")
    after_excel = excel_inventory()
    timings["final_input_and_process_checks_ms"] = round(
        (time.perf_counter() - final_check_started) * 1000, 2
    )
    if process_id in after_excel:
        raise AssertionError(f"Python left its Excel PID {process_id} running")
    if forced_excel_termination:
        raise AssertionError(f"Python needed forced cleanup for Excel PID {process_id}")
    if not app_quit_succeeded or not timings["shutdown_ms"]["excel_process_exited"]:
        raise AssertionError(f"Python Excel shutdown was not clean: {timings['shutdown_ms']}")
    failures = [book for book in books if book["status"] == "failed"]
    successes = [book for book in books if book["status"] == "success"]
    if any(book.get("unexpectedly_opened_invalid_input") for book in books):
        raise AssertionError("Python unexpectedly opened the intentionally invalid workbook")
    if inject_failure:
        if len(failures) != 1 or len(successes) != count:
            raise AssertionError(f"Python did not continue after one bad workbook: {books}")
    elif failures or len(successes) != count:
        raise AssertionError(f"Python batch failed: {books}")

    stage_total = (
        timings["com_initialize_ms"]
        + timings.get("excel_application_com_activation_ms", 0)
        + timings.get("application_properties_total_ms", 0)
        + timings.get("excel_hwnd_pid_lookup_ms", 0)
        + timings.get("ownership_assertion_ms", 0)
        + sum(
            (book.get("open_ms") or 0)
            + (book.get("calculate_ms") or 0)
            + (book.get("save_ms") or 0)
            + book.get("close_ms", 0)
            + book.get("cache_verify_ms", 0)
            for book in books
        )
        + timings["shutdown_ms"]["total_ms"]
        + timings["final_input_and_process_checks_ms"]
    )
    total_ms = round((time.perf_counter() - total_started) * 1000, 2)
    return {
        "kind": "python",
        "requested_valid_books": count,
        "request_count": len(entries),
        "instance_creation_count": 1,
        "excel_process_id": process_id,
        "excel_process_created_unix": process_created,
        "excel_process_count_before": len(baseline),
        "excel_process_count_after": len(after_excel),
        "workbooks_succeeded": len(successes),
        "workbooks_failed": len(failures),
        "workbooks_skipped": 0,
        "books": books,
        "timings": timings,
        "total_ms": total_ms,
        "accounted_stage_sum_ms": round(stage_total, 2),
        "unaccounted_ms": round(total_ms - stage_total, 2),
        "input_hashes": input_hashes,
        "clean_shutdown": True,
    }


def rust_batch(
    originals: list[Path], batch_dir: Path, count: int, inject_failure: bool, audit_bin: Path
):
    entries, input_hashes = copy_batch_inputs(originals, batch_dir, count, inject_failure)
    baseline = excel_inventory()
    total_started = time.perf_counter()
    spawn_started = time.perf_counter()
    creation_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    process = subprocess.Popen(
        [str(audit_bin), "__audit-native-com-session", "excel"],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.DEVNULL,
        text=True,
        encoding="utf-8",
        bufsize=1,
        creationflags=creation_flags,
    )
    spawn_call_ms = round((time.perf_counter() - spawn_started) * 1000, 2)
    ready_wait_started = time.perf_counter()
    excel_pid = None
    excel_created = None
    worker_pid = process.pid
    records: list[dict[str, Any]] = []
    process_exited = False
    forced_excel_termination = False
    shutdown_report: dict[str, Any] | None = None
    try:
        line = process.stdout.readline()
        if not line:
            raise RuntimeError("Rust COM worker exited before session_ready")
        ready = json.loads(line)
        ready_wait_after_spawn_ms = round((time.perf_counter() - ready_wait_started) * 1000, 2)
        spawn_to_ready_ms = round((time.perf_counter() - total_started) * 1000, 2)
        if ready.get("status") != "session_ready" or not ready.get("com_available"):
            raise RuntimeError(f"Rust COM worker unavailable: {ready}")
        if ready.get("engine_process_owned") is not True:
            raise AssertionError(f"Rust worker did not own its Office process: {ready}")
        excel_pid = int(ready["engine_process_id"])
        created_ticks = int(ready["engine_process_created"])
        excel_created = created_ticks / 10_000_000 - FILETIME_EPOCH_SECONDS
        if excel_pid in baseline:
            raise AssertionError(f"Rust worker connected to pre-existing Excel PID {excel_pid}")
        new_excel = set(excel_inventory()) - set(baseline)
        if excel_pid not in new_excel:
            raise AssertionError(f"Rust COM Excel PID {excel_pid} was not newly created")
        initialization_ms = int(ready.get("session_create_ms", 0))
        ownership_check_started = time.perf_counter()
        worker_timing = {
            "spawn_call_ms": spawn_call_ms,
            "ready_wait_after_spawn_ms": ready_wait_after_spawn_ms,
            "spawn_to_ready_ms": spawn_to_ready_ms,
            "worker_bootstrap_and_ready_ipc_estimate_ms": max(0, spawn_to_ready_ms - initialization_ms),
            "session_create_ms": initialization_ms,
            "initialization_stages_ms": ready.get("initialization_stages_ms"),
        }
        worker_created_wall = psutil.Process(worker_pid).create_time()
        ready_wall = time.time()
        worker_timing["worker_process_create_to_ready_ms_estimate"] = round(
            (ready_wall - worker_created_wall) * 1000, 2
        )
        worker_timing["ownership_assertion_ms"] = round(
            (time.perf_counter() - ownership_check_started) * 1000, 2
        )
        worker_timing["new_excel_pids_at_ready"] = sorted(new_excel)

        for entry in entries:
            source: Path = entry["path"]
            output = batch_dir / f"output-{entry['valid_index']:02d}.xlsx" if entry["valid_index"] is not None else batch_dir / "invalid-output.xlsx"
            started = time.perf_counter()
            request = {
                "type": "recalculate",
                "input": str(source),
                "output": str(output),
                "mode": "full",
            }
            process.stdin.write(json.dumps(request, ensure_ascii=False) + "\n")
            process.stdin.flush()
            response_line = process.stdout.readline()
            if not response_line:
                raise RuntimeError(f"Rust worker exited during {source.name}")
            response = json.loads(response_line)
            request_ms = round((time.perf_counter() - started) * 1000, 2)
            record: dict[str, Any] = {
                "input": source.name,
                "engine_process_id": excel_pid,
                "request_roundtrip_ms": request_ms,
                "status": response.get("status"),
                "stages_ms": response.get("stages_ms"),
                "request_unattributed_ms_estimate": round(
                    request_ms - sum((response.get("stages_ms") or {}).values()), 2
                ),
            }
            if entry["valid_index"] is None:
                if response.get("status") != "failed":
                    record["unexpectedly_opened_invalid_input"] = True
                else:
                    record["error"] = response.get("error")
                record["cache_verify_ms"] = 0.0
                records.append(record)
                continue
            if response.get("status") != "recalculated":
                record["error"] = response.get("error")
                records.append(record)
                continue
            if int(response.get("engine_process_id", -1)) != excel_pid:
                raise AssertionError(f"Rust did not reuse Excel PID {excel_pid}: {response}")
            verify_started = time.perf_counter()
            record["verification"] = verify_output(output)
            record["cache_verify_ms"] = round((time.perf_counter() - verify_started) * 1000, 2)
            record["output"] = str(output)
            records.append(record)

        shutdown_started = time.perf_counter()
        process.stdin.write(json.dumps({"type": "shutdown"}) + "\n")
        process.stdin.flush()
        ack_started = time.perf_counter()
        shutdown_line = process.stdout.readline()
        shutdown_ack_wait_ms = round((time.perf_counter() - ack_started) * 1000, 2)
        if shutdown_line:
            shutdown_report = json.loads(shutdown_line)
        process.stdin.close()
        worker_wait_started = time.perf_counter()
        return_code = process.wait(timeout=15)
        worker_exit_wait_ms = round((time.perf_counter() - worker_wait_started) * 1000, 2)
        worker_exited = return_code == 0
        excel_wait_started = time.perf_counter()
        office_exited = wait_for_process_exit(excel_pid, excel_created, 5)
        if not office_exited:
            forced_excel_termination = terminate_exact_excel(excel_pid, excel_created)
            office_exited = wait_for_process_exit(excel_pid, excel_created, 5)
        excel_exit_wait_ms = round((time.perf_counter() - excel_wait_started) * 1000, 2)
        shutdown_ms = round((time.perf_counter() - shutdown_started) * 1000, 2)
        process_exited = True
        shutdown_timings = (shutdown_report or {}).get("shutdown_stages_ms") or {}
        shutdown_timings = {
            **shutdown_timings,
            "shutdown_ack_wait_ms": shutdown_ack_wait_ms,
            "worker_exit_wait_ms": worker_exit_wait_ms,
            "excel_process_exit_wait_ms": excel_exit_wait_ms,
            "driver_shutdown_total_ms": shutdown_ms,
            "excel_quit_succeeded": (shutdown_report or {}).get("excel_quit_succeeded"),
            "excel_process_exited": office_exited,
            "forced_excel_termination": forced_excel_termination,
            "worker_return_code": return_code,
            "worker_exited": worker_exited,
        }
        final_check_started = time.perf_counter()
        for path_text, before_hash in input_hashes.items():
            if sha256(Path(path_text)) != before_hash:
                raise AssertionError(f"Rust modified input copy: {path_text}")
        after_excel = excel_inventory()
        final_input_and_process_checks_ms = round(
            (time.perf_counter() - final_check_started) * 1000, 2
        )
        if excel_pid in after_excel:
            raise AssertionError(f"Rust left its Excel PID {excel_pid} running")
        if forced_excel_termination:
            raise AssertionError(f"Rust needed forced cleanup for Excel PID {excel_pid}")
        if not worker_exited or shutdown_timings["excel_quit_succeeded"] is not True or not office_exited:
            raise AssertionError(f"Rust COM shutdown was not clean: {shutdown_timings}")
        failures = [book for book in records if book["status"] == "failed"]
        successes = [book for book in records if book["status"] == "recalculated"]
        if any(book.get("unexpectedly_opened_invalid_input") for book in records):
            raise AssertionError("Rust unexpectedly opened the intentionally invalid workbook")
        if inject_failure:
            if len(failures) != 1 or len(successes) != count:
                raise AssertionError(f"Rust did not continue after one bad workbook: {records}")
        elif failures or len(successes) != count:
            raise AssertionError(f"Rust batch failed: {records}")

        total_ms = round((time.perf_counter() - total_started) * 1000, 2)
        stage_total = spawn_to_ready_ms + worker_timing["ownership_assertion_ms"] + sum(
            record["request_roundtrip_ms"] + record.get("cache_verify_ms", 0)
            for record in records
        ) + shutdown_ms + final_input_and_process_checks_ms
        return {
            "kind": "rust",
            "requested_valid_books": count,
            "request_count": len(entries),
            "instance_creation_count": 1,
            "worker_process_id": worker_pid,
            "excel_process_id": excel_pid,
            "excel_process_created_unix": excel_created,
            "excel_process_count_before": len(baseline),
            "excel_process_count_after": len(after_excel),
            "workbooks_succeeded": len(successes),
            "workbooks_failed": len(failures),
            "workbooks_skipped": 0,
            "books": records,
            "timings": {
                "worker": worker_timing,
                "shutdown": shutdown_timings,
                "final_input_and_process_checks_ms": final_input_and_process_checks_ms,
            },
            "total_ms": total_ms,
            "accounted_stage_sum_ms": round(stage_total, 2),
            "unaccounted_ms": round(total_ms - stage_total, 2),
            "input_hashes": input_hashes,
            "worker_exit_code": return_code,
            "worker_gone": process.poll() is not None,
            "clean_shutdown": True,
        }
    finally:
        if process.poll() is None:
            process.kill()
            process.wait(timeout=5)
        if not process_exited and excel_pid is not None and excel_created is not None:
            # Emergency cleanup is restricted to the new PID and its creation
            # timestamp from the worker's ready message.
            terminate_exact_excel(excel_pid, excel_created)


def median(values: list[float]) -> float | None:
    return round(statistics.median(values), 2) if values else None


def summarize(runs: list[dict[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for size in (2, 10, 33):
        result[str(size)] = {}
        for kind in ("python", "rust"):
            group = [run for run in runs if run["requested_valid_books"] == size and run["kind"] == kind]
            result[str(size)][kind] = {
                "batches": len(group),
                "total_median_ms": median([run["total_ms"] for run in group]),
                "average_ms_per_valid_book_median": median(
                    [run["total_ms"] / run["requested_valid_books"] for run in group]
                ),
                "excel_initialization_median_ms": median(
                    [
                        float(
                            run["timings"].get("worker", {}).get("session_create_ms")
                            or run["timings"].get("excel_application_com_activation_ms")
                            or 0
                        )
                        for run in group
                    ]
                ),
                "successes": sum(run["workbooks_succeeded"] for run in group),
                "failures": sum(run["workbooks_failed"] for run in group),
                "skipped": sum(run["workbooks_skipped"] for run in group),
                "excel_instances_created": sum(run["instance_creation_count"] for run in group),
                "forced_excel_terminations": sum(
                    bool(run["timings"].get("shutdown", run["timings"].get("shutdown_ms", {})).get("forced_excel_termination"))
                    for run in group
                ),
                "cache_outputs_verified": sum(
                    1
                    for run in group
                    for book in run["books"]
                    if book.get("status") in {"success", "recalculated"}
                    and book.get("verification", {}).get("formula_cells") == FORMULAS_PER_BOOK
                    and book.get("verification", {}).get("cached_values") == CACHED_VALUES_PER_BOOK
                    and book.get("verification", {}).get("formula_results_verified") == FORMULAS_PER_BOOK
                ),
            }
    return result


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--audit-bin", type=Path, default=Path(os.environ.get("AUDIT_BIN", DEFAULT_BIN)))
    parser.add_argument("--rounds-small", type=int, default=5)
    parser.add_argument("--rounds-large", type=int, default=2)
    parser.add_argument("--output-root", type=Path, default=Path(r"F:\RustBuild\audit"))
    args = parser.parse_args()
    audit_bin = args.audit_bin.resolve()
    if not audit_bin.is_file():
        raise FileNotFoundError(f"Rust Audit executable is missing: {audit_bin}")
    args.output_root.mkdir(parents=True, exist_ok=True)
    run_root = Path(tempfile.mkdtemp(prefix="s1f1-final-acceptance-20261030-", dir=args.output_root))
    originals = make_originals(run_root)
    original_hashes = {str(path): sha256(path) for path in originals}
    baseline_excel = excel_inventory()
    pythoncom = __import__("pythoncom")
    # Warm the two activation paths before recording. Warm-ups are discarded
    # and use one synthetic workbook, then assert both processes were cleaned.
    warm_source = originals[:1]
    warm_py = python_batch(warm_source, run_root / "warm-py", 1, False)
    warm_rust = rust_batch(warm_source, run_root / "warm-rust", 1, False, audit_bin)
    runs: list[dict[str, Any]] = []
    errors: list[dict[str, Any]] = []
    for size, rounds in ((2, args.rounds_small), (10, args.rounds_large), (33, args.rounds_large)):
        for round_index in range(rounds):
            inject_failure = size == 2 and round_index == 0
            for kind in ("python", "rust"):
                batch_dir = run_root / f"batch-{size}-{round_index}-{kind}"
                try:
                    if kind == "python":
                        result = python_batch(originals, batch_dir, size, inject_failure)
                    else:
                        result = rust_batch(originals, batch_dir, size, inject_failure, audit_bin)
                    result["round_index"] = round_index
                    result["failure_injected"] = inject_failure
                    runs.append(result)
                    print(
                        f"{kind.upper()} size={size} round={round_index + 1}/{rounds}: "
                        f"total={result['total_ms']}ms success={result['workbooks_succeeded']} "
                        f"failed={result['workbooks_failed']} ExcelPID={result['excel_process_id']}"
                    )
                except Exception as error:
                    item = {
                        "kind": kind,
                        "size": size,
                        "round_index": round_index,
                        "failure_injected": inject_failure,
                        "error": f"{type(error).__name__}: {error}",
                    }
                    errors.append(item)
                    print(f"ERROR {item}")

    paired_cache_differences = []
    for python_run in [run for run in runs if run["kind"] == "python"]:
        rust_run = next(
            (
                run
                for run in runs
                if run["kind"] == "rust"
                and run["requested_valid_books"] == python_run["requested_valid_books"]
                and run["round_index"] == python_run["round_index"]
            ),
            None,
        )
        if rust_run is None:
            continue
        py_results = [
            book.get("verification")
            for book in python_run["books"]
            if book.get("status") == "success"
        ]
        rust_results = [
            book.get("verification")
            for book in rust_run["books"]
            if book.get("status") == "recalculated"
        ]
        if len(py_results) != len(rust_results) or any(
            py.get("formula_hash") != rust.get("formula_hash")
            or py.get("result_hash") != rust.get("result_hash")
            for py, rust in zip(py_results, rust_results, strict=True)
        ):
            paired_cache_differences.append(
                {
                    "size": python_run["requested_valid_books"],
                    "round_index": python_run["round_index"],
                    "python": py_results,
                    "rust": rust_results,
                }
            )
    if paired_cache_differences:
        errors.append({"scope": "python_rust_cache_parity", "differences": paired_cache_differences})

    original_hashes_after = {str(path): sha256(path) for path in originals}
    if original_hashes_after != original_hashes:
        errors.append({"scope": "originals", "error": "synthetic source workbook hash changed"})
    remaining_excel = excel_inventory()
    remaining_original_ids = set(baseline_excel)
    leaked_owned = [
        pid
        for run in runs
        if (pid := run.get("excel_process_id")) is not None and pid in remaining_excel
    ]
    if leaked_owned:
        errors.append({"scope": "excel_cleanup", "pids": leaked_owned})
    report = {
        "task": "S1-F1 COM final performance and batch acceptance",
        "timestamp_local": time.strftime("%Y-%m-%d %H:%M:%S %z"),
        "audit_binary": str(audit_bin),
        "audit_binary_sha256": sha256(audit_bin),
        "python_version": sys.version,
        "pythoncom_initialized_in_existing_benchmark_host": True,
        "excel_private_activation": "Python win32com.client.DispatchEx / Rust CoCreateInstance(CLSCTX_LOCAL_SERVER)",
        "baseline_excel_processes": {str(pid): created for pid, created in sorted(baseline_excel.items())},
        "final_excel_processes": {str(pid): created for pid, created in sorted(remaining_excel.items())},
        "warmups_excluded": {"python": warm_py, "rust": warm_rust},
        "workbook_shape": {
            "rows": ROWS,
            "formula_cells_per_workbook": FORMULAS_PER_BOOK,
            "cached_values_per_workbook": CACHED_VALUES_PER_BOOK,
            "formula_type": "VLOOKUP against a same-workbook reference sheet",
            "formula_result_validation": "all 6,000 results compared to independent numeric expectations",
        },
        "original_workbook_hashes_before": original_hashes,
        "original_workbook_hashes_after": original_hashes_after,
        "runs": runs,
        "paired_cache_differences": paired_cache_differences,
        "summary": summarize(runs),
        "errors": errors,
        "remaining_excel_ids_from_preexisting_baseline": sorted(remaining_original_ids & set(remaining_excel)),
        "new_excel_processes_not_owned_by_this_harness": {
            str(pid): created
            for pid, created in remaining_excel.items()
            if pid not in remaining_original_ids
            and pid not in {run.get("excel_process_id") for run in runs}
        },
    }
    report_path = run_root / "s1f1_final_acceptance.json"
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"report={report_path}")
    print(json.dumps(report["summary"], ensure_ascii=False, indent=2))
    print(f"errors={len(errors)}")
    return 1 if errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
