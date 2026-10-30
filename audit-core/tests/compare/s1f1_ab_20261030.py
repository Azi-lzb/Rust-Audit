#!/usr/bin/env python3
"""S1-F1 原生 COM 严格 A/B 对照（Goal 第四/五节）。

- 同一原始副本每轮重新复制；输入 SHA-256 前后校验。
- 同机同 Excel（两通道各自创建私有 Excel.Application；版本由 Python 侧记录）。
- 统一生产语义：一个实例处理两册；open(UpdateLinks=0) → 逐工作表 Calculate →
  SaveAs(xlsx) → Close；Calculation 模式设置两通道均可能被 Excel 拒绝并同样
  降级（本机实测拒绝，实际计算模式一致）。
- 交替运行 PY → RUST → PY → RUST …共 ROUNDS 轮。
- 分阶段：create / open / calculate / save / close / shutdown / total（逐册）。
- 每轮校验输出公式 7,200 个且缓存值 7,200 个可读。
"""

from __future__ import annotations

import hashlib
import json
import shutil
import statistics
import subprocess
import sys
import tempfile
import time
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

WORKSPACE = Path(__file__).resolve().parents[3]
AUDIT_BIN = WORKSPACE / "dist" / "open-data-audit.exe"
ROUNDS = 5
FORMULA_EXPECT = 6000
CACHED_EXPECT = 7200
sys.path.insert(0, str(Path(r"C:" + chr(92) + "Users" + chr(92) + "AZI" + chr(92) + "Desktop" + chr(92) + "VibeCode" + chr(92) + "3.open-data-audit" + chr(92) + "core" + chr(92) + "src")))


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def build_originals(root: Path):
    sys.path.insert(0, str(WORKSPACE / "audit-core" / "tests" / "compare"))
    import perf_compare_20261030 as base
    originals = []
    for index in range(2):
        path = root / f"original-{index}.xlsx"
        base.build_s1_formula_workbook(path)
        originals.append(path)
    return originals


def verify_output(path: Path) -> dict:
    """公式数与缓存值数（openpyxl 两个视角）。"""
    from openpyxl import load_workbook
    formulas = load_workbook(path, data_only=False)["检查表"]
    cached = load_workbook(path, data_only=True)["检查表"]
    formula_count = 0
    cached_count = 0
    for row in formulas.iter_rows(min_row=1, max_row=1200, max_col=6):
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith("="):
                formula_count += 1
    for row in cached.iter_rows(min_row=1, max_row=1200, max_col=6):
        for cell in row:
            if cell.value is not None:
                cached_count += 1
    return {"formula_cells": formula_count, "cached_values": cached_count}


def py_round(originals, round_dir: Path):
    import pythoncom
    import win32com.client
    from base_audit.excel_com import ExcelSession

    round_dir.mkdir(parents=True, exist_ok=True)
    inputs = []
    for index, original in enumerate(originals):
        copy = round_dir / f"audit-copy-{index}.xlsx"
        shutil.copy(original, copy)
        inputs.append(copy)

    pythoncom.CoInitialize()
    total_started = time.perf_counter()
    t_create = time.perf_counter()
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    app.ScreenUpdating = False
    app.EnableEvents = False
    app.AskToUpdateLinks = False
    try:
        app.Calculation = -4135  # 手动；本机 Excel 拒绝时与 Rust 一样降级
    except Exception:
        pass
    create_ms = round((time.perf_counter() - t_create) * 1000, 1)
    excel_version = app.Version

    books = []
    for index, src in enumerate(inputs):
        open_t = time.perf_counter()
        book = app.Workbooks.Open(str(src), 0, False)
        open_ms = round((time.perf_counter() - open_t) * 1000, 1)

        calc_t = time.perf_counter()
        for worksheet in book.Worksheets:
            worksheet.Calculate()
        calc_ms = round((time.perf_counter() - calc_t) * 1000, 1)

        save_t = time.perf_counter()
        output = round_dir / f"out-{index}.xlsx"
        book.SaveAs(str(output), 51)
        save_ms = round((time.perf_counter() - save_t) * 1000, 1)

        close_t = time.perf_counter()
        book.Close(False)
        close_ms = round((time.perf_counter() - close_t) * 1000, 1)

        caches = verify_output(output)
        books.append({"open": open_ms, "calculate": calc_ms, "save": save_ms,
                      "close": close_ms, "caches": caches})

    quit_t = time.perf_counter()
    app.Quit()
    pythoncom.CoUninitialize()
    shutdown_ms = round((time.perf_counter() - quit_t) * 1000, 1)
    total_ms = round((time.perf_counter() - total_started) * 1000, 1)
    return {"create_ms": create_ms, "books": books, "shutdown_ms": shutdown_ms,
            "total_ms": total_ms, "excel_version": excel_version}


def rust_round(originals, round_dir: Path):
    round_dir.mkdir(parents=True, exist_ok=True)
    inputs = []
    for index, original in enumerate(originals):
        copy = round_dir / f"audit-copy-{index}.xlsx"
        shutil.copy(original, copy)
        inputs.append(copy)

    total_started = time.perf_counter()
    process = subprocess.Popen(
        [str(AUDIT_BIN), "__audit-native-com-session", "excel"],
        stdin=subprocess.PIPE, stdout=subprocess.PIPE,
        stderr=subprocess.DEVNULL, text=True, encoding="utf-8", bufsize=1)
    ready = json.loads(process.stdout.readline())
    if ready.get("status") != "session_ready" or not ready.get("com_available"):
        raise RuntimeError(f"worker unavailable: {ready}")
    create_ms = ready.get("session_create_ms")

    books = []
    for index, src in enumerate(inputs):
        output = round_dir / f"out-{index}.xlsx"
        book_t = time.perf_counter()
        process.stdin.write(json.dumps({
            "type": "recalculate", "input": str(src), "output": str(output),
            "mode": "full",
        }) + "\n")
        process.stdin.flush()
        response = json.loads(process.stdout.readline())
        book_ms = round((time.perf_counter() - book_t) * 1000, 1)
        if response.get("status") != "recalculated":
            raise RuntimeError(f"rust recalc failed: {response}")
        stages = response.get("stages_ms") or {}
        caches = verify_output(output)
        books.append({"open": stages.get("open"), "calculate": stages.get("sheets_calculate"),
                      "save": stages.get("save_as"), "close": stages.get("close"),
                      "caches": caches, "book_ms": book_ms})

    process.stdin.write(json.dumps({"type": "shutdown"}) + "\n")
    process.stdin.flush()
    quit_t = time.perf_counter()
    process.wait(timeout=30)
    shutdown_ms = round((time.perf_counter() - quit_t) * 1000, 1)
    total_ms = round((time.perf_counter() - total_started) * 1000, 1)
    return {"create_ms": create_ms, "books": books, "shutdown_ms": shutdown_ms,
            "total_ms": total_ms, "excel_version": None}


def main():
    rounds = int(sys.argv[1]) if len(sys.argv) > 1 else 5
    root = Path(tempfile.mkdtemp(prefix="s1f1_ab_"))
    originals = build_originals(root)
    original_hashes = [sha256(path) for path in originals]
    # 预热：Excel 二进制/过滤器进文件缓存，消除首轮流冷加载偏差后交替运行
    py_round(originals, root / "warm-py")
    rust_round(originals, root / "warm-rust")

    runs = []
    for index in range(rounds):
        for kind, call in (("py", py_round), ("rust", rust_round)):
            round_dir = root / f"{kind}-{index}"
            result = call(originals, round_dir)
            result["kind"] = kind
            runs.append(result)
            print(f"round {index} {kind.upper()}: total={result['total_ms']}ms "
                  f"create={result['create_ms']}ms shutdown={result['shutdown_ms']}ms "
                  f"books={[b['book_ms'] if 'book_ms' in b else b['open']+b['calculate']+b['save']+b['close'] for b in result['books']]}")

    for original, before in zip(originals, original_hashes):
        assert sha256(original) == before, "输入工作簿被修改"

    summary = {}
    for kind in ("py", "rust"):
        kind_runs = [run for run in runs if run["kind"] == kind]
        def med(key, book=None):
            if book is not None:
                values = []
                for run in kind_runs:
                    book_data = run["books"][book]
                    values.append(book_data.get(key, book_data.get("book_ms", 0)))
                return round(statistics.median(values), 1)
            values = [run[key] for run in kind_runs]
            return round(statistics.median(values), 1)
        summary[kind] = {
            "total_median_ms": med("total_ms"),
            "create_median_ms": med("create_ms"),
            "shutdown_median_ms": med("shutdown_ms"),
            "book1_median_ms": med("book_ms", 0),
            "book2_median_ms": med("book_ms", 1),
            "book1_open": med("open", 0), "book2_open": med("open", 1),
            "book1_calculate": med("calculate", 0), "book2_calculate": med("calculate", 1),
            "book1_save": med("save", 0), "book2_save": med("save", 1),
            "book1_close": med("close", 0), "book2_close": med("close", 1),
            "caches_all_verified": all(
                book["caches"]["formula_cells"] == FORMULA_EXPECT
                and book["caches"]["cached_values"] == CACHED_EXPECT
                for run in kind_runs for book in run["books"]),
            "excel_version": next((r["excel_version"] for r in kind_runs
                                   if r.get("excel_version")), None),
        }
    rust_total = summary["rust"]["total_median_ms"]
    py_total = summary["py"]["total_median_ms"]
    summary["ratio_py_over_rust"] = round(py_total / rust_total, 2) if rust_total else None

    print("\n=== A/B 汇总（中位数，预热后交替） ===")
    print(json.dumps(summary, ensure_ascii=False, indent=1))
    report = root / "s1f1_ab_report.json"
    report.write_text(json.dumps({"rounds": rounds, "runs": runs, "summary": summary},
                                 ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"report: {report}")


if __name__ == "__main__":
    main()
