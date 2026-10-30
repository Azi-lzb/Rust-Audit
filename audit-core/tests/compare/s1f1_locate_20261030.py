#!/usr/bin/env python3
"""S1-F1 / S3-F1 性能热点定位（Goal 第二/三阶段）。

S1-F1：同一批 2 册 7,200 公式工作簿，分别以三种 Rust 语义变体与两种 Python
语义变体测量（全部热会话、逐册计时）：
  R0 = 生产语义：设 Calculation=xlAutomatic + CalculateFull + 轮询
  R1 = 不设 Calculation（保持默认自动），其余同 R0
  R2 = R1 且用工作簿级 Worksheets.Calculate 替代 CalculateFull（旧 Python 生产语义）
  P0 = 旧 Python ExcelSession：calculate_workbook（工作表级 Calculate + 0.1s 轮询）
  P1 = 旧 Python ExcelSession：Application.CalculateFull + 0.1s 轮询
Rust 变体经 `__audit-native-com-session` 的 recalculate_mode 请求字段驱动
（worker 已支持 auto/full/sheet 三种模式，见 native_office_com.rs）。
输出各变体逐册毫秒与均值，写入 stdout。

S3-F1：拆分 Rust 写出阶段的 write 与 validate（worker JSON 的
workbook_write_time_ms 拆成两段上报）。
"""

from __future__ import annotations

import json
import subprocess
import sys
import time
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

WORKSPACE = Path(__file__).resolve().parents[3]
AUDIT_BIN = WORKSPACE / "dist" / "open-data-audit.exe"
sys.path.insert(0, str(WORKSPACE / "audit-core" / "tests" / "compare"))
import perf_compare_20261030 as base  # noqa: E402

# base module already inserts the old Python source root into sys.path
# (DEFAULT_PY_SOURCE); mirror that for the locate script's own imports.
import sys as _sys  # noqa: E402
_python_source = None
for argument_index, argument in enumerate(sys.argv):
    if argument == "--python-source" and argument_index + 1 < len(sys.argv):
        _python_source = Path(sys.argv[argument_index + 1])
if _python_source is None:
    _python_source = base.DEFAULT_PY_SOURCE
_sys.path.insert(0, str(_python_source / "core" / "src"))


def build_inputs(root: Path):
    root.mkdir(parents=True, exist_ok=True)
    inputs = []
    for index in range(2):
        path = root / f"audit-copy-{index}.xlsx"
        base.build_s1_formula_workbook(path)
        inputs.append(path)
    return inputs


def rust_batch(inputs, outputs, mode):
    started = time.perf_counter()
    process = subprocess.Popen(
        [str(AUDIT_BIN), "__audit-native-com-session", "excel"],
        stdin=subprocess.PIPE, stdout=subprocess.PIPE,
        stderr=subprocess.PIPE, text=True, encoding="utf-8", bufsize=1,
    )
    per_book, stages_all = [], []
    try:
        ready = json.loads(process.stdout.readline())
        if ready.get("status") != "session_ready" or not ready.get("com_available"):
            raise RuntimeError(f"worker unavailable: {ready}")
        session_create_ms = ready.get("session_create_ms")
        pid = ready.get("engine_process_id")
        for src, dst in zip(inputs, outputs, strict=True):
            t0 = time.perf_counter()
            request = {"type": "recalculate", "input": str(src), "output": str(dst),
                       "mode": mode}
            process.stdin.write(json.dumps(request, ensure_ascii=False) + "\n")
            process.stdin.flush()
            response = json.loads(process.stdout.readline())
            if response.get("status") != "recalculated":
                raise RuntimeError(f"recalc failed: {response}")
            if response.get("engine_process_id") != pid:
                raise RuntimeError("process not reused")
            per_book.append(round((time.perf_counter() - t0) * 1000, 1))
            stages_all.append(response.get("stages_ms") or {})
        process.stdin.write(json.dumps({"type": "shutdown"}) + "\n")
        process.stdin.flush()
        process.stdin.close()
        if process.wait(timeout=30) != 0:
            raise RuntimeError(process.stderr.read()[-400:])
    finally:
        if process.poll() is None:
            process.kill()
    total = round((time.perf_counter() - started) * 1000, 1)
    return {
        "per_book_ms": per_book,
        "total_ms": total,
        "stages": stages_all,
        "session_create_ms": session_create_ms,
    }


def python_batch(inputs, outputs, mode):
    """mode: sheet=工作表级 Calculate；full=Application.CalculateFull。"""
    from base_audit.excel_com import ExcelSession
    per_book, stages_all = [], []
    session_obj = ExcelSession("Microsoft Excel")
    create_started = time.perf_counter()
    session = session_obj.__enter__()
    session_create_ms = round((time.perf_counter() - create_started) * 1000, 1)
    started = time.perf_counter()
    try:
        run_books(session, inputs, outputs, mode, per_book, stages_all)
    finally:
        session_obj.__exit__(None, None, None)
    total = round((time.perf_counter() - started) * 1000, 1)
    return {
        "per_book_ms": per_book,
        "total_ms": total,
        "stages": stages_all,
        "session_create_ms": session_create_ms,
    }


def run_books(session, inputs, outputs, mode, per_book, stages_all):
        for src, dst in zip(inputs, outputs, strict=True):
            stages = {}
            t0 = time.perf_counter()
            workbook = session.open_workbook(src, read_only=False)
            stages["open"] = round((time.perf_counter() - t0) * 1000, 1)
            t1 = time.perf_counter()
            if mode == "full":
                session.excel.CalculateFull()
            else:
                session.calculate_workbook(workbook)
            stages["calculate"] = round((time.perf_counter() - t1) * 1000, 1)
            t2 = time.perf_counter()
            workbook.SaveAs(str(dst), 51)
            stages["save_as"] = round((time.perf_counter() - t2) * 1000, 1)
            t3 = time.perf_counter()
            session.close_workbook(workbook, save=False)
            stages["close"] = round((time.perf_counter() - t3) * 1000, 1)
            per_book.append(round((time.perf_counter() - t0) * 1000, 1))
            stages_all.append(stages)


def main():
    rounds = 3
    import tempfile
    root = Path(tempfile.mkdtemp(prefix="s1f1_locate_"))
    inputs = build_inputs(root / "inputs")
    variants = [
        ("R0 rust production (set+CalculateFull+poll)", lambda o: rust_batch(inputs, o, "full")),
        ("R1 rust no-set (default auto + CalculateFull + poll)", lambda o: rust_batch(inputs, o, "full_no_set")),
        ("P0 python production (worksheets.Calculate + poll)", lambda o: python_batch(inputs, o, "sheet")),
        ("P1 python Application.CalculateFull + poll", lambda o: python_batch(inputs, o, "full")),
    ]
    report = {"rounds": rounds, "variants": {}}
    for label, call in variants:
        runs = []
        for index in range(rounds):
            outputs = [root / f"out-{label.split()[0]}-{index}-{i}.xlsx" for i in range(2)]
            runs.append(call(outputs))
        totals = [run["total_ms"] for run in runs]
        per_book_avg = [
            round(sum(run["per_book_ms"][book] for run in runs) / rounds, 1)
            for book in range(len(inputs))
        ]
        # average stage map across books and rounds
        stage_names = sorted({k for run in runs for s in run["stages"] for k in s})
        stage_avg = {}
        for name in stage_names:
            values = [s[name] for run in runs for s in run["stages"] if name in s]
            stage_avg[name] = round(sum(values) / len(values), 1) if values else None
        report["variants"][label] = {
            "totals_ms": totals,
            "median_ms": sorted(totals)[len(totals) // 2],
            "per_book_avg_ms": per_book_avg,
            "stage_avg_ms": stage_avg,
        }
        print(f"{label}\n  totals={totals} median={report['variants'][label]['median_ms']}"
              f"\n  per-book avg={per_book_avg}\n  stages avg={stage_avg}\n")

    out = root / "s1f1_locate_report.json"
    out.write_text(json.dumps(report, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"report: {out}")


if __name__ == "__main__":
    main()
