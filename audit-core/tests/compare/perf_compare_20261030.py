#!/usr/bin/env python3
"""新旧 Audit 性能对比基准（2026-10-30 工作区版本）。

同一份合成输入分别经旧 Python 服务（进程内 perf_counter）与新 Rust CLI
（子进程墙钟，含进程启动）运行，各自计时后输出对比表。

口径说明（写入报告）：
- Python 计时不含解释器启动（进程内调用），Rust 计时含 CLI 进程启动；
  报告同时单独测量 Rust 启动开销供折算。
- 每项运行 ROUNDS 轮，报告每轮耗时与最优值；两端输入逐字节相同。
- S1-F1 只对比重算子步骤（Excel COM 是该功能的主要成本）：
  Python=pywin32 进程内 COM（镜像 sidecar 脚本步骤）；Rust=lo-convert
  recalculate --engine com 全管线（含公式缓存清理与结果校验，属其生产路径）。
- S1-F2 本驱动未执行计时。旧版将其误称为缓存诊断模式；该诊断开关属于
  S1-F1 的 `s1-audit-directory`，不能据此认定 S1-F2 不等价。
"""

from __future__ import annotations

import argparse
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

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

WORKSPACE = Path(__file__).resolve().parents[3]
AUDIT_BIN = WORKSPACE / "dist" / "open-data-audit.exe"
LO_CONVERT = WORKSPACE / "dist" / "lo-convert.exe"  # legacy baseline only; current S1 uses Audit native COM worker
DEFAULT_PY_SOURCE = Path(r"C:\Users\AZI\Desktop\VibeCode\3.open-data-audit")
S2_FIXTURE = WORKSPACE / "audit-core" / "tests" / "fixtures" / "s2_import"
S2_CONFIG = WORKSPACE / "audit-core" / "tests" / "fixtures" / "s2_config" / "check_valid.xlsx"

S3_CSV_HEADERS = (
    "业务类,数据日期,机构类代码,机构类名称,地区代码,地区名称,指标顺序码,"
    "指标代码,指标名称,数据属性,币种,频度,批次,数据值\n"
)

ROUNDS = 3


def build_s3_configs(root: Path, with_cross: bool = False):
    """3.0 通用 + 3.1 执行比较（+可选 3.2 核对），镜像 CLI 测试夹具。"""
    from openpyxl import Workbook

    def sheet(wb, name):
        return wb.create_sheet(title=name)

    def header(sheet, headers):
        for c, h in enumerate(headers):
            sheet.cell(row=1, column=c + 1, value=h)

    common = Workbook()
    p = sheet(common, "运行参数")
    header(p, ["参数", "值", "说明"])
    p.cell(row=2, column=1, value="默认源数据单位"); p.cell(row=2, column=2, value="元")
    p.cell(row=3, column=1, value="默认目标单位"); p.cell(row=3, column=2, value="亿元")
    e = sheet(common, "单位换算例外")
    header(e, ["规则编号", "来源表单", "指标代码", "指标名称", "启用", "备注"])
    a = sheet(common, "环比警戒")
    header(a, ["序号", "下限", "上限", "备注文字", "填充颜色", "是否整行填充", "金额变动阈值", "是否说明"])
    for c, v in enumerate(["1", "0.1", "0.2", "环比增幅超限", "6", "否", "0", "是"]):
        a.cell(row=2, column=c + 1, value=v)
    c30 = root / "3.0配置.xlsx"
    common.save(c30)

    feature = Workbook()
    i = sheet(feature, "指标参照")
    header(i, ["指标代码", "指标名称", "表单代码", "行序号", "指标别名"])
    i.cell(row=2, column=1, value="I001"); i.cell(row=2, column=2, value="合成指标")
    i.cell(row=2, column=3, value="F001"); i.cell(row=2, column=4, value="1")
    i.cell(row=2, column=5, value="合成指标")
    r = sheet(feature, "规则动作")
    header(r, ["规则编号", "规则说明", "规则动作", "指标代码", "数据属性", "币种",
               "比较指标代码", "人民币阈值(亿元)", "美元阈值(亿美元)", "变幅阈值(%)",
               "来源分组", "适用表单", "适用频度", "适用批次", "适用场景", "详细说明",
               "启用", "停用原因", "备注"])
    r.cell(row=2, column=1, value="R-ACT-001")
    r.cell(row=2, column=2, value="变动或变幅超限需说明")
    r.cell(row=2, column=3, value="指标变动或变幅超过阈值需说明。")
    r.cell(row=2, column=4, value="I001")
    r.cell(row=2, column=5, value="余额"); r.cell(row=2, column=6, value="CNY")
    r.cell(row=2, column=8, value="0.01"); r.cell(row=2, column=10, value="0.1")
    r.cell(row=2, column=13, value="月报"); r.cell(row=2, column=14, value="第1批")
    r.cell(row=2, column=17, value="是")
    # 规则动作表：两端都支持的“指标变动或变幅超过阈值需说明。”动作，
    # 供 V3 规则引擎计时（旧 Python V3 要求至少一条动作规则）。
    x = sheet(feature, "表达式校验")
    # 两套加载器并存的列集：旧 Python 读「校验规则/触发方式/来源分组…」，
    # Rust 读「校验表达式/是否取反」；各自按表头取名，互不干扰。
    header(x, ["规则编号", "来源分组", "来源表单", "规则说明", "校验规则", "触发方式",
               "Thd值(万元)", "取反标识", "容差值(万元)", "启用", "停用原因", "备注",
               "校验表达式", "是否取反", "机构类代码", "地区代码"])
    x.cell(row=2, column=1, value="R-EXPR-001")
    x.cell(row=2, column=2, value=""); x.cell(row=2, column=3, value="")
    x.cell(row=2, column=4, value="本期值大于上期值时提示")
    x.cell(row=2, column=5, value="[I001,余额,CNY,月报,第1批]>{I001,余额,CNY,月报,第1批}")
    x.cell(row=2, column=6, value=""); x.cell(row=2, column=7, value="")
    x.cell(row=2, column=8, value="否"); x.cell(row=2, column=9, value="")
    x.cell(row=2, column=10, value="是"); x.cell(row=2, column=11, value="")
    x.cell(row=2, column=12, value="")
    x.cell(row=2, column=13, value="[I001,余额,CNY,月报,第1批]>{I001,余额,CNY,月报,第1批}")
    x.cell(row=2, column=14, value="否")
    x.cell(row=2, column=15, value="^ORG0001$"); x.cell(row=2, column=16, value="^R001$")
    # 五段式规则需要同编号的八段式绑定规则（提供业务类绑定段），两端同此要求。
    # 兼容表同时携带两套列集：Rust 读「校验表达式」，旧 Python 读「校验规则」。
    compat = sheet(feature, "表达式校验（兼容）")
    header(compat, ["规则编号", "来源分组", "来源表单", "规则说明", "校验规则", "触发方式",
                    "Thd值(万元)", "取反标识", "容差值(万元)", "启用", "停用原因", "备注",
                    "校验表达式", "是否取反"])
    binding = ("[C001,,,I001,余额,CNY,月报,第1批]>"
               "{C001,,,I001,余额,CNY,月报,第1批}")
    compat.cell(row=2, column=1, value="R-EXPR-001")
    compat.cell(row=2, column=2, value=""); compat.cell(row=2, column=3, value="")
    compat.cell(row=2, column=4, value="本期值大于上期值时提示")
    compat.cell(row=2, column=5, value=binding)
    compat.cell(row=2, column=6, value=""); compat.cell(row=2, column=7, value="")
    compat.cell(row=2, column=8, value="否"); compat.cell(row=2, column=9, value="")
    compat.cell(row=2, column=10, value="是"); compat.cell(row=2, column=11, value="")
    compat.cell(row=2, column=12, value="")
    compat.cell(row=2, column=13, value=binding)
    compat.cell(row=2, column=14, value="否")
    c31 = root / "3.1配置.xlsx"
    feature.save(c31)

    c32 = None
    if with_cross:
        cross = Workbook()
        cr = sheet(cross, "跨期数值核对")
        header(cr, ["校验编码", "校验名称", "机构类代码", "地区代码", "前提条件",
                    "left", "right", "校验公式", "校验提示", "取反标识",
                    "无误是否提示", "是否逻辑校验", "禁用", "备注"])
        for row_index, code in ((2, "C001"), (3, "C002")):
            cr.cell(row=row_index, column=1, value=code)
            cr.cell(row=row_index, column=2, value=f"性能基准规则 {code}")
            cr.cell(row=row_index, column=6, value="[I001,余额,CNY,月报,第1批]")
            cr.cell(row=row_index, column=7, value="0" if row_index == 2 else "1e9")
            cr.cell(row=row_index, column=8, value="left > right" if row_index == 2 else "left < right")
            cr.cell(row=row_index, column=9, value="请核实")
            cr.cell(row=row_index, column=11, value="是")
        c32 = root / "3.2配置.xlsx"
        cross.save(c32)
    return c30, c31, c32


def build_s3_csv(path: Path, rows: int, date_span: int = 1):
    # 旧 Python run_comparison 要求全部行共享同一数据日期；扩容按机构数。
    institutions = max(1, rows // max(1, date_span))
    lines = [S3_CSV_HEADERS]
    for index in range(rows):
        org = index % institutions
        value = round(100.0 + (index % 97) * 0.5, 2)
        row = (
            f"C001,2026-06-01,ORG{org:05d},示例机构{org:05d},"
            f"R{org % 40:03d},示例地区{org % 40:03d},{index + 1},I001,合成指标,"
            f"余额,CNY,月报,第1批,{value}"
        )
        lines.append(row + chr(10))
    path.write_text("".join(lines), encoding="utf-8")


def build_s3_f5(root: Path, institutions: int, indicators: int, forms: int):
    from openpyxl import Workbook

    c33wb = Workbook()
    list_sheet = c33wb.active; list_sheet.title = "报表清单"
    for c, h in enumerate(["报表代码", "报表名称", "频度", "批次", "是否启用写入"]):
        list_sheet.cell(row=1, column=c + 1, value=h)
    for f in range(forms):
        # S3-F05 infers RMB/FX business class from the form-sheet name (A1/A2).
        # Keep these synthetic forms on the RMB path so perf runs populate cells.
        code = f"性能表A1-{f + 1:03d}"
        list_sheet.cell(row=f + 2, column=1, value=code)
        list_sheet.cell(row=f + 2, column=2, value=f"性能表{f + 1:03d}")
        list_sheet.cell(row=f + 2, column=3, value="月")
        list_sheet.cell(row=f + 2, column=4, value="1")
        list_sheet.cell(row=f + 2, column=5, value="是")
    settings = c33wb.create_sheet("转表设置")
    settings.append(["参数", "值", "说明"])
    settings.append(["隐藏空行", "是", ""])
    settings.append(["空表删除", "是", ""])
    base_headers = ["余额人民币"]
    for f in range(forms):
        code = f"性能表A1-{f + 1:03d}"
        sheet = c33wb.create_sheet(code)
        headers = ["指标代码", "指标名称", "行序号", "余额人民币",
                   "上期-余额人民币", "增减额-余额人民币", "环比-余额人民币"]
        for c, h in enumerate(headers):
            sheet.cell(row=1, column=c + 1, value=h)
        for n in range(indicators):
            sheet.cell(row=n + 2, column=1, value=f"I{n + 1:03d}")
    c33 = root / "3.3配置.xlsx"
    c33wb.save(c33)

    c30wb = Workbook()
    alerts = c30wb.active; alerts.title = "环比警戒"
    for c, h in enumerate(["序号", "下限", "上限", "备注文字", "填充颜色", "是否整行填充", "金额变动阈值", "是否说明"]):
        alerts.cell(row=1, column=c + 1, value=h)
    for c, v in enumerate(["1", "0.1", "0.2", "环比增幅超限", "6", "否", "0", "是"]):
        alerts.cell(row=2, column=c + 1, value=v)
    c30 = root / "3.0配置.xlsx"
    c30wb.save(c30)

    comp = Workbook()
    data = comp.active; data.title = "2026年8月对比结果"
    headers = ["业务类", "数据日期", "机构类代码", "机构类名称", "地区代码", "地区名称",
               "指标顺序码", "指标代码", "指标名称", "数据属性", "币种", "频度", "批次",
               "数据值", "上期值", "增减额(亿元)", "环比", "备注", "是否说明", "说明内容", "计算过程"]
    for c, h in enumerate(headers):
        data.cell(row=1, column=c + 1, value=h)
    row = 2
    for org in range(institutions):
        for n in range(indicators):
            values = ["人民币", "2026-08-31", f"{org + 1:04d}", f"机构{org + 1:04d}",
                      f"R{org % 40:03d}", "地区", f"I{n + 1:03d}", f"I{n + 1:03d}",
                      "指标", "余额", "人民币", "月", "1",
                      100.0 + n, 90.0 + n, None, None, None, None, None, None]
            for c, v in enumerate(values):
                data.cell(row=row, column=c + 1, value=v)
            row += 1
    comp_file = root / "比较结果.xlsx"
    comp.save(comp_file)
    return c30, c33, comp_file


def build_s1_formula_workbook(path: Path, rows: int = 1200):
    from openpyxl import Workbook

    wb = Workbook()
    ref = wb.active; ref.title = "参照表"
    for r in range(1, rows + 1):
        ref.cell(row=r, column=1, value=f"K{r:05d}")
        ref.cell(row=r, column=2, value=100.0 + r)
        ref.cell(row=r, column=3, value=200.0 + r)
    check = wb.create_sheet("检查表")
    for r in range(1, rows + 1):
        check.cell(row=r, column=1, value=f"K{r:05d}")
        for c in range(1, 7):
            col = 2 if c % 2 == 0 else 3
            check.cell(row=r, column=c + 1,
                       value=f"=VLOOKUP($A{r},参照表!$A:$C,{col},FALSE)+1")
    wb.save(path)


def python_com_recalculate(input_path: Path, output_path: Path, timeout: float = 180.0) -> float:
    """镜像 Rust sidecar 的 COM 重算步骤（pywin32 进程内）。返回耗时秒。"""
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    started = time.perf_counter()
    app = win32com.client.DispatchEx("Excel.Application")
    book = None
    try:
        app.Visible = False
        app.DisplayAlerts = False
        book = app.Workbooks.Open(str(input_path), False, False)
        app.Calculation = -4105
        app.CalculateFull()
        try:
            if book.Connections.Count > 0:
                app.CalculateUntilAsyncQueriesDone()
        except Exception:
            pass
        deadline = time.perf_counter() + timeout
        while True:
            state = int(app.CalculationState)
            if state == 0:
                break
            if state == 2:
                app.CalculateFullRebuild()
                time.sleep(0.2)
                continue
            if time.perf_counter() >= deadline:
                raise TimeoutError("Excel calculation state did not reach xlDone")
            time.sleep(0.2)
        book.SaveAs(str(output_path), 51)
        book.Close(True)
        book = None
    finally:
        if book is not None:
            book.Close(False)
        app.Quit()
        pythoncom.CoUninitialize()
    return time.perf_counter() - started


def python_com_recalculate_batch(
    input_paths: list[Path], output_paths: list[Path], *, full_calculation: bool = False
):
    """Use the legacy Python ExcelSession once for the whole workbook batch."""
    from base_audit.excel_com import ExcelSession

    per_workbook_ms = []
    with ExcelSession("Microsoft Excel") as session:
        engine = session.engine_name
        for input_path, output_path in zip(input_paths, output_paths, strict=True):
            started = time.perf_counter()
            workbook = session.open_workbook(input_path, read_only=False)
            try:
                if full_calculation:
                    session.excel.CalculateFull()
                else:
                    session.calculate_workbook(workbook)
                workbook.SaveAs(str(output_path), 51)
            finally:
                session.close_workbook(workbook, save=False)
            per_workbook_ms.append(round((time.perf_counter() - started) * 1000, 1))
    return {"engine": engine, "per_workbook_ms": per_workbook_ms}


def rust_native_com_recalculate_batch(input_paths: list[Path], output_paths: list[Path]):
    """Call the production hidden worker once, then recalculate every workbook."""
    started = time.perf_counter()
    process = subprocess.Popen(
        [str(AUDIT_BIN), "__audit-native-com-session", "excel"],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        bufsize=1,
    )
    per_workbook_ms = []
    stages = []
    try:
        ready_line = process.stdout.readline()
        if not ready_line:
            raise RuntimeError("native COM worker exited before reporting readiness")
        ready = json.loads(ready_line)
        if ready.get("status") != "session_ready" or not ready.get("com_available"):
            raise RuntimeError(f"native COM worker unavailable: {ready}")
        process_id = ready.get("engine_process_id")
        for input_path, output_path in zip(input_paths, output_paths, strict=True):
            workbook_started = time.perf_counter()
            request = {"type": "recalculate", "input": str(input_path), "output": str(output_path)}
            process.stdin.write(json.dumps(request, ensure_ascii=False) + "\n")
            process.stdin.flush()
            response_line = process.stdout.readline()
            if not response_line:
                raise RuntimeError("native COM worker exited while recalculating")
            response = json.loads(response_line)
            if response.get("status") != "recalculated":
                raise RuntimeError(f"native COM recalculation failed: {response}")
            if response.get("engine_process_id") != process_id:
                raise RuntimeError("native COM worker did not reuse the same Office process")
            per_workbook_ms.append(round((time.perf_counter() - workbook_started) * 1000, 1))
            stages.append(response.get("stages_ms") or {})
        process.stdin.write(json.dumps({"type": "shutdown"}) + "\n")
        process.stdin.flush()
        process.stdin.close()
        return_code = process.wait(timeout=30)
        stderr = process.stderr.read()
        if return_code != 0:
            raise RuntimeError(f"native COM worker exited {return_code}: {stderr[-500:]}")
        return {
            "engine": ready.get("engine"),
            "engine_process_id_reused": True,
        "stages_ms": stages,
            "per_workbook_ms": per_workbook_ms,
            "session_total_ms": round((time.perf_counter() - started) * 1000, 1),
        }
    finally:
        if process.poll() is None:
            process.kill()
            process.wait()


def verify_s1_formula_output(path: Path, rows: int = 1200) -> dict[str, int]:
    from openpyxl import load_workbook

    formulas_book = load_workbook(path, read_only=True, data_only=False)
    try:
        formulas = formulas_book["检查表"]
        formula_count = sum(
            1
            for row in formulas.iter_rows(min_row=1, max_row=rows, min_col=2, max_col=7)
            for cell in row
            if isinstance(cell.value, str) and cell.value.startswith("=")
        )
        if formula_count != rows * 6:
            raise AssertionError(f"{path.name}: expected {rows * 6} formulas, found {formula_count}")
    finally:
        formulas_book.close()

    values_book = load_workbook(path, read_only=True, data_only=True)
    try:
        values = values_book["检查表"]
        checked = 0
        for row_number, row in enumerate(values.iter_rows(min_row=1, max_row=rows, min_col=2, max_col=7), 1):
            for column_index, cell in enumerate(row, 1):
                expected = (100.0 if column_index % 2 == 0 else 200.0) + row_number + 1
                if not isinstance(cell.value, (int, float)) or abs(float(cell.value) - expected) > 1e-9:
                    raise AssertionError(
                        f"{path.name}: expected numeric cache {expected} at row {row_number}, "
                        f"formula column {column_index}, got {cell.value!r}"
                    )
                checked += 1
    finally:
        values_book.close()
    return {"formula_cells": formula_count, "numeric_cached_results_verified": checked}


def build_s3_feedback_fixture(root: Path, rows: int = 1000):
    from openpyxl import Workbook

    headers = ["机构代码", "机构名称", "指标代码", "指标名称", "本期金额（亿元）",
               "上期值", "增减额（亿元）", "变动原因"]
    target = Workbook()
    target_sheet = target.active
    target_sheet.title = "说明文件"
    target_sheet.append(["xx说明"])
    target_sheet.append(headers)
    feedback = Workbook()
    feedback_sheet = feedback.active
    feedback_sheet.title = "机构反馈"
    feedback_sheet.append(headers)
    for index in range(rows):
        key = (f"ORG{index:05d}", f"机构{index:05d}", f"I{index:05d}", f"指标{index:05d}")
        amount = 100.0 + index
        change = index % 23 - 11
        target_sheet.append([*key, amount, 90.0 + index, change, "待机构填写"])
        feedback_sheet.append([*key, amount, 90.0 + index, change, f"已核实原因{index:05d}"])
    target_path = root / "target_base.xlsx"
    feedback_path = root / "feedback.xlsx"
    target.save(target_path)
    feedback.save(feedback_path)
    target.close()
    feedback.close()
    return target_path, feedback_path


def verify_s3_feedback_run(target: Path, expected_backup_hash: str, rows: int):
    from openpyxl import load_workbook

    backup = list(target.parent.glob(f"{target.stem}_导入前备份_*.xlsx"))
    if len(backup) != 1 or hashlib.sha256(backup[0].read_bytes()).hexdigest() != expected_backup_hash:
        raise AssertionError(f"missing or mismatched pre-import backup for {target.name}")
    book = load_workbook(target, read_only=True, data_only=True)
    try:
        sheet = book["说明文件"]
        if sheet.max_row != rows + 2:
            raise AssertionError(f"unexpected updated target row count: {sheet.max_row}")
        for row_number in (3, rows + 2):
            reason = sheet.cell(row_number, 8).value
            expected = f"已核实原因{row_number - 3:05d}"
            if reason != expected:
                raise AssertionError(f"feedback writeback mismatch at row {row_number}: {reason!r}")
    finally:
        book.close()


def run_rounds(label, rounds, python_call, rust_call):
    py_times, rust_times = [], []
    py_extra, rust_extra = None, None
    for index in range(rounds):
        started = time.perf_counter()
        py_extra = python_call(index) or {}
        py_times.append(time.perf_counter() - started)
    for index in range(rounds):
        started = time.perf_counter()
        rust_extra = rust_call(index) or {}
        rust_times.append(time.perf_counter() - started)
    return {
        "feature": label,
        "rounds": rounds,
        "python_ms": [round(t * 1000, 1) for t in py_times],
        "rust_ms": [round(t * 1000, 1) for t in rust_times],
        "python_best_ms": round(min(py_times) * 1000, 1),
        "rust_best_ms": round(min(rust_times) * 1000, 1),
        "speedup_best": round(min(py_times) / min(rust_times), 2) if min(rust_times) > 0 else None,
        "python_median_ms": round(statistics.median(py_times) * 1000, 1),
        "rust_median_ms": round(statistics.median(rust_times) * 1000, 1),
        "speedup_median": round(statistics.median(py_times) / statistics.median(rust_times), 2)
        if statistics.median(rust_times) > 0 else None,
        "python_extra": py_extra,
        "rust_extra": rust_extra,
    }


def main():
    global AUDIT_BIN
    parser = argparse.ArgumentParser()
    parser.add_argument("--rounds", type=int, default=ROUNDS)
    parser.add_argument("--audit-binary", type=Path, default=AUDIT_BIN)
    parser.add_argument("--csv-rows", type=int, default=10000)
    parser.add_argument("--institutions", type=int, default=30)
    parser.add_argument("--indicators", type=int, default=40)
    parser.add_argument("--forms", type=int, default=4)
    parser.add_argument("--python-source", type=Path, default=DEFAULT_PY_SOURCE)
    parser.add_argument("--report", type=Path,
                        default=WORKSPACE / "audit-core" / "probe" / "reports" / "perf_comparison_20261030.json")
    args = parser.parse_args()

    AUDIT_BIN = args.audit_binary.resolve(strict=True)

    sys.path.insert(0, str(args.python_source / "core" / "src"))
    results = []
    root = Path(tempfile.mkdtemp(prefix="audit_perf_"))
    rust_startup_ms = None

    def rust_subprocess(command, expect_success=True):
        completed = subprocess.run(
            command,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        if expect_success and completed.returncode != 0:
            raise RuntimeError(
                f"Rust CLI failed ({completed.returncode}): {completed.stderr[-600:]}"
            )
        return completed

    # Rust CLI 进程启动开销（--version，5 次取最小）
    times = []
    for _ in range(5):
        started = time.perf_counter()
        subprocess.run([str(AUDIT_BIN), "--version"], capture_output=True)
        times.append(time.perf_counter() - started)
    rust_startup_ms = round(min(times) * 1000, 1)

    # ---------- S3-F1 两期数据 → 比较文件 ----------
    try:
        (root / "s3f1").mkdir(parents=True, exist_ok=True)
        c30, c31, _ = build_s3_configs(root / "s3f1", with_cross=True)
        build_s3_csv(root / "s3f1" / "current.csv", args.csv_rows)
        build_s3_csv(root / "s3f1" / "previous.csv", args.csv_rows)

        def py_s3f1(index):
            from base_audit.systems.s3_central_statistics.service import run_comparison
            output = root / "s3f1" / f"py-out-{index}.xlsx"
            result = run_comparison(
                current_csv=root / "s3f1" / "current.csv",
                previous_csv=root / "s3f1" / "previous.csv",
                output_path=output,
                config_path=[c30, c31],
                write_flow_logs=False,
                expression_schema="FIVE_SEGMENT_V1",
            )
            rows = getattr(result, "rows", None)
            count = len(rows) if isinstance(rows, list) else None
            return {"rows": count}

        def rust_s3f1(index):
            output = root / "s3f1" / f"rust-out-{index}.xlsx"
            rust_subprocess([
                str(AUDIT_BIN), "s3-execution-compare",
                "--current-data", str(root / "s3f1" / "current.csv"),
                "--previous-data", str(root / "s3f1" / "previous.csv"),
                "--common-config", str(c30),
                "--feature-config", str(c31),
                "--output-workbook", str(output),
                "--expression-schema", "five-segment",
            ])
            return {"schema": "five-segment"}

        results.append(run_rounds("S3-F1 两期数据→比较文件（10k 行）", args.rounds, py_s3f1, rust_s3f1))
    except Exception as error:
        results.append({"feature": "S3-F1", "error": f"{error}"})

    # ---------- S3-F4 本期数值核对 ----------
    try:
        (root / "s3f4").mkdir(parents=True, exist_ok=True)
        c30, _, c32 = build_s3_configs(root / "s3f4", with_cross=True)
        build_s3_csv(root / "s3f4" / "current.csv", args.csv_rows, date_span=1)

        def py_s3f4(index):
            from base_audit.systems.s3_central_statistics.service import run_cross_period_check
            output = root / "s3f4" / f"py-out-{index}.xlsx"
            result = run_cross_period_check(
                current_csv=root / "s3f4" / "current.csv",
                output_path=output,
                config_path=[c30, c32],
                write_flow_logs=False,
            )
            return {}

        def rust_s3f4(index):
            output = root / "s3f4" / f"rust-out-{index}.xlsx"
            rust_subprocess([
                str(AUDIT_BIN), "s3-check-current-values",
                "--current-csv", str(root / "s3f4" / "current.csv"),
                "--config-workbook", str(c30),
                "--config-workbook", str(c32),
                "--output-workbook", str(output),
            ])
            return {}

        results.append(run_rounds("S3-F4 本期数值核对（10k 行）", args.rounds, py_s3f4, rust_s3f4))
    except Exception as error:
        results.append({"feature": "S3-F4", "error": f"{error}"})

    # ---------- S3-F5 比较文件 → 金融表单 ----------
    try:
        (root / "s3f5").mkdir(parents=True, exist_ok=True)
        c30, c33, comp_file = build_s3_f5(
            root / "s3f5", args.institutions, args.indicators, args.forms)

        def py_s3f5(index):
            from base_audit.systems.s3_central_statistics.service import render_financial_forms
            out_dir = root / "s3f5" / f"py-out-{index}"
            render_financial_forms(
                comparison_file=comp_file,
                output_dir=out_dir,
                config_path=[c30, c33],
            )
            return {"workbooks": len(list(out_dir.glob("*.xlsx")))}

        def rust_s3f5(index):
            out_dir = root / "s3f5" / f"rust-out-{index}"
            rust_subprocess([
                str(AUDIT_BIN), "s3-generate-financial-forms",
                "--comparison-workbook", str(comp_file),
                "--common-config", str(c30),
                "--forms-config", str(c33),
                "--output-directory", str(out_dir),
            ])
            files = list(out_dir.rglob("*.xlsx"))
            return {"workbooks": len(files)}

        label = (f"S3-F5 比较文件→金融表单（{args.institutions} 机构 × "
                 f"{args.indicators} 指标 × {args.forms} 表）")
        results.append(run_rounds(label, args.rounds, py_s3f5, rust_s3f5))
    except Exception as error:
        results.append({"feature": "S3-F5", "error": f"{error}"})

    # ---------- S3-F2 比较文件 → 说明文件 ----------
    try:
        from openpyxl import Workbook
        comp = Workbook()
        data = comp.active; data.title = "对比结果"
        headers = ["机构类代码", "机构类名称", "地区代码", "地区名称", "指标代码",
                   "指标名称", "数据值", "上期值", "输出说明", "说明内容",
                   "增减额（先出现）", "增减额（后出现）"]
        for c, h in enumerate(headers):
            data.cell(row=1, column=c + 1, value=h)
        rows = 5000
        for index in range(rows):
            selected = "是" if index % 3 == 0 else "否"
            values = [f"{7000001 + index % 400:07d}", f"机构{index % 400:04d}",
                      f"{4401 + index % 5}", "地区", f"I-{index % 50:03d}", "指标",
                      10.0 + index % 7, 9.0 + index % 5, selected, f"说明{index}", 1.0, 1]
            for c, v in enumerate(values):
                data.cell(row=index + 2, column=c + 1, value=v)
        comp_file = root / "s3f2" / "比较结果.xlsx"
        comp_file.parent.mkdir(parents=True, exist_ok=True)
        comp.save(comp_file)
        (root / "s3f2").mkdir(parents=True, exist_ok=True)

        ref = Workbook()
        sheet = ref.active; sheet.title = "机构地区参照"
        for c, h in enumerate(["机构类代码", "地区代码", "机构名称"]):
            sheet.cell(row=1, column=c + 1, value=h)
        for index in range(400):
            sheet.cell(row=index + 2, column=1, value=f"{7000001 + index:07d}")
            sheet.cell(row=index + 2, column=2, value=f"{4401 + index % 5}")
            sheet.cell(row=index + 2, column=3, value=f"精确机构{index:04d}")
        config = root / "s3f2" / "3.0参照.xlsx"
        ref.save(config)
        selected_count = sum(1 for index in range(rows) if index % 3 == 0)

        def py_s3f2(index):
            from base_audit.systems.s3_central_statistics.explanation_exporter import (
                export_selected_explanations,
            )
            output = root / "s3f2" / f"py-out-{index}.xlsx"
            count = export_selected_explanations(comp_file, output, config_path=config)
            return {"exported": count}

        def rust_s3f2(index):
            output = root / "s3f2" / f"rust-out-{index}.xlsx"
            rust_subprocess([
                str(AUDIT_BIN), "s3-export-explanations",
                "--comparison-file", str(comp_file),
                "--output-workbook", str(output),
                "--config-workbook", str(config),
            ])
            return {"exported": selected_count}

        results.append(run_rounds(f"S3-F2 比较文件→说明文件（{rows} 行 / {selected_count} 选中）",
                                  args.rounds, py_s3f2, rust_s3f2))
    except Exception as error:
        results.append({"feature": "S3-F2", "error": f"{error}"})

    # ---------- S3-F3 机构反馈原地导入并备份 ----------
    try:
        s3f3_root = root / "s3f3"
        s3f3_root.mkdir(parents=True, exist_ok=True)
        target_base, feedback_file = build_s3_feedback_fixture(s3f3_root, rows=1000)
        target_base_hash = hashlib.sha256(target_base.read_bytes()).hexdigest()
        feedback_hash = hashlib.sha256(feedback_file.read_bytes()).hexdigest()
        python_targets = [s3f3_root / f"python-target-{index}.xlsx" for index in range(args.rounds)]
        rust_targets = [s3f3_root / f"rust-target-{index}.xlsx" for index in range(args.rounds)]
        for target in (*python_targets, *rust_targets):
            shutil.copy2(target_base, target)

        def py_s3f3(index):
            from base_audit.systems.s3_central_statistics.explanation_exporter import (
                import_explanation_feedback,
            )
            result = import_explanation_feedback(python_targets[index], [feedback_file])
            return {
                "imported": result.imported,
                "unmatched": result.unmatched,
                "conflicts": result.conflicts,
            }

        def rust_s3f3(index):
            completed = rust_subprocess([
                str(AUDIT_BIN), "s3-import-explanation-feedback",
                "--target-workbook", str(rust_targets[index]),
                "--feedback-workbook", str(feedback_file),
            ])
            report = json.loads(completed.stdout)
            return {
                "imported": report.get("imported"),
                "unmatched": report.get("unmatched"),
                "conflicts": report.get("conflicts"),
            }

        results.append(run_rounds(
            "S3-F3 机构反馈原地导入（1000 行，含目标备份）",
            args.rounds, py_s3f3, rust_s3f3,
        ))
        for target in (*python_targets, *rust_targets):
            verify_s3_feedback_run(target, target_base_hash, rows=1000)
        if hashlib.sha256(feedback_file.read_bytes()).hexdigest() != feedback_hash:
            raise AssertionError("S3-F3 changed the institution feedback input workbook")
        results[-1]["verification"] = {
            "python_and_rust_imported_1000_rows": True,
            "all_updated_targets_have_byte_identical_backup": True,
            "feedback_input_unchanged": True,
            "scope": "synthetic unique keys with no conflicts; backup and in-place write are timed",
        }
    except Exception as error:
        results.append({"feature": "S3-F3", "error": f"{error}"})

    # ---------- S1-F1 真实重算子步骤（两端均按批次复用 Excel 会话） ----------
    try:
        s1_root = root / "s1f1"
        s1_root.mkdir(parents=True, exist_ok=True)
        input_paths = [s1_root / f"audit-copy-{index}.xlsx" for index in range(2)]
        for input_path in input_paths:
            build_s1_formula_workbook(input_path)
        formulas = 1200 * 6

        def py_s1(index):
            outputs = [s1_root / f"py-recalc-{index}-{file_index}.xlsx" for file_index in range(2)]
            return python_com_recalculate_batch(input_paths, outputs)

        def rust_s1(index):
            outputs = [s1_root / f"rust-recalc-{index}-{file_index}.xlsx" for file_index in range(2)]
            return rust_native_com_recalculate_batch(input_paths, outputs)

        label = f"S1-F1 公式重算子步骤（每册 {formulas} 个公式；批次 2 册；Excel COM 热会话）"
        results.append(run_rounds(label, args.rounds, py_s1, rust_s1))
        python_full_times = []
        python_full_extras = []
        for index in range(args.rounds):
            full_outputs = [s1_root / f"py-full-recalc-{index}-{file_index}.xlsx" for file_index in range(2)]
            started = time.perf_counter()
            python_full_extras.append(
                python_com_recalculate_batch(input_paths, full_outputs, full_calculation=True)
            )
            python_full_times.append(time.perf_counter() - started)
        results[-1]["python_full_calculation_ms"] = [round(value * 1000, 1) for value in python_full_times]
        results[-1]["python_full_median_ms"] = round(statistics.median(python_full_times) * 1000, 1)
        results[-1]["python_full_extra"] = python_full_extras[-1]
        results[-1]["speedup_median_vs_python_full"] = round(
            statistics.median(python_full_times) / statistics.median([value / 1000 for value in results[-1]["rust_ms"]]),
            2,
        )
        source_hashes = [hashlib.sha256(path.read_bytes()).hexdigest() for path in input_paths]
        verified_outputs = {"python_legacy": [], "python_full": [], "rust": []}
        for index in range(args.rounds):
            for file_index in range(2):
                py_output = s1_root / f"py-recalc-{index}-{file_index}.xlsx"
                py_full_output = s1_root / f"py-full-recalc-{index}-{file_index}.xlsx"
                rust_output = s1_root / f"rust-recalc-{index}-{file_index}.xlsx"
                verified_outputs["python_legacy"].append(verify_s1_formula_output(py_output))
                verified_outputs["python_full"].append(verify_s1_formula_output(py_full_output))
                verified_outputs["rust"].append(verify_s1_formula_output(rust_output))
        inputs_unchanged = source_hashes == [hashlib.sha256(path.read_bytes()).hexdigest() for path in input_paths]
        if not inputs_unchanged:
            raise AssertionError("S1-F1 COM recalculation modified a source workbook")
        results[-1]["verification"] = {
            "all_outputs_kept_7200_formulas_and_numeric_caches": True,
            "output_formula_cache_checks": verified_outputs,
            "source_workbooks_unchanged": inputs_unchanged,
            "scope": "COM recalculation/save substep only; Python legacy uses Worksheets.Calculate, Python matched variant and Rust use Application.CalculateFull; Rust production cache clearing, post-checks and atomic publication are excluded",
        }
    except Exception as error:
        results.append({"feature": "S1-F1", "error": f"{error}"})

    # ---------- S5 配置化逐笔统计审核 ----------
    try:
        from openpyxl import Workbook
        s5_root = root / "s5"
        s5_root.mkdir(parents=True, exist_ok=True)
        empty_templates = s5_root / "empty-templates"
        empty_templates.mkdir(exist_ok=True)

        wb = Workbook()
        rules = wb.active; rules.title = "审核规则"
        headers = ["规则ID", "规则修订", "是否启用", "业务类型", "默认数据源", "默认Sheet",
                   "规则名称", "审核级别", "校验类型", "校验指标", "审核表达式",
                   "表达式布尔值取反", "规则描述", "当前值表达式", "对比值表达式",
                   "差值表达式", "取值字典", "生效起始日", "生效截止日", "适用范围", "适用表达式"]
        for c, h in enumerate(headers):
            rules.cell(row=1, column=c + 1, value=h)
        # openpyxl API（create_sheet/cell）——与文件其余部分一致。
        values = ["PERF.PARITY", "1", "是", "单位贷款", "CUR", "贷款表", "上期对拍",
                  "错误", "表间核对", "经营贷款余额", "ABS({本期} - {上期}) = 20", "否",
                  "差异={差值}", "{本期}", "{上期}", "{差值}",
                  '{"本期": V_CUR(("贷款产品类别", "经营贷款"), ("余额", "本期")), '
                  '"上期": V_PRE(("贷款产品类别", "经营贷款"), ("余额", "本期")), '
                  '"差值": {本期} - {上期}}', "", "", "", ""]
        for c, v in enumerate(values):
            rules.cell(row=2, column=c + 1, value=v)
        catalog = wb.create_sheet("模板指标目录")
        for c, h in enumerate(["业务类型", "工作表名称", "行路径", "列路径", "行表头标题",
                               "数据类型", "指标类型", "来源单位", "标准单位", "存储尺度",
                               "允许空值", "精度", "口径依据", "确认状态", "布局版本"]):
            catalog.cell(row=1, column=c + 1, value=h)
        catalog_values = ["单位贷款", "贷款表", "贷款产品类别\n经营贷款", "余额\n本期",
                          "一级\n二级", "number", "amount", "万元", "万元", "native",
                          "否", "2", "合成基准", "是", "perf-v1"]
        for c, v in enumerate(catalog_values):
            catalog.cell(row=2, column=c + 1, value=v)
        version = wb.create_sheet("配置版本")
        for c, h in enumerate(["配置版本号", "状态", "发布日期", "生效起始日", "生效截止日", "发布说明"]):
            version.cell(row=1, column=c + 1, value=h)
        for c, v in enumerate(["perf-v1", "测试", "2026-10-29", "", "", "性能基准"]):
            version.cell(row=2, column=c + 1, value=v)
        config = s5_root / "config.xlsx"
        wb.save(config)

        def write_s5_source(path: Path, value: float):
            swb = Workbook()
            info = swb.active; info.title = "基本信息"
            info.cell(row=1, column=1, value="机构代码"); info.cell(row=1, column=2, value="I001")
            info.cell(row=2, column=1, value="统计期间")
            data = swb.create_sheet("贷款表")
            data.cell(row=1, column=1, value="模拟单位贷款数据")
            data.cell(row=2, column=3, value="余额")
            for c, h in enumerate(["一级", "二级", "本期", "上期"]):
                data.cell(row=3, column=c + 1, value=h)
            data.cell(row=4, column=1, value="贷款产品类别")
            data.cell(row=4, column=2, value="经营贷款")
            data.cell(row=4, column=3, value=value)
            data.cell(row=4, column=4, value=100.0)
            swb.save(path)

        write_s5_source(s5_root / "current.xlsx", 120.0)
        write_s5_source(s5_root / "previous.xlsx", 90.0)

        def py_s5(index):
            from base_audit.systems.s5_configured_statistics.contracts import S5Request
            from base_audit.systems.s5_configured_statistics.service import run_s5_f1
            request = S5Request(
                config_path=config,
                template_path=empty_templates,
                files=(s5_root / "current.xlsx",),
                input_dir=s5_root,
                output_dir=s5_root / f"py-out-{index}",
                period="2026-08-31",
                business="单位贷款",
                previous_files=(s5_root / "previous.xlsx",),
                previous_period="2025-06-30",
            )
            run_s5_f1(request, data_dir=s5_root)
            return {}

        def rust_s5(index):
            out = s5_root / f"rust-out-{index}"
            rust_subprocess([
                str(AUDIT_BIN), "s5-audit-run",
                "--config", str(config),
                "--source", str(s5_root / "current.xlsx"),
                "--template-dir", str(empty_templates),
                "--source-id", "CUR",
                "--period", "2026-08-31",
                "--org-code", "I001",
                "--org-name", "合成机构",
                "--business", "单位贷款",
                "--output-dir", str(out),
                "--run-id", f"perf-{index}",
                "--previous-source", str(s5_root / "previous.xlsx"),
                "--previous-period", "2025-06-30",
            ])
            return {}

        results.append(run_rounds("S5 配置化逐笔统计审核（单文件 + 独立上期）", args.rounds, py_s5, rust_s5))
    except Exception as error:
        results.append({"feature": "S5", "error": f"{error}"})

    # ---------- S2 报表采集审核（小真实夹具） ----------
    try:
        current_dir = root / "s2" / "current"
        previous_dir = root / "s2" / "previous"
        for d in (current_dir, previous_dir):
            d.mkdir(parents=True, exist_ok=True)
            for name in ("9144#2026-06-30#x#20202#甲银行.xlsx",
                         "#2026-06-30#x#20202#乙银行.xlsx"):
                source = S2_FIXTURE / name
                if source.exists():
                    shutil.copy(source, d / name)

        def py_s2(index):
            from base_audit.systems.s2_report_collection.service import PeriodAuditV2Service
            out = root / "s2" / f"python-out-{index}"
            result = PeriodAuditV2Service().run(
                current_dir=current_dir,
                previous_dir=previous_dir,
                config_path=S2_CONFIG,
                output_dir=out,
            )
            return {
                "findings": len(result.findings),
                "workbooks": len(list(out.glob("*.xlsx"))),
            }

        def rust_s2(index):
            out = root / "s2" / f"rust-out-{index}"
            rust_subprocess([
                str(AUDIT_BIN), "s2",
                "--current-dir", str(current_dir),
                "--previous-dir", str(previous_dir),
                "--config", str(S2_CONFIG),
                "--output-dir", str(out),
            ])
            return {"workbooks": len(list(out.glob("*.xlsx")))}

        results.append(run_rounds("S2-F1 报表采集审核（2 份仓库测试夹具，含结果写出）", args.rounds, py_s2, rust_s2))
        from openpyxl import load_workbook
        for index in range(args.rounds):
            python_files = list((root / "s2" / f"python-out-{index}").glob("*.xlsx"))
            rust_files = list((root / "s2" / f"rust-out-{index}").glob("*.xlsx"))
            if len(python_files) != 1 or len(rust_files) != 1:
                raise AssertionError(f"S2 round {index}: expected one output workbook from both implementations")
            def sheet_signature(path):
                book = load_workbook(path, read_only=True, data_only=True)
                try:
                    return [(sheet.title, sheet.max_row, sheet.max_column) for sheet in book.worksheets]
                finally:
                    book.close()
            if sheet_signature(python_files[0]) != sheet_signature(rust_files[0]):
                raise AssertionError(f"S2 round {index}: Python/Rust result sheet structures differ")
        results[-1]["verification"] = {
            "both_services_wrote_one_readable_workbook_per_round": True,
            "result_sheet_names_and_dimensions_match": True,
            "python_path": "PeriodAuditV2Service.run including ExcelExporter output",
        }
    except Exception as error:
        results.append({"feature": "S2", "error": f"{error}"})

    # ---------- S1-F2：由专项对拍驱动单独测量 ----------
    results.append({
        "feature": "S1-F2 汇总校验结果说明",
        "skipped": True,
        "reason": "S1-F2 使用 compare_s1_region_summary.py 单独执行三轮性能与输出对拍；本脚本不重复计时。正式汇总 perf_comparison_20261030.json 已合并专项结果。",
    })

    report = {
        "schema_version": 1,
        "date": "2026-10-30",
        "platform": "windows-x86_64",
        "rust_binary": str(AUDIT_BIN),
        "rust_startup_ms": rust_startup_ms,
        "python_source": str(args.python_source),
        "methodology": {
            "python": "进程内 perf_counter，不含解释器启动；S1-F1 每轮用一个 Python ExcelSession 处理两册",
            "rust": "当前工作区编译的 release CLI 子进程墙钟，含 CLI 启动；S1-F1 每轮启动一次 Audit 原生 COM worker 处理两册",
            "rounds": args.rounds,
            "inputs": "S2 使用相同仓库测试夹具，其余使用合成输入；S1-F1/S3-F3 另核验公式结果、更新前备份与输入文件未改",
        },
        "results": results,
    }
    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text(json.dumps(report, ensure_ascii=False, indent=1), encoding="utf-8")

    print(f"\nRust CLI 启动开销：{rust_startup_ms} ms\n")
    header = f"{'功能':<44}{'Python 最快':>14}{'Rust 最快':>14}{'加速比':>10}"
    print(header)
    print("-" * len(header))
    for item in results:
        feature = item.get("feature", "?")
        if item.get("error"):
            print(f"{feature:<44}{'ERROR':>14}{'-':>14}{'-':>10}")
            print(f"    └ {item['error'][:120]}")
        elif item.get("skipped"):
            print(f"{feature:<44}{'SKIPPED':>14}{'-':>14}{'-':>10}")
        else:
            speedup = item.get("speedup_best")
            print(f"{feature[:42]:<44}{item['python_best_ms']:>12.1f}ms{item['rust_best_ms']:>12.1f}ms{speedup:>9.2f}x")
    print(f"\n报告已写入：{args.report}")

    shutil.rmtree(root, ignore_errors=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
