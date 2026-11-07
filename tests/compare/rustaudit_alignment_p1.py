#!/usr/bin/env python3
"""RustAudit 对齐 3.open-data-audit 基线（P1 八功能真实数据对比）。

基线：3.open-data-audit/test-output/rustaudit_alignment/{S1F1,S1F2,S2F1,S3F1..S3F5}.json
（desktop@da7b1c3，真实数据、生产入口跑出；引擎设置见 environment.engine_settings_note：
 自动引擎(Excel COM) + conditional_format_evaluator=PYTHON + DIRECT_OOXML 公式/外部写入；
 对应 Rust 侧：calc_engine=auto（Windows=Excel COM）、S1 条件格式提取来源=原始源文件
 （PYTHON 求值语义）、空白格命中=纳入。）

用法：
  python tests/compare/rustaudit_alignment_p1.py [--runs 2] [--features s1f1,...]

每个功能跑 --runs 次，逐次与基线对比；关键值两次运行必须一致。
对比口径：
  - 计数类（成功/失败/条数/填充色分布）直接与基线 JSON 数值比对；
  - 内容类（S1F1 逐条 issue、S1F2/S3 工作簿）与基线 _work 产物做逐格值对比
    （写入器实现不同导致文件哈希必然不同，哈希只记录不作为一致依据）；
  - 原始输入文件前后 SHA-256 必须不变。
输出：probe/reports/rustaudit_alignment_report.json
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(r"C:\Users\AZI\Desktop\VibeCode\Rust-Audit")
EXE = ROOT / "dist" / "open-data-audit.exe"
OLD = Path(r"C:\Users\AZI\Desktop\VibeCode\3.open-data-audit")
ALIGN = OLD / "test-output" / "rustaudit_alignment"
DATA = OLD / "test-data"
CONFIG_REAL = OLD / "config-real"
WORK = ROOT / "probe" / "reports" / "rustaudit_alignment_work"
REPORT = ROOT / "probe" / "reports" / "rustaudit_alignment_report.json"

S1F1_DIR = DATA / "S1F1_external_fallback_20261106"
S1F1_INPUTS = S1F1_DIR / "核查表"
S1F1_TEMPLATES = S1F1_DIR / "模板文件_命名区域"
MANIFEST = ROOT / "测试数据" / "S1F1_新旧对比_20261105_231620" / "rust-manifest-with-external.json"
S1F2_DIR = DATA / "S1F2"
S2F1_DIR = DATA / "S2F1"
S3_DIR = DATA / "S3"
S1F2_TEMPLATE = S1F1_TEMPLATES / "！汇总校验结果及报送说明.xlsx"
S1_HISTORY = CONFIG_REAL / "1.逐笔统计系统_配置.xlsx"

EXTERNAL_20261016 = CONFIG_REAL / "外部文件-大集中相关数据20261016.xlsx"
EXTERNAL_DEFAULT = CONFIG_REAL / "外部文件-大集中相关数据.xlsx"
S2_EXTERNAL = EXTERNAL_20261016 if EXTERNAL_20261016.is_file() else EXTERNAL_DEFAULT


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1 << 20), b""):
            digest.update(block)
    return digest.hexdigest()


def run_exe(args: list[str], env_extra: dict | None = None, timeout: int = 3600) -> dict:
    env = os.environ.copy()
    if env_extra:
        env.update(env_extra)
    started = time.monotonic()
    proc = subprocess.run(
        [str(EXE), *args], capture_output=True, text=True,
        encoding="utf-8", errors="replace", timeout=timeout, env=env,
    )
    return {
        "returncode": proc.returncode,
        "stdout": proc.stdout or "",
        "stderr": proc.stderr or "",
        "elapsed": round(time.monotonic() - started, 1),
    }


def dump_workbook(path: Path, max_rows: int | None = None) -> dict:
    """工作簿 → {sheet: {dims, rows:[[cell...]]}}（仅值；datetime 转 ISO 字符串）。"""
    from openpyxl import load_workbook

    wb = load_workbook(path, read_only=True, data_only=True)
    sheets = {}
    for name in wb.sheetnames:
        ws = wb[name]
        rows = []
        for index, row in enumerate(ws.iter_rows(values_only=True)):
            if max_rows is not None and index >= max_rows:
                break
            rows.append([
                value.isoformat() if hasattr(value, "isoformat") else value
                for value in row
            ])
        sheets[name] = {"dims": [ws.max_row, ws.max_column], "rows": rows}
    wb.close()
    return sheets


def diff_workbooks(left: dict, right: dict, label: str, limit: int = 20) -> dict:
    """逐格值对比；返回差异摘要（前 limit 条明细）。"""
    differences: list[str] = []
    for name in sorted(set(left) | set(right)):
        if name not in left or name not in right:
            differences.append(f"{label}｜工作表仅一侧存在：{name}")
            continue
        ls, rs = left[name], right[name]
        if ls["dims"] != rs["dims"]:
            differences.append(
                f"{label}｜{name}｜行列不一致：Rust={ls['dims']} 基线={rs['dims']}")
        for r, (lrow, rrow) in enumerate(zip(ls["rows"], rs["rows"])):
            width = max(len(lrow), len(rrow))
            for c in range(width):
                lv = lrow[c] if c < len(lrow) else None
                rv = rrow[c] if c < len(rrow) else None
                if lv != rv and not _num_equal(lv, rv):
                    if len(differences) < limit:
                        differences.append(
                            f"{label}｜{name}｜r{r + 2}c{c + 1}｜Rust={lv!r} 基线={rv!r}")
                    else:
                        differences.append("…（明细超限截断）")
                        return {"sheet_diffs": differences, "cell_diff_total": -1}
    # 行数不一致时补齐多出的行
    for name in sorted(set(left) & set(right)):
        lr, rr = len(left[name]["rows"]), len(right[name]["rows"])
        if lr != rr and len(differences) < limit:
            differences.append(f"{label}｜{name}｜行数 Rust={lr} 基线={rr}")
    return {"sheet_diffs": differences, "cell_diff_total": len(differences)}


def _num_equal(lv, rv) -> bool:
    if isinstance(lv, (int, float)) and isinstance(rv, (int, float)):
        return abs(float(lv) - float(rv)) <= 1e-9
    if lv is None and rv == "":
        return True
    if lv == "" and rv is None:
        return True
    if isinstance(lv, str) and isinstance(rv, (int, float)):
        try:
            return abs(float(lv) - float(rv)) <= 1e-9
        except ValueError:
            return False
    return False


def input_hashes(paths: list[Path]) -> dict[str, str]:
    return {str(p): sha(p) for p in paths if p.is_file()}


def collect_sources() -> list[Path]:
    sources = sorted(S1F1_INPUTS.glob("*.xls*")) + sorted(S1F1_TEMPLATES.glob("*.xls*"))
    sources += sorted(S1F2_DIR.glob("*/*.xlsx"))
    sources += [S1F1_DIR / "外部文件-大集中相关数据.xlsx", S1_HISTORY,
                S2_EXTERNAL, S1F2_TEMPLATE]
    sources += sorted(S2F1_DIR.rglob("*.xls*"))
    sources += sorted(S3_DIR.glob("*.csv")) + sorted(S3_DIR.glob("*.xls"))
    return [p for p in sources if p.is_file()]


# ---------------------------------------------------------------- S1F1

def prepare_state(state_dir: Path) -> None:
    """对齐引擎设置：自动引擎 + 条件格式提取来源=原始源文件 + 空白格纳入。"""
    state_dir.mkdir(parents=True, exist_ok=True)
    settings = {
        "calc_engine": "auto",
        "edit_engine": "auto",
        "preferences": {
            "s1ConditionalFormatSource": "source",
            "conditionalFormatIncludeBlankCells": True,
            "centralTolerance": 100.0,
            "centralTrimIndicatorNames": False,
            "centralAutoOpenResult": False,
        },
    }
    (state_dir / "settings.json").write_text(
        json.dumps(settings, ensure_ascii=False, indent=2), encoding="utf-8")


def norm_rust_issues(manifest_path: Path) -> list[dict]:
    data = json.loads(manifest_path.read_text(encoding="utf-8"))
    issues = []
    for file in data.get("files", []):
        source_name = Path(file.get("source_file") or file.get("source") or "").name
        for issue in file.get("formula_issues", []) + file.get("conditional_issues", []):
            issues.append({
                "source_file_name": source_name,
                "sheet_name": issue.get("sheet_name"),
                "severity": issue.get("severity"),
                "formula_cell": issue.get("formula_cell"),
                "message": issue.get("message"),
                "check_field": issue.get("check_field"),
                "target_value": "" if issue.get("target_value") in (None, "null")
                                else str(issue.get("target_value")),
            })
    return issues


def find_recovery_manifest(state_dir: Path) -> Path | None:
    candidates = sorted(state_dir.rglob("*.json"))
    for path in candidates:
        if path.name == "settings.json":
            continue
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            continue
        if isinstance(data, dict) and "files" in data and "batch_id" in data:
            return path
    return None


def run_s1f1(run_dir: Path) -> dict:
    state_dir = run_dir / "state"
    prepare_state(state_dir)
    output_dir = run_dir / "s1f1"
    result = run_exe(
        ["acceptance-run", "--manifest", str(MANIFEST), "--output-dir", str(output_dir)],
        env_extra={"AUDIT_STATE_DIR": str(state_dir)},
    )
    entry: dict = {"exit": result["returncode"], "elapsed": result["elapsed"]}
    report_path = output_dir / "acceptance-report.json"
    if report_path.is_file():
        report = json.loads(report_path.read_text(encoding="utf-8"))
        step = report["steps"][0]
        entry["key_counts"] = step.get("key_counts")
    manifest_path = find_recovery_manifest(state_dir)
    entry["recovery_manifest"] = str(manifest_path) if manifest_path else None
    if manifest_path:
        entry["rust_issues"] = norm_rust_issues(manifest_path)
    # 汇总核查表校验.xlsx 也留档供逐格复查
    for result_xlsx in output_dir.rglob("汇总核查表校验.xlsx"):
        entry["result_workbook"] = str(result_xlsx)
        break
    return entry


def compare_s1f1(baseline: dict, entries: list[dict]) -> dict:
    base_issues = []
    for batch in baseline["batches"]:
        base_issues.extend(batch["issues"])
    key_fields = ["sheet_name", "severity", "formula_cell", "message"]

    def key(issue: dict, file_first: bool) -> tuple:
        prefix = issue.get("source_file_name", "") if file_first else ""
        return (prefix, *[str(issue.get(f) or "") for f in key_fields])

    comparison: dict = {"baseline": {
        "totals": baseline["totals"], "issue_count": len(base_issues)}}
    for index, entry in enumerate(entries, 1):
        counts = entry.get("key_counts") or {}
        rust_issues = entry.get("rust_issues") or []
        rust_index = {key(i, True): i for i in rust_issues}
        base_index = {key(i, True): i for i in base_issues}
        only_rust = sorted(set(rust_index) - set(base_index))
        only_base = sorted(set(base_index) - set(rust_index))
        value_diffs = []
        for k in sorted(set(rust_index) & set(base_index)):
            a, b = rust_index[k], base_index[k]
            for field in ("check_field", "target_value"):
                av, bv = a.get(field), b.get(field)
                av = "" if av in (None, "null") else str(av)
                bv = "" if bv in (None, "null") else str(bv)
                if av != bv and not _num_equal(av, bv):
                    value_diffs.append(f"{k[0][:24]}…｜{k[1]}｜{k[3]}｜{field}：Rust={av!r} 基线={bv!r}")
        comparison[f"run{index}"] = {
            "exit": entry.get("exit"),
            "key_counts": counts,
            "rust_issue_count": len(rust_issues),
            "count_match": (
                counts.get("totalIssues") == len(base_issues)
                and counts.get("formulaIssues") == 3345
                and counts.get("conditionalIssues") == 257
            ) if counts else None,
            "only_rust": [str(k) for k in only_rust[:30]],
            "only_baseline": [str(k) for k in only_base[:30]],
            "only_rust_total": len(only_rust),
            "only_baseline_total": len(only_base),
            "value_field_diffs": value_diffs[:30],
            "value_field_diff_total": len(value_diffs),
            "elapsed": entry.get("elapsed"),
        }
    # 两次运行一致性
    if len(entries) >= 2:
        a = entries[0].get("rust_issues") or []
        b = entries[1].get("rust_issues") or []
        ka = sorted(key(i, True) for i in a)
        kb = sorted(key(i, True) for i in b)
        comparison["two_runs_identical"] = ka == kb
    return comparison


# ---------------------------------------------------------------- S1F2

def run_s1f2(run_dir: Path) -> dict:
    flat = run_dir / "s1f2_flat"
    flat.mkdir(parents=True, exist_ok=True)
    files = sorted(p for p in S1F2_DIR.glob("*/*.xlsx")
                   if not p.name.startswith("汇总"))
    for f in files:
        target = flat / f.name
        if not target.is_file():
            shutil.copy2(f, target)
    output = run_dir / "s1f2" / "汇总校验结果说明.xlsx"
    output.parent.mkdir(parents=True, exist_ok=True)
    result = run_exe([
        "s1-region-summary",
        "--template-workbook", str(S1F2_TEMPLATE),
        "--input-dir", str(flat),
        "--output-workbook", str(output),
        "--mapping-config", str(S1_HISTORY),
        "--history-workbook", str(S1_HISTORY),
        "--feature-name", "任务说明",
        "--feature-name", "业务说明",
        "--feature-name", "本地校验结果",
        "--recursive", "false",
    ])
    return {
        "exit": result["returncode"],
        "stdout_tail": result["stdout"][-500:],
        "stderr_tail": result["stderr"][-500:],
        "output": str(output) if output.is_file() else None,
        "elapsed": result["elapsed"],
    }


def compare_s1f2(baseline: dict, entries: list[dict]) -> dict:
    baseline_workbook = next((ALIGN / "_work" / "s1f2").glob("汇总校验结果说明_*.xlsx"))
    base_dump = dump_workbook(baseline_workbook)
    base_rows = sum(d["dims"][0] - 1 for d in base_dump.values())
    comparison: dict = {
        "baseline": {"input_files": baseline["input_files"],
                     "rows_total": baseline["rows_total"],
                     "dump_rows_total": base_rows,
                     "sheets": {k: v["dims"] for k, v in base_dump.items()}}}
    for index, entry in enumerate(entries, 1):
        record = {"exit": entry["exit"], "elapsed": entry["elapsed"]}
        output = entry.get("output")
        if output and Path(output).is_file():
            dump = dump_workbook(Path(output))
            record["sheets"] = {k: v["dims"] for k, v in dump.items()}
            record["rows_total"] = sum(d["dims"][0] - 1 for d in dump.values())
            record["workbook_diff"] = diff_workbooks(dump, base_dump, "S1F2")
        else:
            record["stderr_tail"] = entry.get("stderr_tail")
        comparison[f"run{index}"] = record
    if len(entries) >= 2 and all(e.get("output") for e in entries[:2]):
        a = dump_workbook(Path(entries[0]["output"]))
        b = dump_workbook(Path(entries[1]["output"]))
        comparison["two_runs_identical"] = a == b
    return comparison


# ---------------------------------------------------------------- S2F1

def run_s2f1(run_dir: Path) -> dict:
    output = run_dir / "s2f1"
    output.mkdir(parents=True, exist_ok=True)
    result = run_exe([
        "s2",
        "--current-dir", str(S2F1_DIR / "2026.06 季报"),
        "--previous-dir", str(S2F1_DIR / "2026.03 季报"),
        "--config", str(CONFIG_REAL / "2.报表采集系统_配置.xlsx"),
        "--output-dir", str(output),
        "--external", str(S2_EXTERNAL),
    ], timeout=5400)
    report = None
    for path in output.rglob("*.xlsx"):
        report = path
        break
    summary = None
    try:
        summary = json.loads(result["stdout"]).get("summary")
    except ValueError:
        pass
    return {
        "exit": result["returncode"],
        "summary": summary,
        "report": str(report) if report else None,
        "stdout_tail": result["stdout"][-400:],
        "stderr_tail": result["stderr"][-400:],
        "elapsed": result["elapsed"],
    }


def compare_s2f1(baseline: dict, entries: list[dict]) -> dict:
    baseline_workbook = next((ALIGN / "_work" / "s2f1").glob("报表审核清单_*.xlsx"))
    base_dump = dump_workbook(baseline_workbook)
    comparison: dict = {
        "baseline": {"findings_total": baseline["findings_total"],
                     "findings_by_status": baseline["findings_by_status"],
                     "sheets": {k: v["dims"] for k, v in base_dump.items()}}}
    for index, entry in enumerate(entries, 1):
        record = {"exit": entry["exit"], "summary": entry.get("summary"),
                  "elapsed": entry["elapsed"]}
        report = entry.get("report")
        if report and Path(report).is_file():
            dump = dump_workbook(Path(report))
            record["sheets"] = {k: v["dims"] for k, v in dump.items()}
            record["workbook_diff"] = diff_workbooks(dump, base_dump, "S2F1", limit=30)
        else:
            record["stderr_tail"] = entry.get("stderr_tail")
        comparison[f"run{index}"] = record
    if len(entries) >= 2 and all(e.get("report") for e in entries[:2]):
        a = dump_workbook(Path(entries[0]["report"]))
        b = dump_workbook(Path(entries[1]["report"]))
        comparison["two_runs_identical"] = a == b
    return comparison


# ---------------------------------------------------------------- S3

def run_s3f1(run_dir: Path) -> Path:
    output = run_dir / "s3f1" / "比较结果.xlsx"
    output.parent.mkdir(parents=True, exist_ok=True)
    result = run_exe([
        "s3-execution-compare",
        "--current-data", str(S3_DIR / "2026.10 月报二批中小法人查询.csv"),
        "--previous-data", str(S3_DIR / "2026.09 月报二批中小法人查询.csv"),
        "--common-config", str(CONFIG_REAL / "3.0大集中通用配置.xlsx"),
        "--feature-config", str(CONFIG_REAL / "3.1大集中执行比较_配置.xlsx"),
        "--output-workbook", str(output),
        "--expression-schema", "legacy-eight-segment",
    ])
    if result["returncode"] != 0:
        raise RuntimeError(f"S3F1 失败：{result['stderr'][-500:]}")
    return output


def count_fill_colors(path: Path) -> dict[str, int]:
    """统计第 2 行起“是否说明”判定列填充色分布（openpyxl fgColor.rgb→索引近似）。

    基线 fill_color 是 VBA/openpyxl ColorIndex；Rust 工作簿写入 RGB。
    这里按“值颜色对”统计并与基线的键集合映射后对比由报告层解释。
    """
    from openpyxl import load_workbook

    wb = load_workbook(path, read_only=False, data_only=True)
    counts: dict[str, int] = {}
    ws = wb[wb.sheetnames[0]]
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            if cell.fill and cell.fill.fill_type == "solid":
                rgb = cell.fill.fgColor.rgb if cell.fill.fgColor else None
                key = str(rgb)
            else:
                key = "None"
            counts[key] = counts.get(key, 0) + 1
    wb.close()
    return counts


def run_s3f2(comparison: Path, run_dir: Path) -> Path:
    output = run_dir / "s3f2" / "说明文件.xlsx"
    output.parent.mkdir(parents=True, exist_ok=True)
    result = run_exe([
        "s3-export-explanations",
        "--comparison-file", str(comparison),
        "--output-workbook", str(output),
        "--config-workbook", str(CONFIG_REAL / "3.0大集中通用配置.xlsx"),
    ])
    if result["returncode"] != 0:
        raise RuntimeError(f"S3F2 失败：{result['stderr'][-500:]}")
    return output


def run_s3f3(explanation: Path, run_dir: Path) -> dict:
    feedback_dir = run_dir / "s3f3"
    feedback_dir.mkdir(parents=True, exist_ok=True)
    feedback = feedback_dir / "2026.10 说明文件.xlsx"
    shutil.copy2(S3_DIR / "2026.10 说明文件.xls", feedback)
    result = run_exe([
        "s3-import-explanation-feedback",
        "--target-workbook", str(explanation),
        "--feedback-workbook", str(feedback),
    ])
    return {"exit": result["returncode"],
            "stdout_tail": result["stdout"][-600:],
            "stderr_tail": result["stderr"][-400:],
            "elapsed": result["elapsed"]}


def run_s3f4(run_dir: Path) -> Path:
    output = run_dir / "s3f4" / "核对结果.xlsx"
    output.parent.mkdir(parents=True, exist_ok=True)
    result = run_exe([
        "s3-check-current-values",
        "--current-csv", str(S3_DIR / "2026.10 月报二批中小法人查询.csv"),
        "--previous-csv", str(S3_DIR / "2026.09 月报二批中小法人查询.csv"),
        "--config-workbook", str(CONFIG_REAL / "3.2大集中本期数值核对_配置.xlsx"),
        "--output-workbook", str(output),
    ])
    if result["returncode"] != 0:
        raise RuntimeError(f"S3F4 失败：{result['stderr'][-500:]}")
    return output


def run_s3f5(comparison: Path, run_dir: Path) -> Path:
    output = run_dir / "s3f5"
    output.mkdir(parents=True, exist_ok=True)
    result = run_exe([
        "s3-generate-financial-forms",
        "--comparison-workbook", str(comparison),
        "--common-config", str(CONFIG_REAL / "3.0大集中通用配置.xlsx"),
        "--forms-config", str(CONFIG_REAL / "3.3大集中指标比较拆分_配置.xlsx"),
        "--output-directory", str(output),
    ], timeout=3600)
    if result["returncode"] != 0:
        raise RuntimeError(f"S3F5 失败：{result['stderr'][-500:]}")
    return output


def compare_s3(baseline: dict, name: str, entries: list[dict]) -> dict:
    comparison: dict = {"baseline": {
        k: baseline[k] for k in baseline if k not in ("environment",)}}
    for index, entry in enumerate(entries, 1):
        record = dict(entry)
        record.pop("workbook", None)
        if entry.get("workbook") and Path(entry["workbook"]).is_file():
            dump = dump_workbook(Path(entry["workbook"]))
            record["sheets"] = {k: v["dims"] for k, v in dump.items()}
            base_file = ALIGN / "_work" / {"s3f1": "s3f1/比较结果.xlsx",
                                            "s3f2": "s3f2/说明文件.xlsx",
                                            "s3f4": "s3f4/核对结果.xlsx"}.get(name, "")
            if base_file and (ALIGN / "_work" / base_file).is_file():
                base_dump = dump_workbook(ALIGN / "_work" / base_file)
                record["workbook_diff"] = diff_workbooks(dump, base_dump, name.upper(), limit=30)
        comparison[f"run{index}"] = record
    if len(entries) >= 2 and all(e.get("workbook") for e in entries[:2]):
        a = dump_workbook(Path(entries[0]["workbook"]))
        b = dump_workbook(Path(entries[1]["workbook"]))
        comparison["two_runs_identical"] = a == b
    return comparison


def compare_s3f3(baseline: dict, entries: list[dict]) -> dict:
    comparison: dict = {"baseline": baseline.get("result")}
    for index, entry in enumerate(entries, 1):
        comparison[f"run{index}"] = entry
    return comparison


def compare_s3f5(baseline: dict, entries: list[dict]) -> dict:
    base_names = sorted(f["name"] for f in baseline["rendered_files"])
    comparison: dict = {"baseline": {"count": len(base_names), "names": base_names}}
    for index, entry in enumerate(entries, 1):
        output = Path(entry.get("output", ""))
        files = sorted(output.rglob("*.xlsx")) if output.is_dir() else []
        names = sorted(f.name for f in files)
        record = {
            "count": len(names),
            "names": names,
            "names_match": names == base_names,
            "elapsed": entry.get("elapsed"),
        }
        # 逐文件值内容与基线对比（写入器不同哈希必不同）
        diffs = []
        for base in baseline["rendered_files"]:
            mine = output / base["name"]
            if not mine.is_file():
                diffs.append(f"缺少：{base['name']}")
                continue
            base_path = ALIGN / "_work" / "s3f5" / base["name"]
            if base_path.is_file():
                d = diff_workbooks(dump_workbook(mine), dump_workbook(base_path),
                                   f"S3F5/{base['name']}", limit=8)
                if d["sheet_diffs"]:
                    diffs.append(f"{base['name']}：{len(d['sheet_diffs'])} 处差异 {d['sheet_diffs'][:3]}")
        record["content_diffs"] = diffs[:20]
        comparison[f"run{index}"] = record
    if len(entries) >= 2:
        comparison["two_runs_identical"] = (
            entries[0].get("names") == entries[1].get("names"))
    return comparison


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--runs", type=int, default=2)
    parser.add_argument("--features", default="all")
    args = parser.parse_args()
    features = (["s1f1", "s1f2", "s2f1", "s3f1", "s3f2", "s3f3", "s3f4", "s3f5"]
                if args.features == "all" else args.features.split(","))

    sources_before = input_hashes(collect_sources())
    if REPORT.is_file():
        try:
            report = json.loads(REPORT.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            report = {}
    else:
        report = {}
    report["exe"] = str(EXE)
    report["exe_sha256"] = sha(EXE)
    report["engine_settings"] = ("calc=auto(Excel COM); s1ConditionalFormatSource=source; "
                                 "conditionalFormatEvaluator=PYTHON; "
                                 "conditionalFormatIncludeBlankCells=true; edit=auto")
    if "features" not in report:
        report["features"] = {}
    WORK.mkdir(parents=True, exist_ok=True)

    for feature in features:
        print(f"=== {feature} × {args.runs} runs ===", flush=True)
        entries = []
        for run_index in range(1, args.runs + 1):
            run_dir = WORK / f"{feature}_run{run_index}"
            run_dir.mkdir(parents=True, exist_ok=True)
            started = time.monotonic()
            if feature == "s1f1":
                entry = run_s1f1(run_dir)
            elif feature == "s1f2":
                entry = run_s1f2(run_dir)
            elif feature == "s2f1":
                entry = run_s2f1(run_dir)
            else:
                if feature == "s3f1":
                    workbook = run_s3f1(run_dir)
                    dump = dump_workbook(workbook)
                    entry = {"workbook": str(workbook),
                             "fills": count_fill_colors(workbook),
                             "rows": sum(d["dims"][0] - 1 for d in dump.values()),
                             "elapsed": round(time.monotonic() - started, 1)}
                elif feature in ("s3f2", "s3f4"):
                    comparison = run_s3f1(run_dir / ("cmp_for_" + feature))
                    workbook = (run_s3f2(comparison, run_dir) if feature == "s3f2"
                                else run_s3f4(run_dir))
                    entry = {"workbook": str(workbook),
                             "elapsed": round(time.monotonic() - started, 1)}
                elif feature == "s3f3":
                    comparison = run_s3f1(run_dir / "cmp_for_s3f3")
                    explanation = run_s3f2(comparison, run_dir / "s3f2_for_f3")
                    entry = run_s3f3(explanation, run_dir)
                else:  # s3f5
                    comparison = run_s3f1(run_dir / "cmp_for_s3f5")
                    output = run_s3f5(comparison, run_dir)
                    entry = {"output": str(output),
                             "elapsed": round(time.monotonic() - started, 1)}
            entries.append(entry)
            print(f"  run{run_index}: {round(time.monotonic() - started, 1)}s "
                  f"exit={entry.get('exit')}", flush=True)

        baseline = json.loads((ALIGN / f"{feature.upper()}.json").read_text(encoding="utf-8"))
        if feature == "s1f1":
            comparison = compare_s1f1(baseline, entries)
        elif feature == "s1f2":
            comparison = compare_s1f2(baseline, entries)
        elif feature == "s2f1":
            comparison = compare_s2f1(baseline, entries)
        elif feature == "s3f3":
            comparison = compare_s3f3(baseline, entries)
        elif feature == "s3f5":
            comparison = compare_s3f5(baseline, entries)
        else:
            comparison = compare_s3(baseline, feature, entries)
        report["features"][feature] = comparison
        (REPORT.parent.mkdir(parents=True, exist_ok=True))
        REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2, default=str),
                          encoding="utf-8")
        print(f"  report updated: {REPORT}", flush=True)

    sources_after = input_hashes(collect_sources())
    changed = [p for p in sources_before if sources_before[p] != sources_after.get(p)]
    report["sources_unchanged"] = not changed
    report["changed_sources"] = changed
    REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2, default=str),
                      encoding="utf-8")
    print(f"sources unchanged: {not changed} ({len(sources_before)} files)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
