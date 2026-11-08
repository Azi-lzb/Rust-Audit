#!/usr/bin/env python3
"""S3-F05 编辑引擎 solo 验证：ooxml 强制独跑 ×2 + openpyxl 对照 ×1 + Python 基线逐格对比。

背景：auto 路由为 ooxml 优先、openpyxl 静默回退——ooxml 内部失败会被掩盖。
本脚本用 `edit_engine=ooxml` 强制只注册 ooxml 后端（fail-closed，无回退），
跑通即证明 ooxml 真实执行了整批编辑（DeleteColumn/SetCellValue/…）。

判定（全部满足才通过）：
1. ooxml 两跑均 exit 0、8 份表单、两跑逐格一致；
2. ooxml 与 openpyxl 强制跑逐格 0 差异；
3. ooxml 与 Python 基线（3.open-data-audit _work/s3f5）逐格 0 差异；
4. 输入（比较结果 + 两份配置）SHA-256 全程不变。

只读原始文件；全部产物在 probe/reports/s3f5_engine_solo_work。
"""
from __future__ import annotations

import hashlib
import json
import shutil
import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(r"C:\Users\AZI\Desktop\VibeCode\Rust-Audit")
sys.path.insert(0, str(ROOT / "tests" / "compare"))
from rustaudit_alignment_p1 import (  # noqa: E402
    CONFIG_REAL,
    EXE,
    S3_DIR,
    dump_workbook,
    diff_workbooks,
    run_exe,
)

OLD = Path(r"C:\Users\AZI\Desktop\VibeCode\3.open-data-audit")
BASELINE_DIR = OLD / "test-output" / "rustaudit_alignment" / "_work" / "s3f5"
S3_DIR = OLD / "test-data" / "S3"
WORK = ROOT / "probe" / "reports" / "s3f5_engine_solo_work"
REPORT = ROOT / "probe" / "reports" / "s3f5_engine_solo_report.json"


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1 << 20), b""):
            digest.update(block)
    return digest.hexdigest()


def write_state(state_dir: Path, edit_engine: str) -> None:
    state_dir.mkdir(parents=True, exist_ok=True)
    (state_dir / "settings.json").write_text(
        json.dumps({"calc_engine": "auto", "edit_engine": edit_engine,
                    "preferences": {}}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def run_s3f1_comparison() -> Path:
    output = WORK / "cmp" / "比较结果.xlsx"
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


def run_forms(comparison: Path, out_dir: Path, state_dir: Path) -> dict:
    out_dir.mkdir(parents=True, exist_ok=True)
    import os

    env = os.environ.copy()
    env["AUDIT_STATE_DIR"] = str(state_dir)
    started = time.monotonic()
    proc = subprocess.run(
        [str(EXE), "s3-generate-financial-forms",
         "--comparison-workbook", str(comparison),
         "--common-config", str(CONFIG_REAL / "3.0大集中通用配置.xlsx"),
         "--forms-config", str(CONFIG_REAL / "3.3大集中指标比较拆分_配置.xlsx"),
         "--output-directory", str(out_dir)],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        timeout=3600, env=env,
    )
    return {
        "exit": proc.returncode,
        "stderr_tail": (proc.stderr or "")[-400:],
        "elapsed": round(time.monotonic() - started, 1),
        "files": sorted(p.name for p in out_dir.rglob("*.xlsx")),
    }


def compare_set(left_dir: Path, right_dir: Path, label: str) -> tuple[list[str], list[str]]:
    """逐格值差异（必须为 0）与 dimension 元数据差异（未触碰空表的字节保留语义）分开。"""
    cell_diffs: list[str] = []
    dim_notes: list[str] = []
    # 发布目录带时间戳子目录，按文件名索引
    left_files = {p.name: p for p in left_dir.rglob("*.xlsx")}
    right_files = {p.name: p for p in right_dir.rglob("*.xlsx")}
    for name in sorted(left_files):
        other = right_files.get(name)
        if other is None:
            cell_diffs.append(f"{label} 缺少：{name}")
            continue
        result = diff_workbooks(dump_workbook(left_files[name]), dump_workbook(other),
                                f"{label}/{name}", limit=40)
        for line in result["sheet_diffs"]:
            if "行列不一致" in line:
                dim_notes.append(f"{label} {name}：{line}")
            else:
                cell_diffs.append(f"{label} {name}：{line}")
    return cell_diffs, dim_notes


def main() -> int:
    if sys.stdout.encoding and sys.stdout.encoding.lower() not in ("utf-8", "utf8"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    if WORK.exists():
        shutil.rmtree(WORK)
    WORK.mkdir(parents=True)

    # 先生成 S3F1 比较结果（S3F5 的输入，与对齐基线同口径）
    comparison = run_s3f1_comparison()
    inputs = [comparison,
              CONFIG_REAL / "3.0大集中通用配置.xlsx",
              CONFIG_REAL / "3.3大集中指标比较拆分_配置.xlsx"]
    before = {str(p): sha(p) for p in inputs}

    runs: dict[str, dict] = {}
    for label, engine in (("ooxml_run1", "ooxml"), ("ooxml_run2", "ooxml"),
                          ("openpyxl_run1", "openpyxl")):
        state = WORK / f"state_{label}"
        write_state(state, engine)
        result = run_forms(comparison, WORK / label, state)
        result["engine"] = engine
        runs[label] = result
        print(f"{label}: exit={result['exit']} files={len(result['files'])} "
              f"elapsed={result['elapsed']}s", flush=True)
        if result["exit"] != 0:
            print(f"  stderr: {result['stderr_tail']}")

    payload: dict = {"exe": str(EXE), "exe_sha256": sha(EXE), "runs": runs}

    all_ok = all(runs[label]["exit"] == 0 and len(runs[label]["files"]) == 8
                 for label in runs)
    if all_ok:
        d_ooxml, n_ooxml = compare_set(WORK / "ooxml_run1", WORK / "ooxml_run2", "ooxml两跑")
        d_cross, n_cross = compare_set(WORK / "ooxml_run1", WORK / "openpyxl_run1", "ooxml-vs-openpyxl")
        d_baseline, n_baseline = compare_set(WORK / "ooxml_run1", BASELINE_DIR, "ooxml-vs-python基线")
        payload["ooxml_two_run_cell_diffs"] = d_ooxml
        payload["ooxml_vs_openpyxl_cell_diffs"] = d_cross
        payload["ooxml_vs_python_baseline_cell_diffs"] = d_baseline
        payload["dimension_metadata_notes"] = {
            "ooxml_two_run": n_ooxml,
            "ooxml_vs_openpyxl": n_cross,
            "ooxml_vs_python_baseline": n_baseline,
        }
        all_ok = not d_ooxml and not d_cross and not d_baseline
        print(f"ooxml 两跑逐格差异: {len(d_ooxml)}（dimension 元数据备注 {len(n_ooxml)}）")
        print(f"ooxml vs openpyxl 逐格差异: {len(d_cross)}（元数据备注 {len(n_cross)}）")
        print(f"ooxml vs Python 基线逐格差异: {len(d_baseline)}（元数据备注 {len(n_baseline)}）")
        for line in (d_ooxml + d_cross + d_baseline)[:6]:
            print(f"  {line}")
        for line in n_baseline[:4]:
            print(f"  [元数据] {line[:140]}")

    after = {str(p): sha(p) for p in inputs}
    unchanged = before == after
    payload["inputs_unchanged"] = unchanged
    print(f"输入 SHA-256 不变: {unchanged}")

    REPORT.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"报告：{REPORT}")
    print(f"=== S3-F05 ooxml solo 验证：{'通过' if all_ok and unchanged else '未通过'} ===")
    return 0 if all_ok and unchanged else 1


if __name__ == "__main__":
    sys.exit(main())
