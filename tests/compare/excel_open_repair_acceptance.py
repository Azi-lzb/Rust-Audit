#!/usr/bin/env python3
"""Excel COM 打开修复验收：28 本 Excel 打不开的报送文件。

对每本失败输入的副本运行 Rust `excel-open-repair`（与旧 Python Audit
excel_open_repair.py 逐条一致的修复），要求：
1. 修复统计与旧实现的验证结果逐本一致（补引号数 / 修窗格数）；
2. 修复副本用 Excel COM 只读打开 28/28 成功；
3. 对照组：原样副本抽样打开必须仍失败（证明判定有效）；
4. 负向控制：无问题文件抽样修复必须 changed=false 且字节不变。

全程只读原始文件；所有修复在临时副本上完成并验证。
"""
from __future__ import annotations

import hashlib
import json
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(r"C:\Users\AZI\Desktop\VibeCode\Rust-Audit")
EXE = (ROOT / "target" / "debug" / "open-data-audit.exe").resolve()
OLD = Path(r"C:\Users\AZI\Desktop\VibeCode\3.open-data-audit")
EXPECT = OLD / "test-output" / "s1f1_mode_parity_20261107" / "excel_open_fix_28.json"
INPUTS = OLD / "test-output" / "s1f1_com_isolation_20261107" / "source_bundle" / "inputs_129"
REPORT = ROOT / "probe" / "reports" / "excel_open_repair_acceptance.json"
WORK = ROOT / "probe" / "reports" / "excel_open_repair_work"

CONTROL_COUNT = 6          # 原样副本打开对照（预期失败）数
CLEAN_COUNT = 12           # 无问题文件负向控制数


def run_repair(workbook: Path) -> dict:
    result = subprocess.run(
        [str(EXE), "excel-open-repair", "--workbook", str(workbook)],
        capture_output=True, text=True, encoding="utf-8", timeout=300,
    )
    if result.returncode != 0:
        return {"changed": False, "quoted": 0, "pane_fixed": 0,
                "reason": f"exit={result.returncode} {result.stderr[-200:]}"}
    return json.loads(result.stdout)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1 << 20), b""):
            digest.update(block)
    return digest.hexdigest()


def excel_com_open_batch(paths: list[Path], work: Path, label: str) -> dict[str, str]:
    """Excel COM 单实例串行只读打开，返回 {路径: OK sheets=N | FAIL 说明}。"""
    context_file = work / f"{label}_context.json"
    scan_out = work / f"{label}_result.json"
    context_file.write_text(json.dumps(
        {"files": [str(path) for path in paths], "out": str(scan_out)},
        ensure_ascii=False), encoding="utf-8")
    ps = work / f"{label}_open.ps1"
    ps.write_text(r'''
$ErrorActionPreference = "Stop"
$ctx = Get-Content -Raw -Encoding UTF8 "__CTX__" | ConvertFrom-Json
$app = New-Object -ComObject Excel.Application
$app.Visible = $false
$app.DisplayAlerts = $false
$result = @{}
try {
    foreach ($path in $ctx.files) {
        try {
            $wb = $app.Workbooks.Open($path, 0, $true)
            $result[$path] = "OK sheets=" + $wb.Worksheets.Count
            $wb.Close($false)
        } catch {
            $msg = $_.Exception.Message
            $result[$path] = "FAIL " + $msg.Substring(0, [Math]::Min(120, $msg.Length))
        }
    }
    $json = $result | ConvertTo-Json -Depth 3
    [System.IO.File]::WriteAllText($ctx.out, $json, (New-Object System.Text.UTF8Encoding($true)))
    Write-Output "OPEN-BATCH-OK"
} finally {
    $app.Quit()
}
'''.replace("__CTX__", str(context_file)), encoding="utf-8-sig")
    proc = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(ps)],
        capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=900,
    )
    if "OPEN-BATCH-OK" not in (proc.stdout or ""):
        raise RuntimeError((proc.stderr or proc.stdout)[-400:])
    raw = json.loads(scan_out.read_text(encoding="utf-8-sig"))
    return {k.replace("/", "\\"): v for k, v in raw.items()}


def main() -> int:
    if sys.stdout.encoding and sys.stdout.encoding.lower() not in ("utf-8", "utf8"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if WORK.exists():
        shutil.rmtree(WORK)
    WORK.mkdir(parents=True)
    broken_dir = WORK / "broken"
    clean_dir = WORK / "clean"
    original_dir = WORK / "original"
    broken_dir.mkdir()
    clean_dir.mkdir()
    original_dir.mkdir()

    expectations = json.loads(EXPECT.read_text(encoding="utf-8"))
    broken_names = [entry["file"] for entry in expectations]

    # ---- 1) 28 本失败输入：修复 + 统计比对 ----
    repaired: list[dict] = []
    stats_mismatch: list[dict] = []
    for entry in expectations:
        name = entry["file"]
        copy = broken_dir / name
        shutil.copy2(INPUTS / name, copy)
        report = run_repair(copy)
        record = {
            "file": name,
            "expected": {"quoted": entry["quoted"], "pane_fixed": entry["pane_fixed"]},
            "rust": {"changed": report.get("changed"), "quoted": report.get("quoted"),
                     "pane_fixed": report.get("pane_fixed"), "reason": report.get("reason")},
            "stats_match": (report.get("changed") is True
                            and report.get("quoted") == entry["quoted"]
                            and report.get("pane_fixed") == entry["pane_fixed"]),
        }
        repaired.append(record)
        if not record["stats_match"]:
            stats_mismatch.append(name)
        mark = "OK  " if record["stats_match"] else "DIFF"
        print(f"{mark} {name[:46]} | 期望 引号={entry['quoted']} 窗格={entry['pane_fixed']}"
              f" | Rust 引号={report.get('quoted')} 窗格={report.get('pane_fixed')}", flush=True)

    # ---- 2) 修复副本 Excel COM 打开（28/28） ----
    print("\nExcel COM 打开修复副本 …", flush=True)
    open_results = excel_com_open_batch(
        [broken_dir / entry["file"] for entry in expectations], WORK, "repaired")
    opened = 0
    for entry in expectations:
        key = str(broken_dir / entry["file"])
        verdict = open_results.get(key, "MISSING")
        entry_ok = verdict.startswith("OK")
        opened += entry_ok
        for record in repaired:
            if record["file"] == entry["file"]:
                record["excel_open"] = verdict
        print(f"{'OK  ' if entry_ok else 'FAIL'} {entry['file'][:46]} | {verdict}", flush=True)

    # ---- 3) 对照组：原样副本抽样打开，必须仍失败 ----
    control_names = broken_names[:CONTROL_COUNT]
    for name in control_names:
        shutil.copy2(INPUTS / name, original_dir / name)
    original_results = excel_com_open_batch(
        [original_dir / name for name in control_names], WORK, "original")
    control_fail = sum(1 for name in control_names
                       if not original_results.get(str(original_dir / name), "OK").startswith("OK"))
    print(f"\n对照组原样副本：{control_fail}/{len(control_names)} 仍打不开（预期全部失败）", flush=True)

    # ---- 4) 负向控制：无问题文件必须不重写 ----
    all_inputs = sorted(INPUTS.glob("*.xlsx"))
    broken_set = set(broken_names)
    clean_names = [path.name for path in all_inputs if path.name not in broken_set][:CLEAN_COUNT]
    clean_reports: list[dict] = []
    unchanged = 0
    for name in clean_names:
        copy = clean_dir / name
        shutil.copy2(INPUTS / name, copy)
        before = sha256(copy)
        report = run_repair(copy)
        same = sha256(copy) == before and report.get("changed") is False
        unchanged += same
        clean_reports.append({"file": name, "rust": report, "byte_identical": same})
        print(f"{'OK  ' if same else 'DIFF'} {name[:46]} | changed={report.get('changed')}"
              f" reason={report.get('reason', '')[:40]}", flush=True)
    print(f"\n负向控制：{unchanged}/{len(clean_names)} 字节不变且 changed=false", flush=True)

    report_out = {
        "repair_stats": repaired,
        "repaired_excel_open": {"opened": opened, "total": len(expectations)},
        "original_control": {"still_failed": control_fail, "total": len(control_names)},
        "clean_control": {"unchanged": unchanged, "total": len(clean_names),
                          "details": clean_reports},
    }
    REPORT.write_text(json.dumps(report_out, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\n报告：{REPORT}")
    print(f"=== 验收结论：修复统计 {'一致' if not stats_mismatch else '存在差异'}；"
          f"Excel COM 打开 {opened}/{len(expectations)} ===")
    ok = (not stats_mismatch and opened == len(expectations)
          and control_fail == len(control_names) and unchanged == len(clean_names))
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
