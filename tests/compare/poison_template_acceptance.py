#!/usr/bin/env python3
"""毒模板（自带外部链接的报送模板）专项验收。

背景：`！70.金融基础数据采集（债券业务）202608.xlsx` 的校验区域公式原文
自带 `[1]集中系统数据!` / `[1]参照表!` 外部工作簿引用前缀（模板内
`externalLinks/externalLink1.xml` 指向旧机器绝对路径，属断链但靠 `<v>` 缓存值
兜底，模板自身可打开）。审核副本会把外部表导入为本地 sheet，若保留公式
`[1]` 前缀而丢弃 externalLink 定义，公式将悬空 → Excel COM 拒绝打开
（0x80020009）、WPS 重算丢公式。

本验收要求（毒模板产物，7 本债券报送文件）：
1. 走 S1-F01 正式流程生成审核副本（不经任何手工修补）；
2. 每个副本中指向本地表的 `[N]表名!` 残留为 0（修复在管线内生效）；
3. 副本 Excel COM 只读打开 7/7 成功；
4. 原始报送文件与毒模板 SHA-256 全程不变；
5. 对照组：把 `[N]` 前缀人为写回副本后，Excel COM 必须打不开（证明判定有效）。

只读原始文件；所有中间产物在临时目录。
"""
from __future__ import annotations

import hashlib
import json
import re
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

ROOT = Path(r"C:\Users\AZI\Desktop\VibeCode\Rust-Audit")
EXE = (ROOT / "dist" / "open-data-audit.exe").resolve()
OLD = Path(r"C:\Users\AZI\Desktop\VibeCode\3.open-data-audit")
TEMPLATES = ROOT / "测试数据" / "S1F1_新旧对比_20261105_231620" / "business_templates"
MANIFEST = ROOT / "测试数据" / "S1F1_新旧对比_20261105_231620" / "rust-manifest-with-external.json"
SOURCES = ROOT / "测试数据" / "逐笔统计系统" / "核查表"
REPORT = ROOT / "probe" / "reports" / "poison_template_acceptance.json"
WORK = ROOT / "probe" / "reports" / "poison_template_work"

LOCAL_REF = re.compile(r"\[\d+\]('?[^'!\[\]]+'?!)")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1 << 20), b""):
            digest.update(block)
    return digest.hexdigest()


def poison_templates() -> list[Path]:
    """带 externalLink 部件的业务模板即毒模板。"""
    found = []
    for path in sorted(TEMPLATES.glob("*.xlsx")):
        with zipfile.ZipFile(path) as archive:
            names = archive.namelist()
        if any("externalLink" in name for name in names):
            found.append(path)
    return found


def dangling_refs(path: Path) -> int:
    """统计工作簿中指向本地表的 [N] 外链残留（应恒为 0）。"""
    total = 0
    with zipfile.ZipFile(path) as archive:
        sheet_names = set()
        workbook = archive.read("xl/workbook.xml").decode("utf-8", errors="ignore")
        sheet_names.update(re.findall(r'<sheet [^>]*name="([^"]+)"', workbook))
        for name in archive.namelist():
            if not (name.startswith("xl/worksheets/") and name.endswith(".xml")):
                continue
            text = archive.read(name).decode("utf-8", errors="ignore")
            for match in LOCAL_REF.finditer(text):
                target = match.group(1).strip("'!").replace("''", "'")
                if target in sheet_names:
                    total += 1
    return total


def excel_com_open_batch(paths: list[Path], work: Path, label: str) -> dict[str, str]:
    context_file = work / f"{label}_context.json"
    result_file = work / f"{label}_result.json"
    context_file.write_text(
        json.dumps({"files": [str(p) for p in paths], "out": str(result_file)}, ensure_ascii=False),
        encoding="utf-8",
    )
    ps = work / f"{label}.ps1"
    ps.write_text(
        r'''
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
'''.replace("__CTX__", str(context_file)),
        encoding="utf-8-sig",
    )
    proc = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(ps)],
        capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=900,
    )
    if "OPEN-BATCH-OK" not in (proc.stdout or ""):
        raise RuntimeError((proc.stderr or proc.stdout)[-400:])
    raw = json.loads(result_file.read_text(encoding="utf-8-sig"))
    return {k.replace("/", "\\"): v for k, v in raw.items()}


def main() -> int:
    if sys.stdout.encoding and sys.stdout.encoding.lower() not in ("utf-8", "utf8"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    poisons = poison_templates()
    if not poisons:
        print("未发现毒模板；验收不适用。")
        return 2
    if WORK.exists():
        shutil.rmtree(WORK)
    WORK.mkdir(parents=True)

    source_dir = WORK / "sources"
    source_dir.mkdir()
    templates_dir = WORK / "templates"
    templates_dir.mkdir()

    # 用毒模板对应的源文件（文件名含模板特征词）构造单批次清单
    manifest = json.loads(MANIFEST.read_text(encoding="utf-8-sig"))
    s1 = manifest["s1"]
    poison_names = {p.name for p in poisons}
    selected: list[Path] = []
    for template in poisons:
        # 模板名形如 ！70.金融基础数据采集（债券业务）202608.xlsx → 特征词“债券业务”
        match = re.search(r"[.．]([^0-9]+?)\d{6}\.xlsx$", template.name)
        keyword = match.group(1) if match else ""
        for source in sorted(SOURCES.rglob("*.xlsx")):
            if keyword and keyword in source.name:
                selected.append(source)
        shutil.copy2(template, templates_dir / template.name)
    if not selected:
        print("毒模板未匹配到源文件；验收不适用。")
        return 2
    for source in selected:
        shutil.copy2(source, source_dir / source.name)

    before = {str(p): sha256(p) for p in selected}
    before.update({str(p): sha256(p) for p in poisons})

    # S1-F01 正式流程（副本修复在管线内，无手工修补）
    state_dir = WORK / "state"
    state_dir.mkdir()
    (state_dir / "settings.json").write_text(
        json.dumps(
            {
                "calc_engine": "auto",
                "edit_engine": "auto",
                "preferences": {
                    "s1ConditionalFormatSource": "source",
                    "conditionalFormatEvaluator": "PYTHON",
                },
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    run_manifest = {
        "s1": {
            **s1,
            "sourceDirectory": str(source_dir),
            "templateDirectory": str(templates_dir),
            "sidecar": None,
        }
    }
    manifest_file = WORK / "manifest.json"
    manifest_file.write_text(json.dumps(run_manifest, ensure_ascii=False, indent=1), encoding="utf-8")
    output_dir = WORK / "out"
    import os

    run_env = os.environ.copy()
    run_env["AUDIT_STATE_DIR"] = str(state_dir)
    proc = subprocess.run(
        [str(EXE), "acceptance-run", "--manifest", str(manifest_file), "--output-dir", str(output_dir)],
        capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=3600, env=run_env,
    )
    report = output_dir / "acceptance-report.json"
    if not report.is_file():
        print(f"S1-F01 未产出验收报告：{proc.stderr[-400:]}")
        return 1
    step = json.loads(report.read_text(encoding="utf-8"))["steps"][0]
    counts = step.get("key_counts") or {}

    # 收集审核副本
    copies = sorted(
        p for p in output_dir.rglob("*-recalculated.xlsx")
    )
    print(f"毒模板: {[p.name for p in poisons]}")
    print(f"源文件: {len(selected)} 本；审核副本: {len(copies)} 个")
    print(f"批次统计: 文件 {counts.get('filesTotal')}，结构通过 {counts.get('structurePassed')}，"
          f"公式异常 {counts.get('formulaIssues')}，条件格式 {counts.get('conditionalIssues')}")

    # 1) 副本 [N] 残留必须为 0
    residual = {str(p): dangling_refs(p) for p in copies}
    residual_total = sum(residual.values())

    # 2) Excel COM 打开毒模板副本
    open_results = excel_com_open_batch(copies, WORK, "poison_copies")
    opened = sum(1 for p in copies if open_results.get(str(p), "").startswith("OK"))

    # 3) 对照组：把 [N] 写回一个副本，Excel 必须打不开
    control_ok = None
    if copies:
        control = WORK / "control-with-external-ref.xlsx"
        with zipfile.ZipFile(copies[0]) as src, zipfile.ZipFile(control, "w", zipfile.ZIP_DEFLATED) as dst:
            for name in src.namelist():
                data = src.read(name)
                if name.startswith("xl/worksheets/") and name.endswith(".xml"):
                    text = data.decode("utf-8", errors="ignore")
                    if "集中系统数据!" in text:
                        text = text.replace("集中系统数据!", "[1]集中系统数据!", 1)
                        data = text.encode("utf-8")
                dst.writestr(name, data)
        control_result = excel_com_open_batch([control], WORK, "control")
        control_ok = not control_result.get(str(control), "").startswith("OK")
        print(f"对照组（写回 [1] 外链）: {'打不开 ✓' if control_ok else '仍能打开 ✗'}")

    # 4) 原始文件哈希不变
    after = {str(p): sha256(p) for p in selected}
    after.update({str(p): sha256(p) for p in poisons})
    unchanged = before == after

    print(f"副本 [N] 残留: {residual_total}（要求 0）")
    print(f"Excel COM 打开: {opened}/{len(copies)}")
    print(f"原始文件与模板哈希不变: {unchanged}")

    payload = {
        "poison_templates": [p.name for p in poisons],
        "source_files": [p.name for p in selected],
        "copies": len(copies),
        "batch_counts": counts,
        "dangling_external_refs": residual,
        "dangling_external_refs_total": residual_total,
        "excel_open": {"opened": opened, "total": len(copies), "detail": open_results},
        "control_original_ref_fails": control_ok,
        "inputs_unchanged": unchanged,
    }
    REPORT.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"报告：{REPORT}")

    ok = (
        residual_total == 0
        and opened == len(copies)
        and copies
        and control_ok is True
        and unchanged
    )
    print(f"=== 毒模板验收：{'通过' if ok else '未通过'} ===")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
