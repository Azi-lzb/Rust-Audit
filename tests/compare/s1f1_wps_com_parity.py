#!/usr/bin/env python3
"""Rust 求值 vs WPS COM 渲染（4 类含 CF 的真实业务副本，正确激活刷新配方）。"""
from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

EXE = Path(r'C:\Users\AZI\Desktop\VibeCode\Rust-Audit\dist\OpenDataAudit\open-data-audit.exe')
RUN_DIR = Path(r'F:/RustBuild/audit/tmp/s1f1-auto-route-current-release-129-20261106/acceptance/steps/S1-F01')
REPORT_DIR = Path(r'C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports')

# 四类含 CF 的代表性副本（每类一本）
SAMPLES = [
    ('债券业务', '0012-TCL科技集团财务有限公司_金融基础数据采集_债券业务__2026-08-31_数据核对表-recalculated.xlsx'),
    ('单位存款', '0013-TCL科技集团财务有限公司_金融基础数据采集_单位存款__2026-08-31_数据核对表-recalculated.xlsx'),
    ('委托贷款', '0016-TCL科技集团财务有限公司_金融基础数据采集_委托贷款__2026-08-31_数据核对表-recalculated.xlsx'),
    ('同业业务', '0015-TCL科技集团财务有限公司_金融基础数据采集_同业业务__2026-08-31_数据核对表-recalculated.xlsx'),
]

PS_TEMPLATE = r'''
$ErrorActionPreference = "Stop"
$context = Get-Content -Raw -Encoding UTF8 "__CONTEXT_JSON__" | ConvertFrom-Json

# 启动 WPS 表格（单实例，非只读）
$app = New-Object -ComObject KET.Application
$app.Visible = $false
$app.DisplayAlerts = $false

$out = @{}
try {
    foreach ($entry in $context.files) {
        $wb = $app.Workbooks.Open($entry.path)
        try {
            # 关键：激活工作簿 + 开启屏幕更新，让 WPS 渲染条件格式
            $app.ScreenUpdating = $true
            $wb.Activate()

            $cfSheets = @()
            foreach ($sheetEntry in $entry.sheets) {
                $sheet = $wb.Worksheets.Item($sheetEntry.name)

                # 逐表激活 + 重算（旧 Python Audit 的 WPS 刷新配方）
                $sheet.Activate()
                $app.CalculateFullRebuild()

                $triggered = @()
                $scanned = 0
                foreach ($condition in $sheet.UsedRange.FormatConditions) {
                    $appliesTo = $condition.AppliesTo
                    foreach ($area in $appliesTo.Areas) {
                        foreach ($cell in $area.Cells) {
                            $scanned++
                            $addr = $cell.Address($false,$false)
                            try {
                                $display = $cell.DisplayFormat.Interior.Color
                                $base = $cell.Interior.Color
                                if ($display -ne $base) {
                                    $triggered += ("$addr=$display")
                                }
                            } catch { }
                        }
                    }
                }
                $cfSheets += @{
                    name = $sheetEntry.name
                    scanned = $scanned
                    triggered = $triggered
                }
            }
            $out[$entry.path] = @{ cfSheets = $cfSheets; error = "" }
        } catch {
            $out[$entry.path] = @{ error = $_.Exception.Message; cfSheets = @() }
        } finally {
            $wb.Close($false)
        }
    }
    $json = $out | ConvertTo-Json -Depth 6
    [System.IO.File]::WriteAllText($context.scanOut, $json, (New-Object System.Text.UTF8Encoding($true)))
    Write-Output "WPS-BATCH-OK"
} finally {
    $app.Quit()
    $app = $null
    [System.GC]::Collect()
}
'''


def main() -> int:
    if sys.stdout.encoding and sys.stdout.encoding.lower() not in ('utf-8', 'utf8'):
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    work_dir = REPORT_DIR / 's1f1_wps_parity_work'
    work_dir.mkdir(parents=True, exist_ok=True)

    # 1) 构造 manifest：文件路径 + 带 CF 的工作表名
    files = []
    for label, filename in SAMPLES:
        path = next(RUN_DIR.rglob(filename), None)
        if not path.is_file():
            print(f'文件不存在: {path.name[:50]}')
            continue
        import zipfile
        with zipfile.ZipFile(path) as archive:
            workbook_xml = archive.read('xl/workbook.xml').decode('utf-8', errors='ignore')
            rels_xml = archive.read('xl/_rels/workbook.xml.rels').decode('utf-8', errors='ignore')
            rel_map = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels_xml)) if False else {}
            sheet_entries = re.findall(r'<sheet name="([^"]+)"[^>]*r:id="(rId\d+)"', workbook_xml)
            rel_map = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels_xml))
            cf_sheets = []
            for name, rel_id in sheet_entries:
                target = rel_map.get(rel_id, '').lstrip('/')
                part = target if target.startswith('xl/') else 'xl/' + target
                try:
                    sheet_xml = archive.read(part).decode('utf-8', errors='ignore')
                    if '<conditionalFormatting' in sheet_xml:
                        cf_sheets.append({'name': name})
                except KeyError:
                    continue
            if cf_sheets:
                files.append({'path': str(path), 'sheets': cf_sheets, 'label': label})
                print(f'  {label}: {len(cf_sheets)} 张带 CF 的表')

    # 2) Rust 求值
    rust_results = {}
    for entry in files:
        proc = subprocess.run(
            [str(EXE), 'conditional-evaluate', '--workbook', entry['path']],
            capture_output=True, text=True, encoding='utf-8', timeout=300,
        )
        if proc.returncode == 0:
            data = json.loads(proc.stdout)
            triggered = {(e['sheet'], e['cell']) for e in data.get('triggered', [])}
            rust_results[entry['path']] = triggered
            print(f'  Rust {entry["label"]}: {len(triggered)} triggered')
        else:
            print(f'  Rust {entry["label"]}: ERROR {proc.stderr[:80]}')

    # 3) WPS COM 扫描
    context_file = work_dir / 'wps_context.json'
    scan_out = work_dir / 'wps_scan.json'
    context_file.write_text(json.dumps({
        'files': [{'path': entry['path'], 'sheets': entry['sheets']} for entry in files],
        'scanOut': str(scan_out),
    }, ensure_ascii=False), encoding='utf-8')

    ps_path = work_dir / 'wps_scan.ps1'
    ps_path.write_text(PS_TEMPLATE.replace('__CONTEXT_JSON__', str(context_file)), encoding='utf-8-sig')
    result = subprocess.run(
        ['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', str(ps_path)],
        capture_output=True, text=True, encoding='utf-8', errors='replace', timeout=600,
    )
    combined = (result.stdout or '') + (result.stderr or '')
    if 'WPS-BATCH-OK' not in combined:
        print(f'WPS 扫描失败: {combined[-400:]}')
        return 1

    wps_data = json.loads(scan_out.read_text(encoding='utf-8-sig'))

    # 4) 汇总
    print('\n=== 对拍结果 ===')
    all_match = True
    for entry in files:
        path = entry['path']
        label = entry['label']
        rust_set = rust_results.get(path, set())
        wps_entry = wps_data.get(path, {})
        wps_set = set()
        for sheet in wps_entry.get('cfSheets', []):
            for item in sheet.get('triggered', []):
                addr = item.split('=')[0]
                wps_set.add((sheet['name'], addr))
        match = 'MATCH' if rust_set == wps_set else 'MISMATCH'
        if match == 'MISMATCH':
            all_match = False
        rust_only = sorted(rust_set - wps_set)[:4]
        wps_only = sorted(wps_set - rust_set)[:4]
        print(f'  {label}: rust={len(rust_set)} wps={len(wps_set)} {match}')
        if match == 'MISMATCH':
            if rust_only:
                print(f'    Rust only: {rust_only}')
            if wps_only:
                print(f'    WPS only: {wps_only}')

    if all_match:
        print('\n结论：Rust 求值与 WPS COM 渲染完全一致。')
    else:
        print('\n结论：存在差异，详见上方。')
    return 0


if __name__ == '__main__':
    import re
    sys.exit(main())
