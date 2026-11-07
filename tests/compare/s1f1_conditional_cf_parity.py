#!/usr/bin/env python3
"""S1-F1 条件格式对拍：Rust 求值 vs Excel COM / WPS COM 渲染颜色。

sheet 名与路径经 UTF-8 JSON 文件传给 PowerShell（PS1 源码纯 ASCII），
避免中文在命令行/脚本编码中损坏。
"""
from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

EXE = Path(r'C:\Users\AZI\Desktop\VibeCode\Rust-Audit\dist\OpenDataAudit\open-data-audit.exe')
WORKBOOK = Path(r'F:/RustBuild/audit/tmp/s1f1-auto-route-current-release-129-20261106/acceptance/steps/S1-F01/S1-F1_审核结果_20261106_144557_098/0013-TCL科技集团财务有限公司_金融基础数据采集_单位存款__2026-08-31_数据核对表-recalculated.xlsx')
REPORT_DIR = Path(r'C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports')

SHEETS_WITH_CF = ['单位存款余额', '单位存款发生额', '存量单位存款分产品期限', '单位存款发生额分产品期限']

PS_TEMPLATE = r'''
$ErrorActionPreference = "Stop"
$context = Get-Content -Raw -Encoding UTF8 "__CONTEXT_JSON__" | ConvertFrom-Json
$app = $null
$used = ""
foreach ($name in $context.probes) {
    try { $app = New-Object -ComObject $name; $used = $name; break } catch {}
}
if (-not $app) { Write-Output "NO_ENGINE"; exit 1 }
try {
    $app.Visible = $false
    $app.DisplayAlerts = $false
    $wb = $app.Workbooks.Open($context.workbook, 0, $true)
    try {
        $app.CalculateFullRebuild()
        $result = @{}
        foreach ($sheetName in $context.sheetNames) {
            $sheet = $wb.Worksheets.Item($sheetName)
            $cfFormats = $sheet.Cells.FormatConditions
            $count = $cfFormats.Count
            $triggered = @()
            $scanned = 0
            for ($i = 1; $i -le $count; $i++) {
                $fmt = $cfFormats.Item($i)
                $appliesTo = $null
                try { $appliesTo = $fmt.Parent.Range($fmt.AppliesTo.Address($true,$true)) } catch { continue }
                foreach ($cell in $appliesTo.Cells) {
                    $scanned++
                    $addr = $cell.Address($false,$false)
                    $display = $null
                    try { $display = $cell.DisplayFormat.Interior.Color } catch { }
                    if ($null -eq $display) { continue }
                    $base = 16777215
                    try { $base = $cell.Interior.Color } catch {}
                    if ($display -ne $base) {
                        $triggered += ("$addr=$display")
                    }
                }
            }
            $result[$sheetName] = @{ scanned = $scanned; triggeredCount = $triggered.Count; triggered = $triggered }
        }
        $json = $result | ConvertTo-Json -Depth 4
        [System.IO.File]::WriteAllText($context.scanOut, $json, (New-Object System.Text.UTF8Encoding($true)))
        Write-Output "COM-OK engine=$used"
    } finally {
        $wb.Close($false)
        $app.Quit()
    }
} catch {
    Write-Output ("COM-FAIL: " + $_.Exception.Message)
    exit 1
}
'''


def com_scan(probes: list[str], label: str) -> dict:
    context_file = REPORT_DIR / f'cf_scan_{label}_context.json'
    scan_file = REPORT_DIR / f'cf_scan_{label}.json'
    context_file.write_text(json.dumps({
        'probes': probes,
        'workbook': str(WORKBOOK),
        'sheetNames': SHEETS_WITH_CF,
        'scanOut': str(scan_file),
    }, ensure_ascii=False), encoding='utf-8')
    ps_script = REPORT_DIR / f'cf_scan_{label}.ps1'
    ps_script.write_text(PS_TEMPLATE.replace('__CONTEXT_JSON__', str(context_file)), encoding='ascii')
    result = subprocess.run(
        ['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', str(ps_script)],
        capture_output=True, text=True, encoding='utf-8', errors='replace', timeout=600,
    )
    combined = (result.stdout or '') + (result.stderr or '')
    if 'COM-OK' not in combined:
        return {'available': False, 'detail': combined[-500:]}
    data = json.loads(scan_file.read_text(encoding='utf-8-sig'))
    return {'available': True, 'engine': label, 'sheets': data}


def rust_evaluate() -> dict:
    result = subprocess.run(
        [str(EXE), 'conditional-evaluate', '--workbook', str(WORKBOOK)],
        capture_output=True, text=True, encoding='utf-8', timeout=300,
    )
    if result.returncode != 0:
        raise RuntimeError(f'conditional-evaluate failed: {result.stderr[-400:]}')
    return json.loads(result.stdout)


def main() -> int:
    if sys.stdout.encoding and sys.stdout.encoding.lower() not in ('utf-8', 'utf8'):
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    report: dict = {'workbook': str(WORKBOOK), 'rust': None, 'excel_com': None, 'wps_com': None}

    rust = rust_evaluate()
    report['rust'] = {
        'status': rust.get('status'),
        'rules_total': rust.get('rules_total'),
        'rules_supported': rust.get('rules_supported'),
        'scanned_cells': rust.get('scanned_cells'),
        'triggered_total': len(rust.get('triggered', [])),
        'triggered_sample': rust.get('triggered', [])[:5],
        'diagnostics': rust.get('diagnostics', [])[:5],
    }
    print(f"Rust: status={rust['status']} rules={rust['rules_total']}/{rust['rules_supported']} "
          f"scanned={rust['scanned_cells']} triggered={len(rust.get('triggered', []))}")

    excel = com_scan(['Excel.Application'], 'excel')
    report['excel_com'] = excel
    if excel.get('available'):
        total = sum(v.get('triggeredCount', 0) for v in excel['sheets'].values())
        print(f"Excel COM: sheets={len(excel['sheets'])} triggeredTotal={total}")

    wps = com_scan(['KET.Application', 'ket.Application', 'ET.Application'], 'wps')
    report['wps_com'] = wps
    if wps.get('available'):
        total = sum(v.get('triggeredCount', 0) for v in wps['sheets'].values())
        print(f"WPS COM: sheets={len(wps['sheets'])} triggeredTotal={total}")
    else:
        print('WPS COM unavailable (recorded honestly)')

    out = REPORT_DIR / 's1f1_conditional_cf_parity.json'
    out.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    print(f'report: {out}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
