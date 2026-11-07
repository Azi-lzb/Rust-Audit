#!/usr/bin/env python3
"""S1-F1 全量条件格式对拍：Rust 求值 vs Excel COM vs WPS COM（真实业务 129 本）。

对 run6（s1f1-auto-route-current-release-129-20261106）的全部 recalculated
审核副本执行三路对拍：
  1. Rust: open-data-audit conditional-evaluate（规则条件求值，触发集合）。
  2. Excel COM: 单实例批次，逐本打开 → CalculateFullRebuild → 按
     FormatConditions.AppliesTo.Areas 逐格读 DisplayFormat.Interior.Color vs
     Interior.Color（渲染色 != 基础色 = 触发）。
  3. WPS COM: 同上，但按旧 Python Audit 的刷新语义（ScreenUpdating=True →
     workbook.Activate → sheet.Activate → CalculateFullRebuild）后再读；
     扫描的是临时副本，输入文件不动，Close(false) 不保存。

无 CF 规则的文件以 ZIP 级证据记录（包内无 conditionalFormatting 部件 =
无可渲染对象），不浪费 COM 打开。
输出：JSON + Markdown 报告（逐文件三引擎触发集合与差异）。
"""
from __future__ import annotations

import json
import re
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

EXE = Path(r'C:\Users\AZI\Desktop\VibeCode\Rust-Audit\dist\OpenDataAudit\open-data-audit.exe')
RUN_DIR = Path(r'F:/RustBuild/audit/tmp/s1f1-auto-route-current-release-129-20261106/acceptance/steps/S1-F01')
REPORT_DIR = Path(r'C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports')

PS_EXCEL_TEMPLATE = r'''
$ErrorActionPreference = "Continue"
$context = Get-Content -Raw -Encoding UTF8 "__CONTEXT_JSON__" | ConvertFrom-Json
$app = New-Object -ComObject Excel.Application
$app.Visible = $false
$app.DisplayAlerts = $false
$out = @{}
try {
    foreach ($entry in $context.files) {
        $fileResult = @{ cfSheets = @(); noCf = $false; error = "" }
        try {
            $wb = $app.Workbooks.Open($entry.path, 0, $true)
            try {
                $app.CalculateFullRebuild()
                if ($entry.sheets.Count -eq 0) {
                    $fileResult.noCf = $true
                } else {
                    $cfSheets = @()
                    foreach ($sheetEntry in $entry.sheets) {
                        $sheet = $wb.Worksheets.Item($sheetEntry.name)
                        $app.CalculateFullRebuild()
                        $triggered = @()
                        $scanned = 0
                        foreach ($condition in $sheet.UsedRange.FormatConditions) {
                            $appliesTo = $condition.AppliesTo
                            foreach ($area in $appliesTo.Areas) {
                                foreach ($cell in $area.Cells) {
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
                        }
                        $cfSheets += @{ name = $sheetEntry.name; scanned = $scanned; triggered = $triggered }
                    }
                    $fileResult.cfSheets = $cfSheets
                }
            } finally {
                $wb.Close($false)
            }
        } catch {
            $fileResult.error = $_.Exception.Message
        }
        $out[$entry.path] = $fileResult
    }
    $json = $out | ConvertTo-Json -Depth 6
    [System.IO.File]::WriteAllText($context.scanOut, $json, (New-Object System.Text.UTF8Encoding($true)))
    Write-Output "BATCH-OK"
} finally {
    $app.Quit()
}
'''

PS_WPS_TEMPLATE = r'''
$ErrorActionPreference = "Continue"
$context = Get-Content -Raw -Encoding UTF8 "__CONTEXT_JSON__" | ConvertFrom-Json
$app = New-Object -ComObject KET.Application
$app.Visible = $false
$app.DisplayAlerts = $false
$out = @{}
try {
    foreach ($entry in $context.files) {
        $fileResult = @{ cfSheets = @(); noCf = $false; error = "" }
        try {
            $wb = $app.Workbooks.Open($entry.path)
            try {
                if ($entry.sheets.Count -eq 0) {
                    $fileResult.noCf = $true
                } else {
                    $cfSheets = @()
                    $wb.Activate()
                    $app.ScreenUpdating = $true
                    foreach ($sheetEntry in $entry.sheets) {
                        $sheet = $wb.Worksheets.Item($sheetEntry.name)
                        $sheet.Activate()
                        try { $app.CalculateFullRebuild() } catch { try { $app.Calculate() } catch {} }
                        $triggered = @()
                        $scanned = 0
                        foreach ($condition in $sheet.UsedRange.FormatConditions) {
                            $appliesTo = $condition.AppliesTo
                            foreach ($area in $appliesTo.Areas) {
                                foreach ($cell in $area.Cells) {
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
                        }
                        $cfSheets += @{ name = $sheetEntry.name; scanned = $scanned; triggered = $triggered }
                    }
                    $fileResult.cfSheets = $cfSheets
                    $app.ScreenUpdating = $false
                }
            } finally {
                $wb.Close($false)
            }
        } catch {
            $fileResult.error = $_.Exception.Message
        }
        $out[$entry.path] = $fileResult
    }
    $json = $out | ConvertTo-Json -Depth 6
    [System.IO.File]::WriteAllText($context.scanOut, $json, (New-Object System.Text.UTF8Encoding($true)))
    Write-Output "BATCH-OK"
} finally {
    $app.Quit()
}
'''


def enumerate_files() -> list[Path]:
    return sorted(RUN_DIR.glob('*/[0-9][0-9][0-9][0-9]-*-recalculated.xlsx'))


def cf_inventory(path: Path) -> tuple[int, list[dict]]:
    """ZIP 级统计 CF 块数并列出带 CF 的工作表名。"""
    with zipfile.ZipFile(path) as archive:
        names = set(archive.namelist())
        workbook_xml = archive.read('xl/workbook.xml').decode('utf-8', errors='ignore')
        rels_xml = archive.read('xl/_rels/workbook.xml.rels').decode('utf-8', errors='ignore')
        rel_map = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels_xml))
        sheet_entries = re.findall(r'<sheet name="([^"]+)"[^>]*r:id="(rId\d+)"', workbook_xml)
        cf_sheets: list[dict] = []
        total = 0
        for name, rel_id in sheet_entries:
            target = rel_map.get(rel_id, '')
            target = target.lstrip('/')
            part = target if target in names else 'xl/' + target
            if part not in names:
                continue
            sheet_xml = archive.read(part).decode('utf-8', errors='ignore')
            count = sheet_xml.count('<conditionalFormatting')
            if count:
                total += count
                cf_sheets.append({'name': name, 'cfBlocks': count})
        return total, cf_sheets


def run_rust(files: list[Path]) -> dict:
    results = {}
    for index, path in enumerate(files):
        proc = subprocess.run(
            [str(EXE), 'conditional-evaluate', '--workbook', str(path)],
            capture_output=True, text=True, encoding='utf-8', timeout=300,
        )
        if proc.returncode != 0:
            results[str(path)] = {'error': proc.stderr[-300:]}
            continue
        data = json.loads(proc.stdout)
        triggered = {(e['sheet'], e['cell']) for e in data.get('triggered', [])}
        results[str(path)] = {
            'status': data.get('status'),
            'scanned': data.get('scanned_cells'),
            'triggered': sorted(f'{s}|{c}' for s, c in triggered),
        }
        if (index + 1) % 25 == 0:
            print(f'  rust {index + 1}/{len(files)}')
    return results


def run_com_batch(files: list[Path], engine: str, work_dir: Path) -> dict:
    """files: 已有 CF 的文件清单。engine: excel | wps。"""
    template = PS_EXCEL_TEMPLATE if engine == 'excel' else PS_WPS_TEMPLATE
    manifest = []
    for path in files:
        scan_path = path
        if engine == 'wps':
            copy_dir = work_dir / 'wps-copies'
            copy_dir.mkdir(parents=True, exist_ok=True)
            copy_path = copy_dir / path.name
            if not copy_path.exists():
                shutil.copy2(path, copy_path)
            scan_path = copy_path
        _, cf_sheets = cf_inventory(path)
        manifest.append({
            'path': str(scan_path),
            'sheets': [{'name': s['name']} for s in cf_sheets],
        })
    context_file = work_dir / f'cf_batch_{engine}_context.json'
    scan_out = work_dir / f'cf_batch_{engine}.json'
    context_file.write_text(json.dumps({'files': manifest, 'scanOut': str(scan_out)}, ensure_ascii=False), encoding='utf-8')
    template_text = template.replace('__CONTEXT_JSON__', str(context_file))
    if engine == 'wps':
        template_text = template_text.replace('$applies_to.Areas', '$appliesTo.Areas')
    ps_path = work_dir / f'cf_batch_{engine}.ps1'
    ps_path.write_text(template_text, encoding='ascii')
    result = subprocess.run(
        ['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', str(ps_path)],
        capture_output=True, text=True, encoding='utf-8', errors='replace', timeout=1800,
    )
    combined = (result.stdout or '') + (result.stderr or '')
    if 'BATCH-OK' not in combined:
        return {'available': False, 'detail': combined[-600:]}
    data = json.loads(scan_out.read_text(encoding='utf-8-sig'))
    return {'available': True, 'files': data}


def triggered_set(entry: dict) -> set:
    result = set()
    for sheet in entry.get('cfSheets', []):
        for item in sheet.get('triggered', []):
            addr = item.split('=')[0]
            result.add((sheet['name'], addr))
    return result


def main() -> int:
    if sys.stdout.encoding and sys.stdout.encoding.lower() not in ('utf-8', 'utf8'):
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    work_dir = REPORT_DIR / 's1f1_cf_parity_work'
    work_dir.mkdir(parents=True, exist_ok=True)

    files = enumerate_files()
    print(f'真实业务副本总数: {len(files)}')
    inventory = {}
    cf_files: list[Path] = []
    for path in files:
        total, cf_sheets = cf_inventory(path)
        inventory[str(path)] = {'cfBlocks': total, 'cfSheets': cf_sheets}
        if total:
            cf_files.append(path)
    print(f'含条件格式的副本: {len(cf_files)}；无条件格式: {len(files) - len(cf_files)}（ZIP 级证据）')

    print('--- Rust 求值（全部文件） ---')
    rust = run_rust(files)

    print('--- Excel COM 批次（CF 文件） ---')
    excel = run_com_batch(cf_files, 'excel', work_dir)

    print('--- WPS COM 批次（CF 文件，含刷新） ---')
    wps = run_com_batch(cf_files, 'wps', work_dir)

    details = []
    totals = {'rust': 0, 'excel': 0, 'wps': 0, 'match': 0, 'mismatch': 0}
    for path in cf_files:
        key = str(path)
        name = path.name[:60]
        rust_entry = rust.get(key, {})
        # Rust 触发是 'sheet|cell' 字符串；Excel/WPS 是按 cfSheets 分组的
        # 'cell=颜色'。统一成 (sheet, cell) 元组后对比。
        rust_set = {
            tuple(entry.split('|', 1))
            for entry in rust_entry.get('triggered', [])
        }
        excel_entry = excel.get('files', {}).get(str(path), {}) if excel.get('available') else {}
        excel_set = triggered_set(excel_entry) if excel_entry else set()
        copy_key = str(work_dir / 'wps-copies' / path.name)
        wps_entry = wps.get('files', {}).get(copy_key, {}) if wps.get('available') else {}
        wps_set = triggered_set(wps_entry) if wps_entry else set()
        match_re = 'n/a'
        if excel.get('available'):
            match_re = 'MATCH' if rust_set == excel_set else 'MISMATCH'
            totals['match' if match_re == 'MATCH' else 'mismatch'] += 1
        detail = {
            'file': name,
            'rustTriggered': len(rust_set),
            'excelTriggered': len(excel_set),
            'wpsTriggered': len(wps_set),
            'rustVsExcel': match_re,
            'wpsAvailable': wps.get('available', False),
            'rustOnly': sorted(rust_set - excel_set)[:8] if excel.get('available') else [],
            'excelOnly': sorted(excel_set - rust_set)[:8] if excel.get('available') else [],
        }
        details.append(detail)
        totals['rust'] += len(rust_set)
        totals['excel'] += len(excel_set)
        totals['wps'] += len(wps_set)
        print(f"  {name[:50]}: rust={len(rust_set)} excel={len(excel_set)} {match_re}")

    report = {
        'runDir': str(RUN_DIR),
        'totalFiles': len(files),
        'filesWithCf': len(cf_files),
        'filesWithoutCf': len(files) - len(cf_files),
        'totals': totals,
        'details': details,
        'inventory': inventory,
    }
    json_out = REPORT_DIR / 's1f1_cf_parity_batch.json'
    json_out.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')

    lines = [
        '# S1-F1 条件格式对拍报告（Rust vs Excel COM vs WPS COM，真实业务副本全量）', '',
        f'- 副本总数：{len(files)}；含条件格式：{len(cf_files)}；无条件格式（ZIP 级证据）：{len(files) - len(cf_files)}',
        f'- 触发格总数：Rust={totals["rust"]} Excel={totals["excel"]} WPS={totals["wps"]}',
        f'- 含 CF 副本的 Rust vs Excel 一致性：MATCH {totals["match"]} / MISMATCH {totals["mismatch"]}', '',
        '| 文件 | Rust | Excel | WPS | Rust vs Excel |', '|---|---:|---:|---:|---|',
    ]
    for detail in details:
        lines.append(f"| {detail['file']} | {detail['rustTriggered']} | {detail['excelTriggered']} | {detail['wpsTriggered']} | {detail['rustVsExcel']} |")
    lines += ['', '## 差异明细（Rust 有 Excel 无 / Excel 有 Rust 无）', '']
    for detail in details:
        if detail['rustVsExcel'] == 'MISMATCH':
            lines.append(f"### {detail['file']}")
            lines.append(f"- Rust only: {detail['rustOnly']}")
            lines.append(f"- Excel only: {detail['excelOnly']}")
    md_out = REPORT_DIR / 's1f1_cf_parity_batch.md'
    md_out.write_text('\n'.join(lines), encoding='utf-8')
    print(f'report: {json_out}')
    print(f'report: {md_out}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
