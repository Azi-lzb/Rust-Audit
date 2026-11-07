#!/usr/bin/env python3
"""四本含 CF 的真实业务副本 × 三种求值方式全量对拍。

1. Rust conditional-evaluate（OOXML 规则静态求值）
2. Excel COM DisplayFormat.Interior.Color（渲染色 vs 基础色差值）
3. Excel COM Evaluate（逐条 CF 公式代入单元格求值，看结果是否 TRUE）
"""
import json, subprocess, sys, zipfile, re
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

EXE = Path(r'C:\Users\AZI\Desktop\VibeCode\Rust-Audit\dist\OpenDataAudit\open-data-audit.exe').resolve()
RUN = Path(r'F:/RustBuild/audit/tmp/s1f1-auto-route-current-release-129-20261106/acceptance/steps/S1-F01/S1-F1_审核结果_20261106_144557_098')
OUT = Path(r'C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports')

SAMPLES = [
    ('bond',     '0012-TCL科技集团财务有限公司_金融基础数据采集_债券业务__2026-08-31_数据核对表-recalculated.xlsx'),
    ('deposit',  '0013-TCL科技集团财务有限公司_金融基础数据采集_单位存款__2026-08-31_数据核对表-recalculated.xlsx'),
    ('entrust',  '0016-TCL科技集团财务有限公司_金融基础数据采集_委托贷款__2026-08-31_数据核对表-recalculated.xlsx'),
    ('interbank','0015-TCL科技集团财务有限公司_金融基础数据采集_同业业务__2026-08-31_数据核对表-recalculated.xlsx'),
]

def rust_eval(wb_path):
    r = subprocess.run([str(EXE), 'conditional-evaluate', '--workbook', str(wb_path)],
                       capture_output=True, text=True, encoding='utf-8', timeout=300)
    if r.returncode != 0:
        return {'error': r.stderr[:200], 'triggered': []}
    d = json.loads(r.stdout)
    return {'status': d.get('status'), 'scanned': d.get('scanned_cells'),
            'triggered': sorted((e['sheet'], e['cell']) for e in d.get('triggered', []))}

def excel_com_scan(wb_path, cf_sheets, label):
    """DisplayFormat 渲染色扫描：返回触发格集合。sheet 名经 JSON context 传入。"""
    out_file = f'C:/Users/AZI/AppData/Local/Temp/cf_scan_{label}.json'
    context = {'wbPath': str(wb_path), 'sheets': cf_sheets, 'outFile': out_file}
    context_file = Path(f'C:/Users/AZI/AppData/Local/Temp/cf_scan_{label}_ctx.json')
    context_file.write_text(json.dumps(context, ensure_ascii=False), encoding='utf-8')

    ps_lines = [
        '$ErrorActionPreference = "Stop"',
        '$ctx = Get-Content -Raw -Encoding UTF8 "' + context_file.as_posix() + '" | ConvertFrom-Json',
        '$app = New-Object -ComObject Excel.Application',
        '$app.Visible = $false',
        '$app.DisplayAlerts = $false',
        '$wb = $app.Workbooks.Open($ctx.wbPath, 0, $true)',
        'try {',
        '    $app.CalculateFullRebuild()',
        '    $result = @{}',
        '    foreach ($sheetName in $ctx.sheets) {',
        '        $sheet = $wb.Worksheets.Item($sheetName)',
        '        $triggered = @()',
        '        $scanned = 0',
        '        foreach ($condition in $sheet.UsedRange.FormatConditions) {',
        '            $appliesTo = $condition.AppliesTo',
        '            foreach ($area in $appliesTo.Areas) {',
        '                foreach ($cell in $area.Cells) {',
        '                    $scanned++',
        '                    $addr = $cell.Address($false,$false)',
        '                    try {',
        '                        $display = $cell.DisplayFormat.Interior.Color',
        '                        $base = $cell.Interior.Color',
        '                        if ($display -ne $base) { $triggered += "$addr" }',
        '                    } catch { }',
        '                }',
        '            }',
        '        }',
        '        $result[$sheetName] = @{ scanned = $scanned; triggered = $triggered }',
        '    }',
        '    $json = $result | ConvertTo-Json -Depth 4',
        '    [System.IO.File]::WriteAllText($ctx.outFile, $json, (New-Object System.Text.UTF8Encoding($true)))',
        '    Write-Output "OK"',
        '} finally {',
        '    $wb.Close($false)',
        '    $app.Quit()',
        '}',
    ]
    ps_file = Path(f'C:/Users/AZI/AppData/Local/Temp/cf_scan_{label}.ps1')
    ps_file.write_text('\n'.join(ps_lines), encoding='ascii')
    r = subprocess.run(['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', str(ps_file)],
                       capture_output=True, text=True, encoding='utf-8', errors='replace', timeout=300)
    if 'OK' not in (r.stdout or ''):
        return {'error': (r.stderr or r.stdout)[-200:]}
    scan = json.loads(Path(out_file).read_text(encoding='utf-8-sig'))
    triggered = set()
    for sheet_name, data in scan.items():
        for addr in data.get('triggered', []):
            triggered.add((sheet_name, addr))
    return {'triggered': sorted(f'{s}|{c}' for s, c in triggered)}

def get_cf_info(wb_path):
    """从 XLSX 包提取 CF 规则的 sheet 名、sqref、formula。"""
    z = zipfile.ZipFile(wb_path)
    wbxml = z.read('xl/workbook.xml').decode('utf-8')
    relsxml = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    rmap = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', relsxml))
    sheet_entries = re.findall(r'<sheet name="([^"]+)"[^>]*r:id="(rId\d+)"', wbxml)
    result = []
    for name, rid in sheet_entries:
        target = rmap.get(rid, '').lstrip('/')
        part = target if target.startswith('xl/') else 'xl/' + target
        try:
            d = z.read(part).decode('utf-8', errors='ignore')
        except KeyError:
            continue
        for m in re.finditer(r'<conditionalFormatting sqref="([^"]+)">(.*?)</conditionalFormatting>', d, re.S):
            sqref = m.group(1)
            formulas = re.findall(r'<formula>([^<]+)</formula>', m.group(2))
            result.append({'sheet': name, 'sqref': sqref, 'formulas': formulas})
    z.close()
    return result

# ================================================================

print('=' * 70)
for label, filename in SAMPLES:
    wb_path = RUN / filename
    if not wb_path.is_file():
        print(f'\n{label}: FILE NOT FOUND'); continue

    print(f'\n=== {label} ({filename[:40]}...) ===')

    # Rust 求值
    rust = rust_eval(wb_path)
    rust_triggered = set(rust.get('triggered', []))
    rust_n = len(rust_triggered)
    print(f'  Rust 求值: scanned={rust.get("scanned_cells")} triggered={rust_n}')

    # 从包里提取 CF 信息（sheet 名 + 公式 + sqref）
    cf_info = get_cf_info(wb_path)
    cf_sheet_names = sorted(set(item['sheet'] for item in cf_info))
    total_formulas = sum(len(item['formulas']) for item in cf_info)
    print(f'  CF 规则: {len(cf_info)} 块 / {total_formulas} 条公式 / {len(cf_sheet_names)} 张表')

    # Excel COM DisplayFormat 渲染
    render = excel_com_scan(wb_path, cf_sheet_names, label)
    if 'error' in render:
        print(f'  Excel COM 渲染: ERROR {render["error"][:80]}')
        render_triggered = set()
    else:
        render_triggered = set(tuple(s.split('|', 1)) for s in render['triggered'])
        render_n = len(render_triggered)
        print(f'  Excel COM 渲染: triggered={render_n}')

    # 对比 Rust vs 渲染
    if 'error' not in render:
        match = '✅ MATCH' if rust_triggered == render_triggered else f'❌ MISMATCH (rust_only={len(rust_triggered-render_triggered)} excel_only={len(render_triggered-rust_triggered)})'
        print(f'  Rust vs 渲染: {match}')
        if rust_triggered - render_triggered:
            for item in sorted(rust_triggered - render_triggered)[:3]:
                print(f'    Rust only: {item[0]}!{item[1]}')
        if render_triggered - rust_triggered:
            for item in sorted(render_triggered - rust_triggered)[:3]:
                print(f'    Render only: {item[0]}!{item[1]}')

print('\n' + '=' * 70)
print('完成')
