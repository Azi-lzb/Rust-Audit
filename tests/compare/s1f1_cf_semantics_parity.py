#!/usr/bin/env python3
"""S1-F1 条件格式语义对拍：Rust conditional-evaluate vs Python checkpoint COM。

读取 Python 侧 checkpoint.jsonl 的 excel.com_evaluate/com_display positions，
在对应的 recalculated 审核副本上跑 Rust conditional-evaluate，逐文件对比
触发位置集合。差异必须为 0，否则逐条列出。
"""
from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

EXE = Path(r'C:\Users\AZI\Desktop\VibeCode\Rust-Audit\dist\OpenDataAudit\open-data-audit.exe')
CHECKPOINT = Path(r'C:/Users/AZI/Desktop/VibeCode/3.open-data-audit/test-output/s1f1_mode_parity_20261107/checkpoint.jsonl')
SEARCH_ROOTS = [
    Path(r'F:/RustBuild/audit/tmp/s1f1-auto-route-current-release-129-20261106/acceptance/steps/S1-F01'),
]
REPORT_DIR = Path(r'C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports')


def find_recalculated(name_stem: str) -> Path | None:
    for root in SEARCH_ROOTS:
        for path in root.rglob('*-recalculated.xlsx'):
            if name_stem in path.name:
                return path
    return None


def rust_evaluate(workbook: Path) -> dict:
    result = subprocess.run(
        [str(EXE), 'conditional-evaluate', '--workbook', str(workbook)],
        capture_output=True, text=True, encoding='utf-8', timeout=300,
    )
    if result.returncode != 0:
        return {'error': result.stderr[-300:], 'triggered': [], 'scanned_cells': 0}
    return json.loads(result.stdout)


def normalize_positions(positions: list) -> set:
    """COM positions 格式为 [sheet, cell] 或 'Sheet!Cell' 字符串。统一为 (sheet, cell) 集。"""
    result = set()
    for entry in positions:
        if isinstance(entry, list) and len(entry) >= 2:
            result.add((str(entry[0]), str(entry[1])))
        elif isinstance(entry, str):
            parts = entry.rsplit('!', 1)
            if len(parts) == 2:
                result.add((parts[0].strip("'"), parts[1]))
    return result


def main() -> int:
    if sys.stdout.encoding and sys.stdout.encoding.lower() not in ('utf-8', 'utf8'):
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    REPORT_DIR.mkdir(parents=True, exist_ok=True)

    lines = CHECKPOINT.read_text(encoding='utf-8').splitlines()
    print(f'checkpoint 行数: {len(lines)}')

    match_count = 0
    mismatch_count = 0
    skipped_no_file = 0
    skipped_no_com = 0
    details = []

    for line in lines:
        checkpoint = json.loads(line)
        file_stem = checkpoint['file'].replace('.xlsx', '').replace('.xlsm', '')
        excel_data = checkpoint.get('excel') or {}
        if not isinstance(excel_data, dict):
            excel_data = {}

        # 取 COM evaluate positions（优先）或 display positions
        com_evaluate = (excel_data.get('com_evaluate') or {})
        com_display = (excel_data.get('com_display') or {})
        com_positions = com_evaluate.get('positions') or com_display.get('positions') or []

        # 找对应的 recalculated 副本
        workbook = find_recalculated(file_stem)
        if workbook is None:
            skipped_no_file += 1
            continue
        if not com_positions:
            skipped_no_com += 1
            continue

        # Rust 求值
        rust_output = rust_evaluate(workbook)
        rust_triggered = {
            (e['sheet'], e['cell'])
            for e in rust_output.get('triggered', [])
        }
        com_set = normalize_positions(com_positions)

        if rust_triggered == com_set:
            match_count += 1
        else:
            mismatch_count += 1
            rust_only = sorted(rust_triggered - com_set)[:6]
            com_only = sorted(com_set - rust_triggered)[:6]
            details.append({
                'file': file_stem[:60],
                'rustCount': len(rust_triggered),
                'comCount': len(com_set),
                'rustOnly': rust_only,
                'comOnly': com_only,
            })
            print(f'  MISMATCH: {file_stem[:50]} rust={len(rust_triggered)} com={len(com_set)}')

    print(f'\n=== 对拍汇总 ===')
    print(f'匹配: {match_count} / 不匹配: {mismatch_count}')
    print(f'跳过（找不到副本）: {skipped_no_file}')
    print(f'跳过（COM positions 为空）: {skipped_no_com}')

    report = {
        'total_checkpoint': len(lines),
        'match': match_count,
        'mismatch': mismatch_count,
        'skipped_no_file': skipped_no_file,
        'skipped_no_com_positions': skipped_no_com,
        'mismatch_details': details,
    }
    out = REPORT_DIR / 's1f1_cf_semantics_parity.json'
    out.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    print(f'report: {out}')
    return 0 if mismatch_count == 0 else 1


if __name__ == '__main__':
    sys.exit(main())
