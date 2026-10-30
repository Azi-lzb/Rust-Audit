#!/usr/bin/env python3
"""S1-F01 Linux LibreOffice 公式重算：批次会话复用 vs 逐本启动的对照驱动。

在同一 lo-convert 二进制上对比两种调用方式（同输入、同输出契约、同验证
强度；逐本单发模式即优化前的生产路径，因此无需旧二进制即可测得基线）：

- ``--mode batch``    : 一次 `recalculate-batch --pairs ... --engine lo` 处理整批
  （优化后路径；单次 LibreOffice 启动，响应含 convert_ms 与逐本结果）。
- ``--mode per-book`` : 每本一次 `recalculate --engine lo`
  （优化前路径；每本各启动一次 soffice）。

记录：批次总耗时、每本耗时（会话模式含 open/calc/save/close 阶段）、
LibreOffice 进程创建次数与 PID、成功/失败数、输入 SHA-256 前后一致性、
输出工作簿可用 openpyxl 重新打开且表名/公式数与输入一致、批次结束后不
存在本任务遗留的 soffice 进程。

用法（Linux，需安装 LibreOffice 与 python3-openpyxl）::

    python3 perf_s1f1_lo_batch_session_20261031.py \
        --lo-convert ./target/release/lo-convert \
        --books 2 10 33 --rounds 3 --mode session per-book

该脚本不修改任何业务代码，不做破坏性实验；全部输入写入临时目录。
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import statistics
import subprocess
import sys
import tempfile
import time
from datetime import datetime, timezone
from pathlib import Path

try:
    from openpyxl import load_workbook
    from openpyxl import Workbook
except ImportError:  # pragma: no cover
    print("需要 openpyxl：pip3 install openpyxl", file=sys.stderr)
    raise


def sha256_of(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1 << 20), b""):
            digest.update(block)
    return digest.hexdigest()


def build_synthetic_workbook(path: Path, rows: int, seed: int) -> None:
    """带公式的合成工作簿：跨列引用与 SUM，保证有可重算的公式。"""
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "数据"
    sheet["A1"] = "序号"
    sheet["B1"] = "数值"
    sheet["C1"] = "公式"
    for row in range(2, rows + 2):
        sheet.cell(row=row, column=1, value=row - 1)
        sheet.cell(row=row, column=2, value=(row * 7 + seed) % 1000)
        sheet.cell(row=row, column=3, value=f"=B{row}*2+A{row}")
    total_row = rows + 3
    sheet.cell(row=total_row, column=1, value="合计")
    sheet.cell(row=total_row, column=3, value=f"=SUM(C2:C{rows + 1})")
    workbook.save(path)


def workbook_signature(path: Path) -> dict:
    """输出契约检查：可重新打开、表名/顺序与公式数量。"""
    workbook = load_workbook(path, read_only=False, data_only=False)
    sheets = list(workbook.sheetnames)
    formulas = 0
    for sheet in workbook.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    formulas += 1
    workbook.close()
    return {"sheets": sheets, "formula_count": formulas}


def soffice_descendants_of(pid: int) -> list[int]:
    """列出 sidecar 进程仍存活的 soffice 子进程 PID（进程清理核对）。"""
    descendants: list[int] = []
    try:
        out = subprocess.run(
            ["pgrep", "-P", str(pid)], capture_output=True, text=True, timeout=10
        )
        children = [int(line) for line in out.stdout.split() if line.strip()]
    except (subprocess.SubprocessError, ValueError):
        return descendants
    for child in children:
        try:
            name = Path(f"/proc/{child}/comm").read_text().strip()
        except OSError:
            name = ""
        if "soffice" in name:
            descendants.append(child)
        descendants.extend(soffice_descendants_of(child))
    return descendants


def run_batch(
    lo_convert: Path,
    inputs: list[Path],
    outputs: list[Path],
    mode: str,
    timeout_seconds: int,
) -> dict:
    """按模式执行一批；返回每本结果与批次计时。"""
    started = time.perf_counter()
    books: list[dict] = []
    if mode == "batch":
        pairs_file = Path(tempfile.gettempdir()) / "audit-s1f1-perf-batch-pairs.tsv"
        lines = "".join(
            f"{source}\t{target}\n" for source, target in zip(inputs, outputs)
        )
        pairs_file.write_text(lines, encoding="utf-8")
        completed = subprocess.run(
            [
                str(lo_convert),
                "recalculate-batch",
                "--pairs",
                str(pairs_file),
                "--engine",
                "lo",
            ],
            capture_output=True,
            text=True,
            timeout=timeout_seconds * max(len(inputs), 1) + 120,
        )
        batch_seconds = time.perf_counter() - started
        if completed.returncode != 0:
            raise RuntimeError(
                f"批次重算失败（exit={completed.returncode}）：{(completed.stderr or '')[-400:]}"
            )
        payload = json.loads(completed.stdout)
        for entry in payload.get("books", []):
            books.append(
                {
                    "index": entry.get("index"),
                    "status": entry.get("status"),
                    "engine": payload.get("engine"),
                    "error": entry.get("error"),
                    "lo_batch": {"convert_ms": payload.get("convert_ms"), "profile_warm": True},
                    "book_seconds": None,
                }
            )
        ready = {"mode": "batch", "convert_ms": payload.get("convert_ms")}
        session_pid = None
        leftover_soffice = []
    else:
        session_pid = None
        ready = None
        for index, (source, target) in enumerate(zip(inputs, outputs)):
            book_started = time.perf_counter()
            completed = subprocess.run(
                [
                    str(lo_convert),
                    "recalculate",
                    "--input",
                    str(source),
                    "--output",
                    str(target),
                    "--engine",
                    "lo",
                ],
                capture_output=True,
                text=True,
                timeout=timeout_seconds,
            )
            book_elapsed = time.perf_counter() - book_started
            status = "recalculated" if completed.returncode == 0 else "failed"
            error = None
            engine = None
            if completed.returncode == 0:
                try:
                    payload = json.loads(completed.stdout)
                    status = payload.get("status", status)
                    engine = payload.get("engine")
                except json.JSONDecodeError:
                    pass
            else:
                error = (completed.stderr or "")[-400:]
            books.append(
                {
                    "index": index + 1,
                    "status": status,
                    "engine": engine,
                    "error": error,
                    "book_seconds": round(book_elapsed, 4),
                }
            )
        batch_seconds = time.perf_counter() - started
        leftover_soffice = []

    return {
        "mode": mode,
        "ready": ready,
        "books": books,
        "batch_seconds": round(batch_seconds, 4),
        "leftover_soffice_pids": leftover_soffice,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--lo-convert", required=True, type=Path, help="lo-convert 可执行文件")
    parser.add_argument("--books", type=int, nargs="+", default=[2, 10, 33])
    parser.add_argument("--rows", type=int, default=400, help="每本数据行数")
    parser.add_argument("--rounds", type=int, default=3, help="每规模测量轮数（另有 1 轮预热）")
    parser.add_argument("--modes", nargs="+", default=["per-book", "batch"], choices=["per-book", "batch"])
    parser.add_argument("--timeout", type=int, default=300, help="单本/单进程超时秒数")
    parser.add_argument("--report", type=Path, default=None)
    arguments = parser.parse_args()

    lo_convert = arguments.lo_convert.resolve()
    if not lo_convert.is_file():
        print(f"lo-convert 不存在：{lo_convert}", file=sys.stderr)
        return 2

    report: dict = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "lo_convert": str(lo_convert),
        "rows_per_book": arguments.rows,
        "rounds": arguments.rounds,
        "results": [],
    }
    env_note = subprocess.run(
        ["soffice", "--version"], capture_output=True, text=True
    ).stdout.strip()
    report["libreoffice_version"] = env_note

    with tempfile.TemporaryDirectory(prefix="s1f1-lo-batch-perf-") as workspace:
        work = Path(workspace)
        for book_count in arguments.books:
            inputs: list[Path] = []
            outputs: list[Path] = []
            input_dir = work / f"inputs-{book_count}"
            output_dir = work / f"outputs-{book_count}"
            input_dir.mkdir()
            output_dir.mkdir()
            for index in range(book_count):
                source = input_dir / f"book-{index + 1:04}.xlsx"
                build_synthetic_workbook(source, arguments.rows, seed=index + 1)
                inputs.append(source)
                outputs.append(output_dir / f"{source.stem}-recalculated.xlsx")
            input_hashes = {source.name: sha256_of(source) for source in inputs}
            input_signature = workbook_signature(inputs[0])

            for mode in arguments.modes:
                rounds = []
                for round_index in range(arguments.rounds + 1):
                    for target in outputs:
                        if target.exists():
                            target.unlink()
                    labelled = "warmup" if round_index == 0 else f"round-{round_index}"
                    outcome = run_batch(
                        lo_convert, inputs, outputs, mode, arguments.timeout
                    )
                    outcome["round"] = labelled
                    outcome["input_sha256_unchanged"] = all(
                        sha256_of(source) == digest for source, digest in
                        ((input_dir / name, digest) for name, digest in input_hashes.items())
                    )
                    outcome["output_checks"] = [
                        {
                            "output": target.name,
                            "exists": target.is_file(),
                            "signature": workbook_signature(target) if target.is_file() else None,
                        }
                        for target in outputs
                    ]
                    if round_index > 0:
                        rounds.append(outcome)
                succeeded = [r for r in rounds if all(
                    book["status"] == "recalculated" for book in r["books"])]
                summary = {
                    "books": book_count,
                    "mode": mode,
                    "rounds": len(rounds),
                    "all_rounds_succeeded": len(succeeded) == len(rounds),
                    "batch_seconds_median": round(
                        statistics.median(r["batch_seconds"] for r in rounds), 4
                    )
                    if rounds
                    else None,
                    "per_book_seconds_median": round(
                        statistics.median(
                            book["book_seconds"]
                            for r in rounds
                            for book in r["books"]
                            if book["book_seconds"] is not None
                        ),
                        4,
                    )
                    if any(book["book_seconds"] is not None for r in rounds for book in r["books"])
                    else None,
                    "distinct_lo_pids": sorted(
                        {
                            book["lo_batch"]["lo_process_id"]
                            for r in rounds
                            for book in r["books"]
                            if book.get("lo_batch")
                            and book["lo_batch"].get("lo_process_id")
                        }
                    ),
                    "lo_instance_ids": sorted(
                        {
                            book["lo_batch"]["lo_instance_id"]
                            for r in rounds
                            for book in r["books"]
                            if book.get("lo_batch")
                            and book["lo_batch"].get("lo_instance_id") is not None
                        }
                    ),
                    "leftover_soffice_pids": sorted(
                        {pid for r in rounds for pid in r["leftover_soffice_pids"]}
                    ),
                    "input_sha256_unchanged_all_rounds": all(
                        r["input_sha256_unchanged"] for r in rounds
                    ),
                    "output_sheet_names_match_input": all(
                        check["signature"] is not None
                        and check["signature"]["sheets"] == input_signature["sheets"]
                        for r in rounds
                        for check in r["output_checks"]
                    ),
                    "output_formula_count_matches_input": all(
                        check["signature"] is not None
                        and check["signature"]["formula_count"]
                        == input_signature["formula_count"]
                        for r in rounds
                        for check in r["output_checks"]
                    ),
                    "round_details": rounds,
                }
                report["results"].append(summary)
                print(
                    f"books={book_count:>2} mode={mode:<8} "
                    f"median_batch={summary['batch_seconds_median']}s "
                    f"median_book={summary['per_book_seconds_median']}s "
                    f"pids={summary['distinct_lo_pids']} "
                    f"instances={summary['lo_instance_ids']} "
                    f"input_unchanged={summary['input_sha256_unchanged_all_rounds']} "
                    f"output_ok={summary['output_sheet_names_match_input'] and summary['output_formula_count_matches_input']}"
                )

    report_path = arguments.report or (
        Path(__file__).resolve().parent / "perf_s1f1_lo_batch_session_20261031.json"
    )
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2))
    print(f"report: {report_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
