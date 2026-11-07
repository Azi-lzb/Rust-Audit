#!/usr/bin/env python3
"""把本机 Python 精简复制为随包隔离运行时（openpyxl worker 用）。

产出目录布局（供 OpenpyxlWorkerConfig::for_product_binary 解析）：
    <dest>/python.exe          ← 复制的前缀解释器
    <dest>/python3xx.dll 等
    <dest>/Lib/...             ← 标准库（含 site-packages）
    <dest>/DLLs/...
    <dest>/site-packages/openpyxl、et_xmlfile（确保可导入）

用法（用系统 Python 运行，系统 Python 必须已装 openpyxl）：
    python bundle_openpyxl_runtime.py --dest <目录>

排除项：tcl/tk（GUI）、include、Tools、__pycache__、test —— 减小体积；
openpyxl 及其依赖不依赖这些目录。复制完成后用 <dest>/python.exe -I 验证
`import openpyxl` 成功，失败即非零退出。
"""
from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from pathlib import Path

IGNORE_TOP_LEVEL = {
    "tcl",
    "tk",
    "Tools",
    "include",
    "idlelib",
    "test",
    "__pycache__",
    "site-packages",  # 单独按需复制
}
IGNORE_LIB = {"test", "idlelib", "tkinter", "turtledemo", "__pycache__"}
SITE_PACKAGES_NEEDED = ("openpyxl", "et_xmlfile")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--dest", required=True, type=Path)
    args = parser.parse_args()

    dest = args.dest.resolve()
    prefix = Path(sys.prefix).resolve()
    if dest == prefix or dest in prefix.parents or prefix in dest.parents:
        print("错误：目标目录不能与源 Python 前缀重叠", file=sys.stderr)
        return 2

    print(f"源前缀: {prefix}")
    print(f"目标:   {dest}")
    if dest.exists():
        shutil.rmtree(dest)
    dest.mkdir(parents=True)

    # 1) 前缀顶层：解释器、DLL、Lib、LICENSE
    for entry in prefix.iterdir():
        if entry.name in IGNORE_TOP_LEVEL:
            continue
        target = dest / entry.name
        if entry.is_dir():
            shutil.copytree(entry, target, dirs_exist_ok=True)
        else:
            shutil.copy2(entry, target)

    # 2) Lib 内排除 tkinter/test
    lib_dest = dest / "Lib"
    for name in IGNORE_LIB:
        shutil.rmtree(lib_dest / name, ignore_errors=True)

    # 3) site-packages：只带 openpyxl 及其依赖
    for source_site in (prefix / "Lib" / "site-packages",):
        if not source_site.is_dir():
            continue
        target_site = lib_dest / "site-packages"
        target_site.mkdir(parents=True, exist_ok=True)
        for name in SITE_PACKAGES_NEEDED:
            source = source_site / name
            if source.is_dir():
                shutil.copytree(source, target_site / name, dirs_exist_ok=True)
            elif source.is_file():
                shutil.copy2(source, target_site / name)

    # 4) 清理所有 __pycache__ 与 .pyc
    for cache in list(dest.rglob("__pycache__")):
        shutil.rmtree(cache, ignore_errors=True)
    for pyc in dest.rglob("*.pyc"):
        pyc.unlink(missing_ok=True)

    # 5) 验证：-I 模式导入 openpyxl
    interpreter = dest / "python.exe"
    if not interpreter.is_file():
        print(f"错误：复制后缺少 {interpreter}", file=sys.stderr)
        return 3
    probe = subprocess.run(
        [str(interpreter), "-I", "-c", "import openpyxl; print(openpyxl.__version__)"],
        capture_output=True,
        text=True,
        timeout=120,
    )
    if probe.returncode != 0:
        print(f"错误：隔离运行时无法导入 openpyxl：{probe.stderr[-400:]}", file=sys.stderr)
        return 4
    version = probe.stdout.strip()
    total_files = sum(1 for _ in dest.rglob("*") if _.is_file())
    print(f"隔离运行时就绪：openpyxl {version}；共 {total_files} 个文件于 {dest}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
