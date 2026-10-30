#!/usr/bin/env bash
# WSL 内运行 S1-F01 LibreOffice 批次会话验证的一次性脚本。
# 用法: wsl -d Ubuntu -- bash /mnt/c/.../wsl_run_driver.sh <mode> <books> <rounds>
set -euo pipefail

VENV="$HOME/openpyxl-venv"
DRIVER="/mnt/c/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/tests/compare/perf_s1f1_lo_batch_session_20261031.py"
LO_CONVERT="$HOME/lo-build/target/release/lo-convert"
REPORT="$HOME/lo-work/report-$1-$2.json"

if [ ! -x "$VENV/bin/python" ]; then
  python3 -m venv "$VENV"
  "$VENV/bin/pip" install --quiet openpyxl
fi

mkdir -p "$HOME/lo-work"
"$VENV/bin/python" "$DRIVER" \
  --lo-convert "$LO_CONVERT" \
  --books "$2" \
  --rounds "$3" \
  --modes "$1" \
  --report "$REPORT"
