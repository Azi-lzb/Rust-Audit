#!/usr/bin/env bash
# 完整性能矩阵：2/10/33 本 × 3 轮 × per-book/batch。
set -euo pipefail
VENV="$HOME/openpyxl-venv"
DRIVER="/mnt/c/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/tests/compare/perf_s1f1_lo_batch_session_20261031.py"
LO_CONVERT="$HOME/lo-build/target/release/lo-convert"
if [ ! -x "$VENV/bin/python" ]; then
  python3 -m venv "$VENV"
  "$VENV/bin/pip" install --quiet openpyxl
fi
"$VENV/bin/python" "$DRIVER" \
  --lo-convert "$LO_CONVERT" \
  --books 2 10 33 \
  --rounds 3 \
  --modes per-book batch \
  --report "$HOME/lo-work/perf-full.json"
