#!/usr/bin/env bash
# doc→docx 批次转换实测（LO 路径）：造 3 份 .doc 夹具 → convert-batch → 校验。
set -euo pipefail
VENV="$HOME/openpyxl-venv"
LO="$HOME/lo-build/target/release/lo-convert"
BASE="$HOME/lo-docx-test"
rm -rf "$BASE"; mkdir -p "$BASE/src" "$BASE/out"

# 造 3 份简单文本文档，用 LO 转成 .doc 作为夹具
for n in 1 2 3; do
  printf '机构测试文档 %s\n第一行数据\n第二行数据\n' "$n" > "$BASE/src/doc$n.txt"
  soffice --headless --convert-to doc --outdir "$BASE/src" "$BASE/src/doc$n.txt" >/dev/null 2>&1
done
ls "$BASE/src"/*.doc

# 写批次清单（.doc → .docx）
: > "$BASE/pairs.tsv"
for n in 1 2 3; do
  printf '%s\t%s\n' "$BASE/src/doc$n.doc" "$BASE/out/doc$n.docx" >> "$BASE/pairs.tsv"
done

"$LO" convert-batch --pairs "$BASE/pairs.tsv" --to docx --engine lo
echo "--- outputs ---"
ls -la "$BASE/out/"
