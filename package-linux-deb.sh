#!/usr/bin/env bash
# Rust-Audit Linux deb 打包（x86_64 Kylin V10 / UOS V20，glibc 2.28 目标）。
#
# 前提：已用 zigbuild 产出 glibc 2.28 目标的 release 二进制：
#   cargo zigbuild --release --target x86_64-unknown-linux-gnu.2.28 \
#     -p audit-core --bin open-data-audit
#
# 用法（在 Linux/WSL 中执行）：
#   ./package-linux-deb.sh <binary> <product-dir> <version> [out-dir]
#
# 产物：open-data-audit_<version>_amd64.deb（无需用户安装 Python/Office/JVM）
set -euo pipefail

BINARY="${1:?用法: package-linux-deb.sh <binary> <product-dir> <version> [out-dir]}"
PRODUCT_DIR="${2:?缺少产品资源目录（前端/配置模板等）}"
VERSION="${3:?缺少版本号（如 0.1.0）}"
OUT_DIR="${4:-./dist}"

[[ -f "$BINARY" ]] || { echo "二进制不存在：$BINARY" >&2; exit 2; }
[[ -d "$PRODUCT_DIR" ]] || { echo "产品目录不存在：$PRODUCT_DIR" >&2; exit 3; }
command -v dpkg-deb >/dev/null || { echo "需要 dpkg-deb（apt install dpkg）" >&2; exit 4; }

WORK="$(mktemp -d)"
trap 'rm -rf "$WORK"' EXIT

PKG_ROOT="$WORK/pkg"
mkdir -p "$PKG_ROOT/DEBIAN" "$PKG_ROOT/opt/open-data-audit" \
         "$PKG_ROOT/usr/bin" "$PKG_ROOT/usr/share/applications"

install -m 0755 "$BINARY" "$PKG_ROOT/opt/open-data-audit/open-data-audit"
cp -a "$PRODUCT_DIR"/. "$PKG_ROOT/opt/open-data-audit/"

ln -s /opt/open-data-audit/open-data-audit "$PKG_ROOT/usr/bin/open-data-audit"

cat > "$PKG_ROOT/DEBIAN/control" <<CONTROL
Package: open-data-audit
Version: ${VERSION}
Section: utils
Priority: optional
Architecture: amd64
Maintainer: Azi-lzb <Azi-lzb@users.noreply.github.com>
Installed-Size: $(du -sk "$PKG_ROOT" | cut -f1)
Depends: libc6 (>= 2.28), libgcc-s1
Recommends: fonts-noto-cjk
Description: Open-Data-Audit 数据审核工具（Rust 实现）
 S1 逐笔统计核查、S2 报表采集核对、S3 大集中统计比较与金融表单生成。
 纯 Rust 计算引擎（betteroffice）内置，无需安装 Excel/WPS/LibreOffice/JVM。
CONTROL

# glibc 2.28 上限校验：防止误打包新版工具链产物
MAX_GLIBC="$(objdump -T "$BINARY" 2>/dev/null \
  | grep -oE 'GLIBC_2\.[0-9]+' | sort -V | tail -1 | sed 's/GLIBC_//')"
if [[ -n "$MAX_GLIBC" ]] && [[ "$(printf '%s\n2.28\n' "$MAX_GLIBC" | sort -V | tail -1)" != "2.28" ]]; then
  echo "拒绝打包：二进制需要 glibc $MAX_GLIBC，超出 2.28 交付基线" >&2
  exit 5
fi
echo "glibc 要求：${MAX_GLIBC:-unknown}（基线 2.28）"

mkdir -p "$OUT_DIR"
DEB="$OUT_DIR/open-data-audit_${VERSION}_amd64.deb"
dpkg-deb --root-owner-group --build "$PKG_ROOT" "$DEB" >/dev/null
echo "已生成：$DEB ($(du -h "$DEB" | cut -f1))"
dpkg-deb --info "$DEB" | head -8
