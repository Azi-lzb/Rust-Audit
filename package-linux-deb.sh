#!/usr/bin/env bash
# Rust-Audit Linux deb 打包（x86_64 Kylin V10 / UOS V20，glibc 2.28 目标）。
#
# 前提：已用 zigbuild 产出 glibc 2.28 目标的 release 二进制：
#   cargo zigbuild --release --target x86_64-unknown-linux-gnu.2.28 \
#     -p audit-core --bin AuditTools
#
# 用法（在 Linux/WSL 中执行）：
#   ./package-linux-deb.sh <binary> <product-dir> <version> [out-dir] [full|s1s2] [docs-dir]
#
# 产物：AuditTools_<version>_amd64.deb（无需用户安装 Python/Office/JVM）
set -euo pipefail

BINARY="${1:?用法: package-linux-deb.sh <binary> <product-dir> <version> [out-dir]}"
PRODUCT_DIR="${2:?缺少产品资源目录（前端/配置模板等）}"
VERSION="${3:?缺少版本号（如 0.1.0）}"
OUT_DIR="${4:-./dist}"
EDITION="${5:-full}"
DOCS_DIR="${6:-}"

[[ -f "$BINARY" ]] || { echo "二进制不存在：$BINARY" >&2; exit 2; }
[[ -d "$PRODUCT_DIR" ]] || { echo "产品目录不存在：$PRODUCT_DIR" >&2; exit 3; }
command -v dpkg-deb >/dev/null || { echo "需要 dpkg-deb（apt install dpkg）" >&2; exit 4; }

WORK="$(mktemp -d)"
trap 'rm -rf "$WORK"' EXIT

case "$EDITION" in
  full)  PKG_SUFFIX="" ;;
  s1s2)  PKG_SUFFIX="-s1s2" ;;
  *)     echo "未知 edition：$EDITION（可用 full / s1s2）" >&2; exit 6 ;;
esac

PKG_ROOT="$WORK/pkg"
mkdir -p "$PKG_ROOT/DEBIAN" "$PKG_ROOT/opt/audittools" \
         "$PKG_ROOT/usr/bin" "$PKG_ROOT/usr/share/applications"

install -m 0755 "$BINARY" "$PKG_ROOT/opt/audittools/AuditTools"
cp -a "$PRODUCT_DIR"/. "$PKG_ROOT/opt/audittools/"

# 真实配置（S1S2 版只带 1./2.）与使用说明
CONFIG_SRC="${CONFIG_SRC:-$(dirname "$BINARY")/../../../../audit-core/config-real}"
PRODUCT_ROOT="$(dirname "$PRODUCT_DIR")"
if [[ -d "$PRODUCT_ROOT/config-real" ]]; then
  mkdir -p "$PKG_ROOT/opt/audittools/config-real"
  if [[ "$EDITION" == "s1s2" ]]; then
    for name in '1.逐笔统计系统_配置.xlsx' '2.报表采集系统_配置.xlsx'; do
      [[ -f "$PRODUCT_ROOT/config-real/$name" ]] &&         cp "$PRODUCT_ROOT/config-real/$name" "$PKG_ROOT/opt/audittools/config-real/"
    done
    # GUI 默认绑定（settings.json 播种）依赖：大集中外部文件与模板目录
    [[ -f "$PRODUCT_ROOT/config-real/外部文件-大集中相关数据.xlsx" ]] &&       cp "$PRODUCT_ROOT/config-real/外部文件-大集中相关数据.xlsx" "$PKG_ROOT/opt/audittools/config-real/"
    [[ -d "$PRODUCT_ROOT/config-real/模板文件_命名区域" ]] &&       cp -r "$PRODUCT_ROOT/config-real/模板文件_命名区域" "$PKG_ROOT/opt/audittools/config-real/"
    GUIDE="${DOCS_DIR:-$PRODUCT_ROOT/docs}/使用说明-S1S2版.docx"
    [[ -f "$GUIDE" ]] && cp "$GUIDE" "$PKG_ROOT/opt/audittools/使用说明.docx"
  else
    cp "$PRODUCT_ROOT/config-real"/*.xlsx "$PKG_ROOT/opt/audittools/config-real/" 2>/dev/null || true
    [[ -d "$PRODUCT_ROOT/config-real/模板文件_命名区域" ]] &&       cp -r "$PRODUCT_ROOT/config-real/模板文件_命名区域" "$PKG_ROOT/opt/audittools/config-real/"
    GUIDE="${DOCS_DIR:-$PRODUCT_ROOT/docs}/使用说明-全量版.docx"
    [[ -f "$GUIDE" ]] && cp "$GUIDE" "$PKG_ROOT/opt/audittools/使用说明.docx"
  fi
else
  echo "警告：未找到 config-real，deb 将不含配置（PRODUCT_DIR 的上级目录）" >&2
fi

ln -s /opt/audittools/AuditTools "$PKG_ROOT/usr/bin/audittools"
if [[ "$EDITION" == "s1s2" ]]; then
  printf 's1s2' > "$PKG_ROOT/opt/audittools/edition-s1s2"
fi

cat > "$PKG_ROOT/DEBIAN/control" <<CONTROL
Package: audittools${PKG_SUFFIX}
Version: ${VERSION}
Section: utils
Priority: optional
Architecture: amd64
Maintainer: Azi-lzb <Azi-lzb@users.noreply.github.com>
Installed-Size: $(du -sk "$PKG_ROOT" | cut -f1)
Depends: libc6 (>= 2.28), libgcc-s1
Recommends: fonts-noto-cjk
Description: AuditTools 审核工具（Rust 实现）
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
DEB="$OUT_DIR/AuditTools${PKG_SUFFIX}_${VERSION}_amd64.deb"
dpkg-deb --root-owner-group --build "$PKG_ROOT" "$DEB" >/dev/null
echo "已生成：$DEB ($(du -h "$DEB" | cut -f1))"
dpkg-deb --info "$DEB" | head -8
