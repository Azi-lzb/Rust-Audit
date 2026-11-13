#!/usr/bin/env bash
# Rust-Audit Linux deb 打包（x86_64 Kylin V10 / UOS V20，glibc 2.28 目标）。
#
# 前提：已用 zigbuild 产出 glibc 2.28 目标的 release 二进制：
#   cargo zigbuild --release --target x86_64-unknown-linux-gnu.2.28 \
#     -p audit-core --bin AuditTools
#
# 用法（在 Linux/WSL 中执行）；加 --validate-only 仅校验资源，不创建 deb：
#   ./package-linux-deb.sh <binary> <product-dir> <version> [out-dir] [full|s1s2] [docs-dir]
#
# 产物：AuditTools_<version>_amd64.deb（无需用户安装 Python/Office/JVM）
set -euo pipefail

VALIDATE_ONLY=0
if [[ "${1:-}" == "--validate-only" ]]; then
  VALIDATE_ONLY=1
  shift
fi
SCRIPT_DIR="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"

BINARY="${1:?用法: package-linux-deb.sh <binary> <product-dir> <version> [out-dir]}"
PRODUCT_DIR="${2:?缺少产品资源目录（前端/配置模板等）}"
VERSION="${3:?缺少版本号（如 0.1.0）}"
OUT_DIR="${4:-./dist}"
EDITION="${5:-full}"
DOCS_DIR="${6:-}"

[[ -f "$BINARY" ]] || { echo "二进制不存在：$BINARY" >&2; exit 2; }
[[ -d "$PRODUCT_DIR" ]] || { echo "产品目录不存在：$PRODUCT_DIR" >&2; exit 3; }
CONFIG_PRESERVATION="$SCRIPT_DIR/audit-core/packaging/config-preservation.sh"
[[ -f "$CONFIG_PRESERVATION" ]] || { echo "Missing configuration preservation helper" >&2; exit 7; }
case "$EDITION" in
  full)
    PKG_SUFFIX=""
    OTHER_EDITION="audittools-s1s2"
    CONFIG_FILES=('1.逐笔统计系统_配置.xlsx' '2.报表采集系统_配置.xlsx'
      '3.0大集中通用配置.xlsx' '3.1大集中执行比较_配置.xlsx'
      '3.2大集中本期数值核对_配置.xlsx' '3.3大集中指标比较拆分_配置.xlsx')
    GUIDE_NAME='使用说明-全量版.docx'
    ;;
  s1s2)
    PKG_SUFFIX="-s1s2"
    OTHER_EDITION="audittools"
    CONFIG_FILES=('1.逐笔统计系统_配置.xlsx' '2.报表采集系统_配置.xlsx')
    GUIDE_NAME='使用说明-S1S2版.docx'
    ;;
  *) echo "未知 edition：$EDITION（可用 full / s1s2）" >&2; exit 6 ;;
esac

PRODUCT_ROOT="$(dirname "$PRODUCT_DIR")"
# Explicit CONFIG_SRC is authoritative: never silently substitute a different configuration.
if [[ -z "${CONFIG_SRC:-}" ]]; then
  if [[ -d "$PRODUCT_DIR/config-real" ]]; then CONFIG_SRC="$PRODUCT_DIR/config-real"
  elif [[ -d "$PRODUCT_ROOT/config-real" ]]; then CONFIG_SRC="$PRODUCT_ROOT/config-real"
  else CONFIG_SRC="$SCRIPT_DIR/audit-core/config-real"
  fi
fi
[[ -d "$CONFIG_SRC" ]] || { echo "必要配置目录不存在：$CONFIG_SRC" >&2; exit 7; }
for name in "${CONFIG_FILES[@]}" '外部文件-大集中相关数据.xlsx'; do
  [[ -f "$CONFIG_SRC/$name" ]] || { echo "缺少必要配置或外部文件：$CONFIG_SRC/$name" >&2; exit 7; }
done
[[ -d "$CONFIG_SRC/模板文件_命名区域" ]] || { echo "缺少必要模板目录：$CONFIG_SRC/模板文件_命名区域" >&2; exit 7; }
if [[ -z "$DOCS_DIR" ]]; then
  if [[ -d "$PRODUCT_ROOT/docs" ]]; then DOCS_DIR="$PRODUCT_ROOT/docs"; else DOCS_DIR="$SCRIPT_DIR/docs"; fi
fi
GUIDE="$DOCS_DIR/$GUIDE_NAME"
[[ -f "$GUIDE" ]] || { echo "缺少使用说明：$GUIDE" >&2; exit 7; }
[[ -f "$PRODUCT_DIR/assets/app-icon.png" ]] || { echo "缺少程序图标：$PRODUCT_DIR/assets/app-icon.png" >&2; exit 7; }
if [[ "$VALIDATE_ONLY" == "1" ]]; then
  printf 'Validation passed: edition=%s; configs=%s; source=%s\n' "$EDITION" "${#CONFIG_FILES[@]}" "$CONFIG_SRC"
  printf 'Package: audittools%s\nConflicts: %s\nReplaces: %s\n' "$PKG_SUFFIX" "$OTHER_EDITION" "$OTHER_EDITION"
  exit 0
fi
command -v dpkg-deb >/dev/null || { echo "需要 dpkg-deb（apt install dpkg）" >&2; exit 4; }
WORK="$(mktemp -d)"
trap 'rm -rf "$WORK"' EXIT

PKG_ROOT="$WORK/pkg"
mkdir -p "$PKG_ROOT/DEBIAN" "$PKG_ROOT/opt/audittools" \
         "$PKG_ROOT/usr/bin" "$PKG_ROOT/usr/share/applications" \
         "$PKG_ROOT/usr/share/icons/hicolor/256x256/apps"

cp -a "$PRODUCT_DIR"/. "$PKG_ROOT/opt/audittools/"
# The selected binary and edition resources win over stale product-directory contents.
install -m 0755 "$BINARY" "$PKG_ROOT/opt/audittools/AuditTools"
rm -rf -- "$PKG_ROOT/opt/audittools/config-real" "$PKG_ROOT/opt/audittools/config-defaults"
rm -f -- "$PKG_ROOT/opt/audittools/edition-s1s2"
mkdir -p "$PKG_ROOT/opt/audittools/config-defaults"
for name in "${CONFIG_FILES[@]}" '外部文件-大集中相关数据.xlsx'; do
  cp "$CONFIG_SRC/$name" "$PKG_ROOT/opt/audittools/config-defaults/"
done
cp -a "$CONFIG_SRC/模板文件_命名区域" "$PKG_ROOT/opt/audittools/config-defaults/"
cp "$GUIDE" "$PKG_ROOT/opt/audittools/使用说明.docx"

ln -s /opt/audittools/AuditTools "$PKG_ROOT/usr/bin/audittools"
# Always own an explicit marker so full replaces any previous slim marker too.
if [[ "$EDITION" == "s1s2" ]]; then
  printf 's1s2' > "$PKG_ROOT/opt/audittools/edition-s1s2"
else
  printf 's1s2s3' > "$PKG_ROOT/opt/audittools/edition-s1s2"
fi

# 桌面菜单入口（Kylin/UOS 的应用菜单，类似 Windows 开始菜单）。
# 图标装进 hicolor 主题并在 .desktop 里按名字引用，主题切换时跟随。
install -m 0644 "$PKG_ROOT/opt/audittools/assets/app-icon.png"   "$PKG_ROOT/usr/share/icons/hicolor/256x256/apps/audittools.png"
if [[ "$EDITION" == "s1s2" ]]; then
  DESKTOP_NAME="审核工具（S1S2）"
else
  DESKTOP_NAME="审核工具"
fi
cat > "$PKG_ROOT/usr/share/applications/audittools.desktop" <<DESKTOP
[Desktop Entry]
Version=1.0
Type=Application
Name=${DESKTOP_NAME}
Name[zh_CN]=${DESKTOP_NAME}
GenericName=数据审核工具
GenericName[zh_CN]=数据审核工具
Comment=逐笔统计 / 报表采集 / 大集中统计审核
Comment[zh_CN]=逐笔统计 / 报表采集 / 大集中统计审核
Exec=/usr/bin/audittools
Icon=audittools
Terminal=false
StartupNotify=true
StartupWMClass=AuditTools
Categories=Office;Finance;
Keywords=audit;excel;审核;财务;
DESKTOP
chmod 0644 "$PKG_ROOT/usr/share/applications/audittools.desktop"

# Embed helpers because preinst cannot rely on files from the not-yet-unpacked package.
for phase in preinst prerm; do
  {
    printf '#!/bin/sh\nset -e\n'
    cat "$CONFIG_PRESERVATION"
    if [[ "$phase" == "preinst" ]]; then
      printf 'case "${1:-}" in install|upgrade) audit_config_snapshot /opt/audittools/config-real /var/lib/audittools/config-real ;; esac\n'
    else
      printf 'case "${1:-}" in remove|upgrade|deconfigure) audit_config_snapshot /opt/audittools/config-real /var/lib/audittools/config-real ;; esac\n'
    fi
  } > "$PKG_ROOT/DEBIAN/$phase"
  chmod 0755 "$PKG_ROOT/DEBIAN/$phase"
done
{
  printf '#!/bin/sh\nset -e\n'
  cat "$CONFIG_PRESERVATION"
  printf 'case "${1:-}" in configure) audit_config_restore /opt/audittools/config-real /var/lib/audittools/config-real /opt/audittools/config-defaults ;; esac\n'
} > "$PKG_ROOT/DEBIAN/postinst"
cat >> "$PKG_ROOT/DEBIAN/postinst" <<'POSTINST'
# 刷新桌面数据库，让应用菜单立即出现图标（无 GUI 环境下失败不影响安装）。
if command -v update-desktop-database >/dev/null 2>&1; then
  update-desktop-database -q /usr/share/applications || true
fi
if command -v gtk-update-icon-cache >/dev/null 2>&1; then
  gtk-update-icon-cache -q -t -f /usr/share/icons/hicolor || true
fi
exit 0
POSTINST
chmod 0755 "$PKG_ROOT/DEBIAN/postinst"

cat > "$PKG_ROOT/DEBIAN/postrm" <<'POSTRM'
#!/bin/sh
set -e
if command -v update-desktop-database >/dev/null 2>&1; then
  update-desktop-database -q /usr/share/applications || true
fi
if command -v gtk-update-icon-cache >/dev/null 2>&1; then
  gtk-update-icon-cache -q -t -f /usr/share/icons/hicolor || true
fi
exit 0
POSTRM
chmod 0755 "$PKG_ROOT/DEBIAN/postrm"

# 依赖口径：二进制实际 NEEDED 只有 libc/libm/libpthread/libdl，libgcc 是保险。
# libgcc1 是 Debian 10 / 麒麟 V10 的包名，libgcc-s1 是 Debian 11+ / Ubuntu 22+ 的
# 包名，用可选依赖同时兼容；写死 libgcc-s1 会导致麒麟 V10 上 dpkg 配置失败
# （依赖不满足 → postinst 不执行 → 桌面菜单不刷新、命令不可用）。
cat > "$PKG_ROOT/DEBIAN/control" <<CONTROL
Package: audittools${PKG_SUFFIX}
Version: ${VERSION}
Section: utils
Priority: optional
Architecture: amd64
Maintainer: Azi-lzb <Azi-lzb@users.noreply.github.com>
Installed-Size: $(du -sk "$PKG_ROOT" | cut -f1)
Depends: libc6 (>= 2.28), libgcc1 | libgcc-s1
Recommends: fonts-noto-cjk
Conflicts: ${OTHER_EDITION}, audit-tools, audit-tools-s1-s2, base-audit-v3
Replaces: ${OTHER_EDITION}, audit-tools, audit-tools-s1-s2, base-audit-v3
Description: AuditTools 审核工具（Rust 实现）
 S1 逐笔统计核查、S2 报表采集核对、S3 大集中统计比较与金融表单生成。
 纯 Rust 计算引擎（betteroffice）内置，无需安装 Excel/WPS/LibreOffice/JVM。
 本包与另一 edition 及旧版 Python 审核工具互斥，不支持同时安装。
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
# 产物文件名带打包时间戳（与 Windows 包的命名口径一致），
# 避免连续重打时新旧包混淆——旧包被进程锁定删不掉的场景吃过亏。
STAMP="$(date +%Y%m%d-%H%M%S)"
DEB="$OUT_DIR/AuditTools${PKG_SUFFIX}_${VERSION}_${STAMP}_amd64.deb"
dpkg-deb --root-owner-group --build "$PKG_ROOT" "$DEB" >/dev/null
echo "已生成：$DEB ($(du -h "$DEB" | cut -f1))"
dpkg-deb --info "$DEB" | head -8
