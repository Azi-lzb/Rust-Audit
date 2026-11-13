#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd -- "$SCRIPT_DIR/../../.." && pwd)"
TARGET="${TARGET:-x86_64-unknown-linux-musl}"
CARGO_TARGET_DIR="${CARGO_TARGET_DIR:-$PROJECT_ROOT/target}"
OUTPUT_DIR="${OUTPUT_DIR:-$PROJECT_ROOT/dist/linux}"
VERSION="${VERSION:-0.1.0}"

command -v cargo >/dev/null 2>&1 || { echo "cargo is required" >&2; exit 127; }
command -v dpkg-deb >/dev/null 2>&1 || { echo "dpkg-deb is required" >&2; exit 127; }
[[ "$TARGET" == "x86_64-unknown-linux-musl" ]] || {
    echo "This package recipe currently supports x86_64-unknown-linux-musl only" >&2
    exit 2
}

export CARGO_TARGET_DIR
cd "$PROJECT_ROOT"
cargo build --offline --locked --release --target "$TARGET" -p audit-core --bin AuditTools

STAGE="$(mktemp -d)"
trap 'rm -rf "$STAGE"' EXIT
mkdir -p \
    "$STAGE/DEBIAN" \
    "$STAGE/usr/bin" \
    "$STAGE/usr/share/doc/open-data-audit" \
    "$STAGE/usr/share/open-data-audit/config"

install -m 0755 \
    "$CARGO_TARGET_DIR/$TARGET/release/AuditTools" \
    "$STAGE/usr/bin/open-data-audit"
if [[ -d "$PROJECT_ROOT/audit-core/config/默认配置" ]]; then
    cp -a "$PROJECT_ROOT/audit-core/config/默认配置" \
        "$STAGE/usr/share/open-data-audit/config/"
fi
if [[ -f "$PROJECT_ROOT/dist/README_RELEASE.md" ]]; then
    install -m 0644 "$PROJECT_ROOT/dist/README_RELEASE.md" \
        "$STAGE/usr/share/doc/open-data-audit/README_RELEASE.md"
fi

cat >"$STAGE/DEBIAN/control" <<EOF
Package: open-data-audit
Version: $VERSION
Section: utils
Priority: optional
Architecture: amd64
Maintainer: Rust-Audit Project
Description: Open-Data-Audit Rust command-line and local web application
 Statically linked musl executables; no Python, Go, office suite, JVM, or
 LibreOffice runtime is required to launch the application. Recalculation
 workflows report an unavailable optional Excel/LibreOffice capability.
EOF

mkdir -p "$OUTPUT_DIR"
PACKAGE_PATH="$OUTPUT_DIR/open-data-audit_${VERSION}_amd64.deb"
[[ ! -e "$PACKAGE_PATH" ]] || {
    echo "Refusing to overwrite existing package: $PACKAGE_PATH" >&2
    exit 3
}
dpkg-deb --root-owner-group -Zgzip --build "$STAGE" "$PACKAGE_PATH"
printf 'Package created: %s\n' "$PACKAGE_PATH"
stat -c 'Package size: %s bytes' "$PACKAGE_PATH"
