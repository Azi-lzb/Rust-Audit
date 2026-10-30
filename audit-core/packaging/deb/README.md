# Linux offline deb package

Run `build.sh` from a Linux x86_64 environment with the locked Cargo sources
available offline and `dpkg-deb` installed. It builds the Audit CLI and
recalculation sidecar as static `x86_64-unknown-linux-musl` executables, then
packages them into `dist/linux/open-data-audit_0.1.0_amd64.deb`. The control
archive uses gzip compression so older `dpkg-deb` versions on the glibc 2.28
baseline can read it.

The static musl build has no glibc runtime floor, so it can run on the target
glibc 2.28 baseline without requiring Python, Go, Excel, WPS, LibreOffice, or a
JVM to launch. S1-F01 still needs an available Excel/LibreOffice calculation
engine at runtime; the application fails clearly when that optional
capability is absent. The deb does not bundle third-party office software.

Validation must separately record the build target, package metadata, static
linkage, a `--version`/`--help` smoke, and the environment used for the smoke.
The Kylin V10 and UOS V20 product environments still require manual acceptance.
