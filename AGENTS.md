# Open-Data-Audit Rust project rules

This repository is the independent Rust migration of Open-Data-Audit. It owns its business code, frontend, configuration, Cargo workspace, and local Excel-Core copies. It does not depend on PyTools-Rust or the former combined workspace. The original Python source at `C:\Users\AZI\Desktop\VibeCode\3.open-data-audit` is read-only.

## Project boundaries

- Keep Audit business logic and frontend integration under `audit-core/`. Excel primitives are maintained in this repository's own `crates/`; no shared path outside this repository, Git submodule, or automatic synchronization is allowed.
- Preserve the product scope and migration requirements in `audit-core/AGENTS.md`, `MIGRATION_INVENTORY.md`, and the applicable files under `audit-core/rules/`. Python comparison helpers are baselines only. The isolated openpyxl worker remains a narrow, existing-workbook modification compatibility path and must not perform Audit business logic.
- Keep real or sensitive workbooks under ignored local-only directories. Commit only fixtures explicitly approved for this repository.
- Rules under `rules/` and `audit-core/rules/` apply within this independent repository, with the root AGENTS.md build authorization rule taking precedence over historical test instructions.

## Build and test authorization

- Do not run any command that may compile, build, execute, or install Rust code unless the current user request explicitly authorizes that class of operation. This includes `cargo build/check/test/run/clippy/bench/install`, `rustc`, and scripts or tools that invoke them.
- Implementing or reviewing code is not build authorization. `cargo fmt --check` and static inspection are allowed. Do not start a build from a run script.
- `run.bat` and its `run-audit-dev.bat` compatibility alias launch only the existing `%~dp0dist\open-data-audit.exe`; they never invoke Cargo and must fail clearly if the EXE is absent. `build-release.bat` is the separate, explicit Release build entry, uses `CARGO_TARGET_DIR=F:\RustBuild\audit`, and copies successful outputs into this project's `dist` directory. `package.bat` packages existing `dist` outputs and must not build. All project BAT files resolve the source root through `%~dp0`; do not point them at Main or a Codex worktree.
- Never run `cargo clean` or remove target caches unless the user explicitly requests cache maintenance.

## Validation and delivery

- Record implementation, unit/smoke-test, Windows real-sample, and Linux/deb validation as separate states. Tests written but not run remain unverified.
- Keep new work incremental and tied to a real Audit feature. Public Excel fixes must have an Audit-specific impact and regression evidence; do not import unrelated PyTools changes. Use the agreed manual feature-acceptance workflow; do not launch a full business audit, full Python parity run, or broad Excel-style comparison unless the user asks. Do not remove existing business behavior without a replacement that has been tested.
- Keep Linux delivery requirements in force: x86_64, glibc 2.28 compatibility, offline debs, and no requirement for a user-installed Python, Go, Office, WPS, LibreOffice, or JVM. Optional external engines remain explicit and capability-specific.
