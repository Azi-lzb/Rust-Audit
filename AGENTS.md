# Rust-Audit project rules

## Formal project and boundaries

- The only default Audit development directory is C:\Users\AZI\Desktop\VibeCode\Rust-Audit. Prefer Codex Local for ordinary work. Do not create a Git worktree or branch, work in old Main or audit-dev, leave work in a temporary Codex worktree, or synchronize old Main unless the user explicitly requests it. If a worktree is requested, report its path and branch, ownership of changes, and a safe integration plan; do not leave the result there alone.
- This is an independent Git/Cargo project. Keep Audit business logic, frontend and configuration under audit-core/; maintain its own Excel reader, writer, editor and OOXML patch sources under crates/. Change them only for an actual Audit need, with scoped evidence. Do not add outside path dependencies, submodules, cross-project synchronization or a dependency on PyTools.
- The Python source at C:\Users\AZI\Desktop\VibeCode\3.open-data-audit and old project/worktree/history material are read-only references. Keep real or sensitive workbooks in ignored local-only directories; commit only explicitly approved fixtures.
- Preserve the agreed business scope in audit-core/AUDIT_RUST_V2_BUSINESS_REQUIREMENTS.md, the current status in DEVELOPMENT_GOAL.md, and the historical evidence in audit-core/MIGRATION_INVENTORY.md. Follow audit-core/AGENTS.md and applicable rules/ and audit-core/rules/ subject to this file and current user instructions; historical workspace, mandatory test and parity gates do not override these rules.
- The isolated openpyxl worker is only a capability-gated existing-workbook modification compatibility path. It must not perform Audit business logic. Python comparison helpers are baselines, not product business dependencies.

## Build, test and launch authorization

- Do not run commands that compile, build, execute or install Rust code unless the current user request explicitly authorizes that class of operation. This includes cargo build/check/test/run/clippy/bench/install, rustc, and indirect scripts. Implementation or review is not authorization. Static inspection, Git diff and cargo fmt --check are allowed.
- Do not run a test suite, packaging or a broad comparison without the applicable user authorization. Never run cargo clean or remove target caches without an explicit cache-maintenance request.
- run.bat and its run-audit-dev.bat alias only launch the existing %~dp0dist\open-data-audit.exe and fail clearly if absent; they never invoke Cargo. build-release.bat and its compatibility alias are explicit build entries, use CARGO_TARGET_DIR=F:\RustBuild\audit and copy successful outputs into this project's dist. package.bat packages existing dist outputs without building. All BAT files resolve the project root through %~dp0, never old Main or a Codex worktree.

## Feature workflow and correctness

- Work incrementally on a specified real Audit feature. Do not restart migration, redesign Excel-Core or the GUI, remove existing behavior without a tested replacement, or invent new business features. S1-F3/F4/F5 and S4 remain deferred; S3 V2 F2 and F3 are separate workflows. S5 uses Rust rules and direct calculation rather than unnecessary template-formula injection.
- The default workflow is implementation, static inspection, delivery for user-initiated build and manual testing, then focused fixes from feedback. A missing broad automated comparison is not a reason to withhold an honest manual-test build. Do not call code presence or menu visibility completed acceptance.
- Do not proactively run full Python/Rust parity, routine font/border/style-by-style comparisons, unrelated engine expansion, purposeless complete source audits, repeated architecture unification, or a new long-running migration plan. Preserve data, formulas, critical structure, audit conclusions and input-file safety; conditional formatting that affects audit judgments is business logic, not ordinary appearance.
- Preserve the existing Rust + HTML GUI. A TUI is optional future work; do not introduce Electron, Tauri, WebKitGTK or TypeScript restructuring without a request.
- Linux delivery remains x86_64 Kylin V10 and UOS V20, offline deb, target glibc 2.28, with no required user-installed Python, Go, Office, WPS, LibreOffice or JVM. External calculation or editing engines are explicit, optional and capability-specific; unavailable capability must be reported.

## Git, validation and delivery

- Preserve all existing changes. Do not use destructive Git operations such as reset --hard, clean, force overwrite or deletion of historical work. Do not commit, push, merge, rebase or publish unless explicitly requested.
- Keep implementation, unit/smoke tests, Windows real-sample/manual acceptance and Linux/deb validation as separate states. Written but unrun tests remain unverified; previous evidence proves only its stated version and scope.
- At task completion report the Git repository and branch, actual changed files, important uncommitted work, completed business functionality, outstanding issues, and whether compilation and tests ran. Do not claim acceptance without evidence.
