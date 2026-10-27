# Audit independent repository baseline

- Baseline source: `C:\Users\AZI\Desktop\VibeCode\audit-dev`, branch `audit-dev`, HEAD `db4a1cc5bd4660ca145183ac06c652db60578aa9`.
- This worktree had the more complete current Audit business implementation, including the active S1/S3/S5 migration work. Its tracked and untracked work was copied; the source worktree remains unchanged.
- The active Codex Audit snapshot at `C:\Users\AZI\.codex\worktrees\88e4\1.rust-workspace`, branch `codex/windows-operable-entries`, HEAD `ddb61a8e02d744676c9c813869e7532c4618925f`, was retained for comparison and recovery. No wholesale replay was performed; the selected Audit tree already contains the GUI and browser service path needed by the product.
- Targeted Main fix included: commit `846414553f32725ebbda6f14760a5b053cede3c5`, which validates moved-sheet worksheet relationships by sheet identity and actual relationship target. It was applied only to this repository's `crates/excel-editor/src/openpyxl_backend.rs`; no other Main changes were copied.
- Existing executables under `F:\RustBuild\audit` are not used by `run.bat`; it launches only `dist\open-data-audit.exe`. The first Release EXE must be built from this repository using `build-release.bat` after explicit user authorization.
- All five original worktrees and their work-in-progress snapshots are preserved in `F:\RustWorkspaceSplitSnapshots\baseline-20261027`. The repository bundle, original paths/HEADs, status files, tracked WIP patches, and complete non-build tree copies are documented in its `manifest.txt`.
- Cargo build, test, check, run, clippy, and executable validation were intentionally not performed for this split. The targeted relationship fix still requires an explicitly authorized focused regression run.
