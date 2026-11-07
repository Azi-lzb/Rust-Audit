# Open-Data-Audit Rust project

This repository is an independent Cargo workspace for Audit. Its Excel-Core copies live under `crates/`; project-specific business code, frontend, configuration, fixtures, and acceptance records live here. It has no path dependency on the former combined workspace or PyTools.

## Windows development launchers

- `run.bat` runs only `dist\open-data-audit.exe serve`; it never invokes Cargo. `run-audit-dev.bat` is a compatibility alias.
- `build-release.bat` is the explicit Windows source build. It uses `F:\RustBuild\audit` and copies the Release `open-data-audit.exe` into `dist\` after the build succeeds. Windows S1 formula recalculation uses native Office COM and does not require `lo-convert.exe`. `build-audit-dev.bat` is a compatibility alias.
- `package.bat` creates a Windows x64 ZIP from the existing `dist\open-data-audit.exe`, `run.bat`, this README, and the six local S1/S2/S3 workbooks in `audit-core\config-real\`. It does not compile or include S5, external data, backups, or templates.
- The Audit GUI starts its local web interface from this repository's working directory. An optional `--config <path>` argument selects the S2 configuration workbook.

The original Open-Data-Audit source snapshot was commit
`cac9e1068367bdc46f27685e043aff0cfcd2294c` from `main`. Public placeholder
configuration is in `audit-core/config/`; local real S1/S2/S3 configuration
workbooks belong in the ignored `audit-core/config-real/` directory. The
Windows package places those six files in `dist/config-real/` beside the EXE,
which uses them as the startup defaults.

## Probes

From the independent repository root. The following `cargo run` commands compile before execution; only run them after explicit build/run authorization under `AGENTS.md`:

```text
cargo run -p audit-core --bin umya-probe -- --input <workbook.xlsx> --output-dir <tmp-dir> --report <report.json>
cargo run -p audit-core --bin rust-fastpath-probe -- read --input <workbook.xlsx> --output <dump.json>
cargo run -p audit-core --bin rust-fastpath-probe -- write-smoke --output <new-workbook.xlsx>
```

The Python scripts under `tests/compare/` use OpenPyXL only as a comparison
baseline. They are not runtime dependencies. `tests/TEST_GROUPS.md` is retained
as a clearly marked legacy Python regression map; its old `core/tests` commands
do not apply to this Rust workspace. Markdown findings are in `GPT修改记录/`;
per-run JSON output goes to `probe/reports/`.

> Current feature status: see the repository-root DEVELOPMENT_GOAL.md. The historical slices below describe their original scope, not the complete current implementation.

## Historical S2 CLI candidate evidence

The file-based S2 audit flow can be invoked without Python:

```text
open-data-audit s2 --current-dir <current-period-dir> --previous-dir <previous-period-dir> --config <s2-config.xlsx> --output-dir <results-dir> [--external <external-workbook.xlsx>]
```

Progress events are emitted as JSON lines on stderr; successful runs print a
JSON summary and result-workbook path on stdout. The CLI and its report output
match the read-only Python service on three synthetic scenarios. Real business
templates, Windows-native execution, target Linux/deb, and production routing
remain unverified; see `MIGRATION_INVENTORY.md` and
`probe/reports/s2_service_parity.json`.

The early reports described here cover a synthetic fixture only. Real sanitized workbook and
Kylin/UOS installation acceptance are still required before business migration
or deb packaging.

## Historical S1 template recommendation slice

The `open-data-audit template-suggest` command reads `.xlsx` worksheet names
through the shared namespace-aware OOXML metadata adapter and applies the
original S1 worksheet-set recommendation rules without writing to the template
or input directories:

```text
open-data-audit template-suggest --template-dir <templates-dir> --input-dir <source-dir> [--recursive]
```

This is the template-discovery slice only. It does not validate S1 rules,
recalculate formulas, write audit results, or update history. Four synthetic
cases match the Python baseline; a read-only smoke comparison also matched the
source `config-real` templates and the 2026-07-31 S1 source tree. One WSL2
measurement on that directory tree completed in a 1.37-second median versus
1.76 seconds for the Python baseline; this is a single-sample observation, not
a general performance claim. See
`probe/reports/template_recommendation_parity.json` and
`probe/reports/template_recommendation_real_smoke.json`.

## Historical S1 conditional-format issue extraction slice

`open-data-audit conditional-audit` evaluates supported static conditional
format rules and converts triggers into Audit `Issue` records in Rust. Supply
the S1 metadata already resolved by a caller; optional scan and structure
areas use `Sheet!A1:B20` notation:

```text
open-data-audit conditional-audit --workbook <source.xlsx> --period <period> --batch-id <id> --audit-time <timestamp> [--source-file <original-path>] [--org-code <code>] [--org-name <name>] [--scan-area Sheet1!A1:Z200] [--structure-area Sheet1!A1:Z1] [--developer-view]
```

The single-workbook command prints issue records as JSON on stdout. For a
selected multi-workbook S1 batch, pass a manifest with per-file identity:

```text
open-data-audit conditional-audit-batch --manifest <batch.json> --period <period> --batch-id <id> --audit-time <timestamp> [--scan-area Sheet1!A1:Z200] [--structure-area Sheet1!A1:Z1] [--developer-view]
```

The manifest has a `files` array; each entry supplies `workbook` and may supply
`source_file`, `org_code`, and `org_name`. Relative paths resolve from the
manifest directory. Workbooks are evaluated in parallel by file; returned
summaries and issues retain manifest order and per-file metadata. Any file
failure fails the whole batch without returning a partial issue summary. Both
commands print actual engine and stage timings in ordinary logs;
`--developer-view` also emits per-file JSON progress events on stderr. Formula
cells use stored caches; no recalculation or workbook write occurs.
`--structure-area` lets the caller resolve `check_field`. This is only
the conditional-format issue-generation stage: S1 config/template selection,
ordinary business rules, output workbook writeback, and history are not yet
connected. The synthetic evaluator, issue mapping, and CLI tests pass; one
read-only real-workbook comparison matched all 348 Python issues field by field
(see `probe/reports/s1_conditional_audit_real.json`). Both runs used the same
temporary ASCII basename to work around WSL Unicode argument handling, so
Issue identity with the original workbook filename remains unverified. Full
S1-F01 remains unverified.

## Rust web UI bridge

The maintained Rust + HTML frontend is in audit-core/frontend (or frontend/ from
this directory). The Rust server serves it on loopback and injects the bridge.
From the repository root, run the existing distribution executable with:

~~~powershell
.\run.bat
~~~

An optional --port or --config argument is passed to the serve command. The
launcher requires dist/open-data-audit.exe, which is produced only by an
explicitly authorized build-release.bat run.

The bridge connects S1-F2, S2-F1, S3-F1, S3 current-period checks, S3
explanation export, and S5-F1 to their Rust handlers. S1-F1 and the S3
institution split are not connected as complete workflows. An API being wired
does not establish business acceptance; see DEVELOPMENT_GOAL.md at the
repository root and the historical audit-core/MIGRATION_INVENTORY.md for
current status and scoped evidence.
