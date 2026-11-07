# Open-Data-Audit implementation rules

The repository-root AGENTS.md is authoritative for independent development, build authorization, manual acceptance and Git safety. Historical references to a shared workspace, old Main, worktree coordination, mandatory full parity or unconditional tests in rules/ and old evidence are superseded. The original Python project is read-only.

## Business and scope

- Keep Audit business behavior and its real frontend integration in audit-core/; Excel primitives are this repository's own crates/. Do not depend on pytools-core or an external source tree.
- Use ../DEVELOPMENT_GOAL.md for current feature states and numbering, AUDIT_RUST_V2_BUSINESS_REQUIREMENTS.md for the confirmed business scope, and MIGRATION_INVENTORY.md for dated historical evidence. Update the current state when a feature changes; do not reinterpret older reports as acceptance of the current version.
- Protect original files, configuration and templates. Real data stays local and ignored. Prefer new output paths and explicit failures over silent loss or false success. Conditional-format triggers and unsupported rules affecting business judgments must be reported.
- Preserve S1's independent audit-copy/explicit recalculation requirements. S5 has been removed from the active product per the current user decision; historical S5 evidence is archival only. Existing optional external engines remain capability-gated except for an explicitly selected openpyxl trial, whose candidates must pass output validation before publication; they are not universal prerequisites for unrelated workflows.
- When an Audit operation invokes Excel, WPS, or LibreOffice for conversion/calculation, follow `../rules/13_Rust_Excel读写修改与引擎路由规范.md` §4.5: reuse one owned engine session across sequential workbooks in a logical batch; prefer native Rust COM for validated Windows Excel/WPS operations. For LibreOffice, prefer one isolated process/profile for the batch and reuse UNO only when that operation has a verified UNO path; where a multi-input `soffice --convert-to` batch is the supported path, make one invocation for the whole batch and do not describe it as a persistent UNO session. Do not infer operation support from a different engine or benchmark.
- Do not automatically resume deferred features, expand unrelated Excel capabilities, or start full-source/whole-project comparison. Check only the call chain and evidence relevant to the requested feature.

## Frontend and verification

- The maintained frontend is frontend/web/index.html with frontend/assets/. src/web_bridge.rs serves it on loopback and injects the bridge. Only describe an API as connected when its handler and business path exist; connected is not the same as accepted. Unsupported calls must fail explicitly.
- Follow the root manual-acceptance workflow and authorization rules. Without authorization, compile/build, tests and runtime checks stay unverified; retain historical Windows and Linux evidence with its original scope.
- Keep the isolated openpyxl worker limited to existing-workbook modifications. An explicitly selected unverified trial may write only a candidate copy; failed output validation must prevent publication. Python comparison files under probe/ and tests/compare/ are helpers, not product business logic.
- Record Audit implementation changes under GPT修改记录/ according to rules/08. Document-only governance changes do not change product versions. Do not commit, push, merge, rebase or publish without an explicit user request.
