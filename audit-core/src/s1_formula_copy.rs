//! S1 formula-copy business planning.
//!
//! This module plans and executes the S1 formula-copy stage through Excel-Core,
//! then requires a real calculation sidecar before publishing formula results.

use crate::models::CopyRange;
use crate::s1_template_rules::{S1TemplateRuleSummary, load_s1_template_rule_plan};
use anyhow::{Context, Result, ensure};
use excel_editor::{
    EditExecutionReport, EditOperation, ExcelEditPlan, ExcelModifier, PreservationRequirement,
    RangeStyleMode, WorkbookProfile, inspect_xlsx_profile,
};
use excel_reader::{SheetDump, WorkbookDump, read_workbook};
use regex::Regex;
use std::{
    collections::{BTreeMap, BTreeSet},
    fs,
    path::{Path, PathBuf},
    sync::atomic::{AtomicU64, Ordering},
};

static STAGING_SEQUENCE: AtomicU64 = AtomicU64::new(0);

#[derive(Clone, Debug)]
pub struct S1FormulaCopyEditPlan {
    pub edit_plan: ExcelEditPlan,
    /// Kept separate because an ImportRange's source workbook also has to
    /// match the selected backend's verified feature set.
    pub template_profile: WorkbookProfile,
    /// Optional external workbook profile. Like the template, it is a
    /// cross-workbook ImportRange source and must fit the chosen backend's
    /// verified feature scope.
    pub external_profile: Option<WorkbookProfile>,
    pub external_plan: Option<S1ExternalSheetPlan>,
    pub template_summary: S1TemplateRuleSummary,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct S1ExternalSheetImport {
    pub sheet_name: String,
    pub source_range: String,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct S1ExternalSheetPlan {
    pub workbook: PathBuf,
    /// `formula_references` when static formula references identify sheets,
    /// `formula_dependencies` when imported formulas need further sheets;
    /// otherwise `all_sheets_fallback`, matching the legacy dynamic-reference
    /// safeguard for INDIRECT/OFFSET and similar formulas.
    pub selection_reason: String,
    pub sheets: Vec<S1ExternalSheetImport>,
}

#[derive(Clone, Debug)]
pub struct S1FormulaCopyExecutionReport {
    pub edit: EditExecutionReport,
    pub recalculation: crate::s1_recalculation::S1RecalculationReport,
    pub formula_cells_prepared: usize,
    pub cached_values_cleared: usize,
    pub external_sheets_imported: usize,
    /// Excel COM 打开修复：仅当本次重算路由到 Excel COM 且副本确实被修复时
    /// 才有值；原始报送文件永远不被修改。
    pub excel_open_repair: Option<crate::excel_open_repair::ExcelOpenRepairReport>,
}

#[derive(Debug)]
pub struct S1FormulaCopyManualRecalcReport {
    pub edit: EditExecutionReport,
    pub audit_copy: PathBuf,
    pub reason: String,
    pub formula_cells_prepared: usize,
    pub cached_values_cleared: usize,
    pub external_sheets_imported: usize,
    pub excel_open_repair: Option<crate::excel_open_repair::ExcelOpenRepairReport>,
}

#[derive(Debug)]
pub enum S1FormulaCopyOutcome {
    Recalculated(S1FormulaCopyExecutionReport),
    WaitingManualRecalc(S1FormulaCopyManualRecalcReport),
}

/// Formula-copy candidate held until a whole S1-F1 batch is ready for one
/// LibreOffice launch. The staging guard keeps the prepared workbook alive
/// until batch recalculation and validation finish.
#[derive(Debug)]
struct PreparedS1FormulaCopy {
    _staging: S1FormulaCopyStagingDirectory,
    template_workbook: PathBuf,
    output_workbook: PathBuf,
    cacheless_formula_copy: PathBuf,
    recalculated_candidate: PathBuf,
    template_rule_comments_imported: bool,
    edit: EditExecutionReport,
    formula_cells_prepared: usize,
    cached_values_cleared: usize,
    external_sheets_imported: usize,
    excel_open_repair: Option<crate::excel_open_repair::ExcelOpenRepairReport>,
}

impl PreparedS1FormulaCopy {
    fn batch_pair(&self) -> (PathBuf, PathBuf) {
        (
            self.cacheless_formula_copy.clone(),
            self.recalculated_candidate.clone(),
        )
    }
}

/// Prepare a separate formula-source copy for an editor that cannot retain
/// cached formula values. The returned copy is only an intermediate input;
/// callers must run a real calculation sidecar before reading or publishing
/// formula results. Ordinary shared/array formulas are rejected by the OOXML
/// patcher before an output copy is created.
pub fn prepare_s1_recalculation_source_copy(
    source_workbook: impl AsRef<Path>,
    output_workbook: impl AsRef<Path>,
) -> Result<Option<ooxml_patch::FormulaCachesClearedReport>> {
    let source_workbook = source_workbook.as_ref();
    let output_workbook = output_workbook.as_ref();
    ensure!(source_workbook.is_file(), "S1 重算源工作簿不存在");
    ensure!(xlsx_extension(source_workbook), "S1 重算源目前仅支持 .xlsx");
    ensure!(
        xlsx_extension(output_workbook),
        "S1 重算中间副本必须为 .xlsx"
    );
    ensure!(!output_workbook.exists(), "S1 重算中间副本已存在，拒绝覆盖");
    let workbook = read_workbook(source_workbook).context("读取 S1 重算源公式")?;
    if workbook
        .sheets
        .iter()
        .all(|sheet| sheet.formulas.is_empty())
    {
        return Ok(None);
    }
    ooxml_patch::clear_formula_caches_for_recalculation(source_workbook, output_workbook)
        .context("创建 S1 公式缓存清除中间副本")
        .map(Some)
}

/// Execute only the formula-copy stage of S1-F1. The final output is
/// published only after Excel-Core accepts the exact edit batch and the
/// calculation sidecar reports a real supported engine.
pub fn execute_s1_formula_copy(
    target_workbook: impl AsRef<Path>,
    template_workbook: impl AsRef<Path>,
    output_workbook: impl AsRef<Path>,
    mapping_config: Option<&Path>,
    external_workbook: Option<&Path>,
    sidecar_path: impl AsRef<Path>,
    modifier: &ExcelModifier,
    calc_engine: &str,
) -> Result<S1FormulaCopyExecutionReport> {
    match execute_s1_formula_copy_with_manual_recovery(
        target_workbook,
        template_workbook,
        output_workbook,
        mapping_config,
        external_workbook,
        Some(sidecar_path.as_ref()),
        modifier,
        calc_engine,
    )? {
        S1FormulaCopyOutcome::Recalculated(report) => Ok(report),
        S1FormulaCopyOutcome::WaitingManualRecalc(report) => anyhow::bail!(
            "S1 公式副本已保留，等待人工重算：{}；副本：{}",
            report.reason,
            report.audit_copy.display()
        ),
    }
}

/// Prepare a formula audit copy, attempt real recalculation when the sidecar
/// is available, and preserve a cache-cleared independent copy when it is not
/// available or recalculation fails. The source workbook is never modified.
pub fn execute_s1_formula_copy_with_manual_recovery(
    target_workbook: impl AsRef<Path>,
    template_workbook: impl AsRef<Path>,
    output_workbook: impl AsRef<Path>,
    mapping_config: Option<&Path>,
    external_workbook: Option<&Path>,
    sidecar_path: Option<&Path>,
    modifier: &ExcelModifier,
    calc_engine: &str,
) -> Result<S1FormulaCopyOutcome> {
    execute_s1_formula_copy_with_recalculator(
        target_workbook,
        template_workbook,
        output_workbook,
        mapping_config,
        external_workbook,
        sidecar_path.is_some(),
        modifier,
        crate::excel_open_repair::excel_com_engine_requested(calc_engine),
        |input, output| {
            crate::s1_recalculation::recalculate_s1_formula_copy(
                sidecar_path.context("公式重算 sidecar 不可用")?,
                input,
                output,
                calc_engine,
            )
        },
    )
}

/// S1-F1 batch adapter which reuses one already-created calculation session.
/// On Windows the native COM worker and Office process start lazily.
pub(crate) fn execute_s1_formula_copy_with_manual_recovery_session(
    target_workbook: impl AsRef<Path>,
    template_workbook: impl AsRef<Path>,
    output_workbook: impl AsRef<Path>,
    mapping_config: Option<&Path>,
    external_workbook: Option<&Path>,
    sidecar_available: bool,
    modifier: &ExcelModifier,
    session: &mut crate::s1_recalculation::S1RecalculationSession,
) -> Result<S1FormulaCopyOutcome> {
    execute_s1_formula_copy_with_recalculator(
        target_workbook,
        template_workbook,
        output_workbook,
        mapping_config,
        external_workbook,
        sidecar_available,
        modifier,
        crate::excel_open_repair::excel_com_engine_requested(session.engine_request()),
        |input, output| session.recalculate(input, output),
    )
}

fn execute_s1_formula_copy_with_recalculator<F>(
    target_workbook: impl AsRef<Path>,
    template_workbook: impl AsRef<Path>,
    output_workbook: impl AsRef<Path>,
    mapping_config: Option<&Path>,
    external_workbook: Option<&Path>,
    sidecar_available: bool,
    modifier: &ExcelModifier,
    excel_open_repair_requested: bool,
    mut recalculate: F,
) -> Result<S1FormulaCopyOutcome>
where
    F: FnMut(&Path, &Path) -> Result<crate::s1_recalculation::S1RecalculationReport>,
{
    let mut prepared = prepare_s1_formula_copy(
        target_workbook,
        template_workbook,
        output_workbook,
        mapping_config,
        external_workbook,
        modifier,
    )?;
    // Excel COM 对报送文件中两类非规范内容拒绝打开：含强制引号字符的表名
    // 裸引用、冻结窗格 topLeftCell 落在冻结区内。重算前在审核副本上做无害
    // 规范化（公式语义不变），让 Excel 优先的 auto/excel 引擎不再被迫回退；
    // 修复失败或无需修复时保持副本原样，绝不打断审核流程。
    if excel_open_repair_requested {
        let repair =
            crate::excel_open_repair::repair_audit_copy_for_excel(&prepared.cacheless_formula_copy);
        if repair.changed {
            prepared.excel_open_repair = Some(repair);
        }
    }

    if !sidecar_available {
        return preserve_prepared_s1_formula_copy(
            prepared,
            "未找到公式重算 sidecar；审核副本已保留，尚未计算公式".to_owned(),
        );
    }

    let recalculated_candidate = prepared.recalculated_candidate.clone();
    match recalculate(&prepared.cacheless_formula_copy, &recalculated_candidate) {
        Ok(recalculation) => complete_prepared_s1_formula_copy(prepared, recalculation),
        Err(error) => preserve_prepared_s1_formula_copy(
            prepared,
            format!("自动公式重算失败；审核副本已保留：{error:#}"),
        ),
    }
}

/// Prepare one formula-copy candidate without starting a calculation engine.
/// This lets the Linux S1-F1 GUI prepare every independent workbook before it
/// sends the whole batch to LibreOffice in a single engine launch.
fn prepare_s1_formula_copy(
    target_workbook: impl AsRef<Path>,
    template_workbook: impl AsRef<Path>,
    output_workbook: impl AsRef<Path>,
    mapping_config: Option<&Path>,
    external_workbook: Option<&Path>,
    modifier: &ExcelModifier,
) -> Result<PreparedS1FormulaCopy> {
    let target_workbook = target_workbook.as_ref();
    let template_workbook = template_workbook.as_ref();
    let output_workbook = output_workbook.as_ref();
    let target_canonical = fs::canonicalize(target_workbook).context("规范化 S1 源工作簿路径")?;
    let template_canonical =
        fs::canonicalize(template_workbook).context("规范化 S1 公式模板路径")?;
    ensure!(
        !same_path(&target_canonical, &template_canonical),
        "S1 公式复制源工作簿与模板必须是不同文件"
    );
    if let Some(external_workbook) = external_workbook {
        let external_canonical =
            fs::canonicalize(external_workbook).context("规范化 S1 外部数据工作簿路径")?;
        ensure!(
            !same_path(&target_canonical, &external_canonical),
            "S1 外部数据工作簿不能与源报送工作簿相同"
        );
    }
    ensure!(
        xlsx_extension(output_workbook),
        "S1 公式审核结果目前仅支持 .xlsx"
    );
    ensure!(!output_workbook.exists(), "S1 公式审核结果已存在，拒绝覆盖");
    let output_parent = output_workbook
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    ensure!(output_parent.is_dir(), "S1 公式审核输出目录不存在");

    let staging = S1FormulaCopyStagingDirectory::create(output_parent)?;
    let cacheless_target = staging.path().join("formula-source.xlsx");
    let cache_report = prepare_s1_recalculation_source_copy(target_workbook, &cacheless_target)
        .context("准备 S1 公式重算源副本")?;
    let cached_values_cleared = cache_report
        .as_ref()
        .map(|report| report.cache_values_removed)
        .unwrap_or_default();
    let editor_input = if cache_report.is_some() {
        cacheless_target.as_path()
    } else {
        target_workbook
    };
    let editor_output = staging.path().join("formula-copy.xlsx");
    let prepared = build_s1_formula_copy_edit_plan(
        editor_input,
        template_workbook,
        &editor_output,
        mapping_config,
        external_workbook,
    )?;
    let template_rule_comments_imported =
        prepared
            .edit_plan
            .operations
            .iter()
            .any(|operation| match operation {
                EditOperation::ImportRange {
                    source_workbook,
                    source_sheet,
                    target_sheet,
                    copy_comments: true,
                    ..
                }
                | EditOperation::ImportRangeWithFormulaTranslation {
                    source_workbook,
                    source_sheet,
                    target_sheet,
                    copy_comments: true,
                    ..
                } => {
                    source_sheet.eq_ignore_ascii_case("审核规则")
                        && target_sheet.eq_ignore_ascii_case("审核规则")
                        && fs::canonicalize(source_workbook)
                            .is_ok_and(|source| same_path(&source, &template_canonical))
                }
                _ => false,
            });
    let external_sheets_imported = prepared
        .external_plan
        .as_ref()
        .map(|plan| plan.sheets.len())
        .unwrap_or_default();
    let edit = modifier
        .execute(&prepared.edit_plan)
        .context("通过 Excel-Core 复制 S1 模板公式和外部数据")?;
    ensure!(
        edit.output == editor_output,
        "Excel-Core 编辑器返回的 S1 候选路径与请求不一致"
    );
    validate_s1_formula_copy_candidate(&prepared, &editor_output)
        .context("验证 S1 模板和外部数据复制结果")?;
    // Re-clear caches after the copy so formulas imported from the template
    // cannot carry stale template results into the manual-recalculation path.
    let cacheless_formula_copy = staging.path().join("formula-copy-cacheless.xlsx");
    let final_cache_report =
        prepare_s1_recalculation_source_copy(&editor_output, &cacheless_formula_copy)?
            .context("S1 公式复制结果中没有可重算的公式")?;
    let formula_cells_prepared = final_cache_report.formula_cells;
    let cached_values_cleared =
        cached_values_cleared.saturating_add(final_cache_report.cache_values_removed);

    let recalculated_candidate = staging.path().join("formula-copy-recalculated.xlsx");
    Ok(PreparedS1FormulaCopy {
        _staging: staging,
        template_workbook: template_workbook.to_path_buf(),
        output_workbook: output_workbook.to_path_buf(),
        cacheless_formula_copy,
        recalculated_candidate,
        template_rule_comments_imported,
        edit,
        formula_cells_prepared,
        cached_values_cleared,
        external_sheets_imported,
        excel_open_repair: None,
    })
}

fn complete_prepared_s1_formula_copy(
    prepared: PreparedS1FormulaCopy,
    mut recalculation: crate::s1_recalculation::S1RecalculationReport,
) -> Result<S1FormulaCopyOutcome> {
    let final_candidate = if prepared.template_rule_comments_imported {
        let comments_restored_candidate = prepared
            ._staging
            .path()
            .join("formula-copy-comments-restored.xlsx");
        let report = ooxml_patch::restore_legacy_comments_for_sheet(
            &prepared.template_workbook,
            &recalculation.output,
            &comments_restored_candidate,
            "审核规则",
        )
        .context("在真实公式重算后恢复模板审核规则表批注")?;
        if report.restored_sheets.is_empty() {
            recalculation.output.clone()
        } else {
            comments_restored_candidate
        }
    } else {
        recalculation.output.clone()
    };
    fs::hard_link(&final_candidate, &prepared.output_workbook).with_context(|| {
        format!(
            "安全发布 S1 重算审核副本失败（输出不得覆盖）：{}",
            prepared.output_workbook.display()
        )
    })?;
    recalculation.output = prepared.output_workbook.clone();
    Ok(S1FormulaCopyOutcome::Recalculated(
        S1FormulaCopyExecutionReport {
            edit: prepared.edit,
            recalculation,
            formula_cells_prepared: prepared.formula_cells_prepared,
            cached_values_cleared: prepared.cached_values_cleared,
            external_sheets_imported: prepared.external_sheets_imported,
            excel_open_repair: prepared.excel_open_repair,
        },
    ))
}

fn preserve_prepared_s1_formula_copy(
    prepared: PreparedS1FormulaCopy,
    reason: String,
) -> Result<S1FormulaCopyOutcome> {
    let excel_open_repair = prepared.excel_open_repair;
    preserve_s1_manual_copy(
        &prepared.cacheless_formula_copy,
        &prepared.output_workbook,
        prepared.edit,
        reason,
        prepared.formula_cells_prepared,
        prepared.cached_values_cleared,
        prepared.external_sheets_imported,
        excel_open_repair,
    )
}

/// Prepare independent formula copies, recalculate the eligible set through
/// one LibreOffice batch invocation, then restore each workbook's normal
/// audit-copy and manual-recovery semantics. A failed batch keeps the prepared
/// formula copies for S1-F3 recovery instead of retrying one cold LO process
/// per workbook.
pub(crate) fn execute_s1_formula_copy_batch_with_manual_recovery(
    items: &[(PathBuf, PathBuf, PathBuf)],
    mapping_config: Option<&Path>,
    external_workbook: Option<&Path>,
    sidecar_path: &Path,
    modifier: &ExcelModifier,
) -> Vec<Result<S1FormulaCopyOutcome>> {
    let mut outcomes: Vec<Option<Result<S1FormulaCopyOutcome>>> =
        (0..items.len()).map(|_| None).collect();
    let mut prepared = Vec::new();
    for (index, (target, template, output)) in items.iter().enumerate() {
        match prepare_s1_formula_copy(
            target,
            template,
            output,
            mapping_config,
            external_workbook,
            modifier,
        ) {
            Ok(candidate) => prepared.push((index, candidate)),
            Err(error) => outcomes[index] = Some(Err(error)),
        }
    }
    if !prepared.is_empty() {
        let pairs = prepared
            .iter()
            .map(|(_, candidate): &(usize, PreparedS1FormulaCopy)| candidate.batch_pair())
            .collect::<Vec<_>>();
        match crate::s1_recalculation::recalculate_s1_formula_copy_batch(sidecar_path, &pairs, "lo")
        {
            Ok(reports) if reports.len() == prepared.len() => {
                for ((index, candidate), report) in prepared.into_iter().zip(reports) {
                    let outcome = match report {
                        Some(report) => complete_prepared_s1_formula_copy(candidate, report),
                        None => preserve_prepared_s1_formula_copy(
                            candidate,
                            "LibreOffice 批次未通过该工作簿的重算验证；审核副本已保留，等待人工重算".to_owned(),
                        ),
                    };
                    outcomes[index] = Some(outcome);
                }
            }
            Ok(reports) => {
                let reason = format!(
                    "LibreOffice 批次返回工作簿数量不一致（{} / {}）；审核副本已保留，等待人工重算",
                    reports.len(),
                    prepared.len()
                );
                for (index, candidate) in prepared {
                    outcomes[index] =
                        Some(preserve_prepared_s1_formula_copy(candidate, reason.clone()));
                }
            }
            Err(error) => {
                let reason =
                    format!("LibreOffice 批次重算失败；审核副本已保留，等待人工重算：{error:#}");
                for (index, candidate) in prepared {
                    outcomes[index] =
                        Some(preserve_prepared_s1_formula_copy(candidate, reason.clone()));
                }
            }
        }
    }
    outcomes
        .into_iter()
        .map(|outcome| outcome.unwrap_or_else(|| Err(anyhow::anyhow!("S1-F1 批次副本状态缺失"))))
        .collect()
}

fn preserve_s1_manual_copy(
    cacheless_formula_copy: &Path,
    output_workbook: &Path,
    edit: EditExecutionReport,
    reason: String,
    formula_cells_prepared: usize,
    cached_values_cleared: usize,
    external_sheets_imported: usize,
    excel_open_repair: Option<crate::excel_open_repair::ExcelOpenRepairReport>,
) -> Result<S1FormulaCopyOutcome> {
    fs::hard_link(cacheless_formula_copy, output_workbook).with_context(|| {
        format!(
            "安全保留待人工重算副本失败（不覆盖已有文件）：{}",
            output_workbook.display()
        )
    })?;
    Ok(S1FormulaCopyOutcome::WaitingManualRecalc(
        S1FormulaCopyManualRecalcReport {
            edit,
            audit_copy: output_workbook.to_path_buf(),
            reason,
            formula_cells_prepared,
            cached_values_cleared,
            external_sheets_imported,
            excel_open_repair,
        },
    ))
}

#[derive(Debug)]
struct S1FormulaCopyStagingDirectory(PathBuf);

impl S1FormulaCopyStagingDirectory {
    fn create(parent: &Path) -> Result<Self> {
        let parent = fs::canonicalize(parent)
            .with_context(|| format!("解析 S1 公式审核输出目录 {}", parent.display()))?;
        ensure!(parent.is_dir(), "S1 公式审核输出路径的父级不是目录");
        for _ in 0..128 {
            let sequence = STAGING_SEQUENCE.fetch_add(1, Ordering::Relaxed);
            let path = parent.join(format!(".audit-s1-f1-{}-{sequence}", std::process::id()));
            match fs::create_dir(&path) {
                Ok(()) => return Ok(Self(path)),
                Err(error) if error.kind() == std::io::ErrorKind::AlreadyExists => continue,
                Err(error) => {
                    return Err(error)
                        .with_context(|| format!("创建 S1 公式审核临时目录 {}", path.display()));
                }
            }
        }
        anyhow::bail!("无法创建唯一的 S1 公式审核临时目录")
    }

    fn path(&self) -> &Path {
        &self.0
    }
}

impl Drop for S1FormulaCopyStagingDirectory {
    fn drop(&mut self) {
        let _ = fs::remove_dir_all(&self.0);
    }
}

/// Build a non-destructive, workbook-level plan for copying formula ranges
/// from one S1 template into a new audit-workbook copy.
///
/// The output must not exist. This function only reads the target/template and
/// creates an edit plan; a production caller must still route it through
/// ExcelModifier and a Passed capability for the exact operation batch.
pub fn build_s1_formula_copy_edit_plan(
    target_workbook: impl AsRef<Path>,
    template_workbook: impl AsRef<Path>,
    output_workbook: impl AsRef<Path>,
    mapping_config: Option<&Path>,
    external_workbook: Option<&Path>,
) -> Result<S1FormulaCopyEditPlan> {
    let target_workbook = target_workbook.as_ref();
    let template_workbook = template_workbook.as_ref();
    let output_workbook = output_workbook.as_ref();

    ensure!(target_workbook.is_file(), "S1 公式复制目标工作簿不存在");
    ensure!(template_workbook.is_file(), "S1 公式模板工作簿不存在");
    ensure!(
        xlsx_extension(target_workbook),
        "S1 公式复制目标目前仅支持 .xlsx"
    );
    ensure!(
        xlsx_extension(template_workbook),
        "S1 公式复制模板目前仅支持 .xlsx"
    );
    ensure!(
        xlsx_extension(output_workbook),
        "S1 公式复制输出目前仅支持 .xlsx"
    );
    ensure!(
        !output_workbook.exists(),
        "S1 公式复制输出文件已存在，拒绝覆盖"
    );
    let output_parent = output_workbook
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    ensure!(output_parent.is_dir(), "S1 公式复制输出目录不存在");

    let target_canonical = fs::canonicalize(target_workbook).context("规范化 S1 目标工作簿路径")?;
    let template_canonical =
        fs::canonicalize(template_workbook).context("规范化 S1 模板工作簿路径")?;
    ensure!(
        !same_path(&target_canonical, &template_canonical),
        "S1 公式复制目标与模板必须是不同工作簿"
    );

    let template = load_s1_template_rule_plan(template_workbook, mapping_config)
        .context("解析 S1 公式复制模板规则")?;
    ensure!(
        !template.definition.copy_ranges.is_empty(),
        "S1 模板没有可复制的公式区域"
    );
    let target_profile = inspect_xlsx_profile(target_workbook).context("检查 S1 目标工作簿结构")?;
    let target_sheet_names = excel_reader::read_xlsx_sheet_names_ooxml(target_workbook)
        .context("读取 S1 目标工作簿工作表名")?;
    let rule_sheet_copy_range = template.rule_sheet_copy_range.clone();
    for range in &template.definition.copy_ranges {
        ensure!(
            target_sheet_names
                .iter()
                .any(|sheet_name| sheet_name == &range.sheet_name),
            "S1 公式复制目标工作簿缺少模板映射工作表“{}”",
            range.sheet_name
        );
    }
    let template_profile =
        inspect_xlsx_profile(template_workbook).context("检查 S1 公式模板结构")?;
    let (external_plan, external_profile) = if let Some(external_path) = external_workbook {
        ensure!(external_path.is_file(), "S1 外部数据工作簿不存在");
        ensure!(
            xlsx_extension(external_path),
            "S1 外部数据工作簿目前仅支持 .xlsx；暂不把 .xlsm 当作已验证的外部表导入能力"
        );
        let external_canonical =
            fs::canonicalize(external_path).context("规范化 S1 外部数据工作簿路径")?;
        ensure!(
            !same_path(&target_canonical, &external_canonical),
            "S1 外部数据工作簿不能与源报送工作簿相同"
        );
        ensure!(
            !same_path(&template_canonical, &external_canonical),
            "S1 外部数据工作簿不能与公式模板相同"
        );

        let template_dump = read_workbook(template_workbook).context("读取模板公式引用")?;
        let external_dump = read_workbook(&external_canonical).context("读取 S1 外部数据工作簿")?;
        // Only formulas in configured `校验区域` ranges are audit rules. Other
        // formulas already present in the template must not cause unrelated
        // external sheets to be copied into every source workbook. The
        // dependency selector still follows references transitively inside
        // the selected external lookup sheets.
        let validation_formulas =
            formulas_in_ranges(&template_dump, &template.definition.copy_ranges)?;
        let (selected_names, selection_reason) = select_external_sheets(
            validation_formulas.iter().map(String::as_str),
            &external_dump.sheets,
        )?;
        let target_names = target_sheet_names
            .iter()
            .map(|name| name.to_lowercase())
            .collect::<BTreeSet<_>>();
        let mut imports = Vec::with_capacity(selected_names.len());
        for sheet_name in selected_names {
            ensure!(
                !target_names.contains(&sheet_name.to_lowercase()),
                "S1 外部工作表“{sheet_name}”与源报送工作簿重名，拒绝覆盖"
            );
            let sheet = external_dump
                .sheets
                .iter()
                .find(|sheet| sheet.name == sheet_name)
                .with_context(|| format!("S1 外部数据缺少工作表“{sheet_name}”"))?;
            imports.push(S1ExternalSheetImport {
                sheet_name,
                source_range: used_range(sheet)?,
            });
        }
        let external_profile =
            inspect_xlsx_profile(&external_canonical).context("检查 S1 外部数据工作簿结构")?;
        (
            Some(S1ExternalSheetPlan {
                workbook: external_canonical,
                selection_reason,
                sheets: imports,
            }),
            Some(external_profile),
        )
    } else {
        (None, None)
    };

    if let (Some(rule_sheet), Some(external_plan)) =
        (rule_sheet_copy_range.as_ref(), external_plan.as_ref())
    {
        ensure!(
            external_plan.sheets.iter().all(|sheet| !sheet
                .sheet_name
                .eq_ignore_ascii_case(&rule_sheet.sheet_name)),
            "S1 外部数据工作表与模板审核规则表重名：{}",
            rule_sheet.sheet_name
        );
    }

    let external_operation_count = external_plan
        .as_ref()
        .map(|plan| plan.sheets.len() * 2)
        .unwrap_or_default();
    let rule_sheet_operation_count = if rule_sheet_copy_range.is_some() {
        3
    } else {
        0
    };
    let mut operations = Vec::with_capacity(
        external_operation_count
            + template.definition.copy_ranges.len()
            + rule_sheet_operation_count
            + 1,
    );
    if let Some(external_plan) = external_plan.as_ref() {
        // External lookup sheets must exist before template formulas are
        // imported and recalculated, matching the Python workflow ordering.
        for sheet in &external_plan.sheets {
            operations.push(EditOperation::CreateSheet {
                name: sheet.sheet_name.clone(),
            });
            operations.push(EditOperation::ImportRange {
                source_workbook: external_plan.workbook.clone(),
                source_sheet: sheet.sheet_name.clone(),
                source_range: sheet.source_range.clone(),
                target_sheet: sheet.sheet_name.clone(),
                // Keep external coordinates unchanged: a used range may
                // start at C4, while template/dependent formulas still refer to C4.
                target_start: range_target_start(&sheet.source_range)?,
                copy_values: true,
                copy_formulas: true,
                style_mode: RangeStyleMode::None,
                copy_comments: false,
                copy_merges: false,
                copy_row_heights: false,
                copy_column_widths: false,
                copy_row_hidden: false,
                copy_column_hidden: false,
                copy_outline_levels: false,
                copy_all_row_dimensions: false,
                copy_all_column_dimensions: false,
                materialize_blank_cells: true,
                hide_zero: false,
            });
        }
    }
    let mut copied_ranges = BTreeSet::new();
    for range in &template.definition.copy_ranges {
        let identity = (range.sheet_name.clone(), range.address.to_ascii_uppercase());
        if !copied_ranges.insert(identity) {
            continue;
        }
        let target_start = range_target_start(&range.address)
            .with_context(|| format!("解析公式复制目标坐标 {}", range.address))?;
        operations.push(EditOperation::ImportRange {
            source_workbook: template_canonical.clone(),
            source_sheet: range.sheet_name.clone(),
            source_range: range.address.clone(),
            target_sheet: range.sheet_name.clone(),
            target_start,
            copy_values: true,
            copy_formulas: true,
            // The validation-region copy contract includes legacy cell comment
            // objects as well as formulas and literal contents.
            style_mode: RangeStyleMode::None,
            copy_comments: true,
            copy_merges: false,
            copy_row_heights: false,
            copy_column_widths: false,
            copy_row_hidden: false,
            copy_column_hidden: false,
            copy_outline_levels: false,
            copy_all_row_dimensions: false,
            copy_all_column_dimensions: false,
            // The S1 copy contract overwrites the complete mapped rectangle,
            // including blank source cells.
            materialize_blank_cells: true,
            hide_zero: false,
        });
    }
    if let Some(rule_sheet) = rule_sheet_copy_range.as_ref() {
        if let Some(existing_name) = target_sheet_names
            .iter()
            .find(|name| name.eq_ignore_ascii_case(&rule_sheet.sheet_name))
        {
            operations.push(EditOperation::DeleteSheet {
                name: existing_name.clone(),
            });
        }
        operations.push(EditOperation::CreateSheet {
            name: rule_sheet.sheet_name.clone(),
        });
        operations.push(EditOperation::ImportRange {
            source_workbook: template_canonical.clone(),
            source_sheet: rule_sheet.sheet_name.clone(),
            source_range: rule_sheet.address.clone(),
            target_sheet: rule_sheet.sheet_name.clone(),
            target_start: range_target_start(&rule_sheet.address)
                .context("解析模板审核规则表复制坐标")?,
            copy_values: true,
            copy_formulas: true,
            style_mode: RangeStyleMode::None,
            copy_comments: true,
            copy_merges: false,
            copy_row_heights: false,
            copy_column_widths: false,
            copy_row_hidden: false,
            copy_column_hidden: false,
            copy_outline_levels: false,
            copy_all_row_dimensions: false,
            copy_all_column_dimensions: false,
            materialize_blank_cells: true,
            hide_zero: false,
        });
    }
    operations.push(EditOperation::SetWorkbookCalculationProperties {
        full_calc_on_load: Some(true),
        force_full_calc: Some(true),
    });

    // Existing values in merged non-anchor cells are still workbook data.
    // Do not inherit the legacy Python save behavior that drops them. A
    // backend that cannot preserve these values must reject this plan.
    let mut drop_allowed = BTreeMap::<String, BTreeSet<PreservationRequirement>>::new();
    if let Some(existing_name) = rule_sheet_copy_range.as_ref().and_then(|rule_sheet| {
        target_sheet_names
            .iter()
            .find(|name| name.eq_ignore_ascii_case(&rule_sheet.sheet_name))
    }) {
        drop_allowed.insert(
            existing_name.clone(),
            [
                PreservationRequirement::CellValuesAndTypes,
                PreservationRequirement::FormulaText,
                PreservationRequirement::CommentsAndPositions,
                PreservationRequirement::MergedRanges,
                PreservationRequirement::ConditionalFormatting,
                PreservationRequirement::DataValidations,
                PreservationRequirement::DefinedNames,
                PreservationRequirement::Relationships,
                PreservationRequirement::SheetIdentityAndOrder,
            ]
            .into_iter()
            .collect(),
        );
    }

    let edit_plan = ExcelEditPlan {
        request_id: "audit-s1-formula-copy".to_owned(),
        input: target_workbook.to_owned(),
        output: output_workbook.to_owned(),
        format: target_profile.format,
        workbook: target_profile,
        operations,
        // Keep the contracts required by formula execution, issue extraction
        // and workbook structure. Ordinary presentation and package-byte
        // normalization are intentionally nonblocking for this workflow.
        preserve: [
            PreservationRequirement::CellValuesAndTypes,
            PreservationRequirement::FormulaText,
            PreservationRequirement::CommentsAndPositions,
            PreservationRequirement::MergedRanges,
            PreservationRequirement::ConditionalFormatting,
            PreservationRequirement::DataValidations,
            PreservationRequirement::DefinedNames,
            PreservationRequirement::Relationships,
            PreservationRequirement::SheetIdentityAndOrder,
        ]
        .into_iter()
        .collect(),
        drop_allowed,
    };
    edit_plan
        .validate()
        .context("校验 S1 公式复制 ExcelEditPlan")?;

    Ok(S1FormulaCopyEditPlan {
        edit_plan,
        template_profile,
        external_profile,
        external_plan,
        template_summary: template.summary,
    })
}

/// Check the copied S1 workbook against the exact source ranges before any
/// external formula engine is allowed to recalculate it. Formula text and
/// literal input values are compared independently; saved formula caches are
/// deliberately excluded because they are cleared before recalculation.
fn validate_s1_formula_copy_candidate(
    prepared: &S1FormulaCopyEditPlan,
    candidate_path: &Path,
) -> Result<()> {
    let source =
        read_workbook(&prepared.edit_plan.input).context("读取 S1 源工作簿以验证复制结果")?;
    let candidate = read_workbook(candidate_path).context("重新读取 S1 公式复制候选工作簿")?;

    let mut expected_sheet_order = Vec::with_capacity(source.sheets.len());
    let mut expected_formulas = BTreeMap::<String, BTreeMap<String, String>>::new();
    let mut expected_values = BTreeMap::<String, BTreeMap<String, serde_json::Value>>::new();
    let mut expected_comments = excel_reader::read_ooxml_comments(&prepared.edit_plan.input)
        .context("读取 S1 源工作簿批注以验证复制结果")?;
    let mut source_comments = BTreeMap::new();
    let mut imported_ranges = BTreeMap::<String, Vec<(u32, u32, u32, u32)>>::new();
    for sheet in &source.sheets {
        expected_sheet_order.push(sheet.name.clone());
        expected_formulas.insert(sheet.name.clone(), normalized_formula_map(sheet)?);
        expected_values.insert(sheet.name.clone(), normalized_literal_value_map(sheet)?);
    }

    let mut source_workbooks = BTreeMap::<PathBuf, WorkbookDump>::new();
    for operation in &prepared.edit_plan.operations {
        match operation {
            EditOperation::DeleteSheet { name } => {
                let index = expected_sheet_order
                    .iter()
                    .position(|existing| existing.eq_ignore_ascii_case(name))
                    .with_context(|| format!("S1 复制计划要删除的工作表不存在：{name:?}"))?;
                let removed_name = expected_sheet_order.remove(index);
                expected_formulas.remove(&removed_name);
                expected_values.remove(&removed_name);
                expected_comments.remove(&removed_name);
            }
            EditOperation::CreateSheet { name } => {
                ensure!(
                    !expected_sheet_order
                        .iter()
                        .any(|existing| existing.eq_ignore_ascii_case(name)),
                    "S1 复制计划中的新工作表与源工作表重名：{name:?}"
                );
                expected_sheet_order.push(name.clone());
                expected_formulas.insert(name.clone(), BTreeMap::new());
                expected_values.insert(name.clone(), BTreeMap::new());
                expected_comments.insert(name.clone(), BTreeMap::new());
            }
            EditOperation::ImportRange {
                source_workbook,
                source_sheet,
                source_range,
                target_sheet,
                target_start,
                copy_values,
                copy_formulas,
                copy_comments,
                materialize_blank_cells,
                ..
            } => {
                ensure!(
                    *copy_values && *copy_formulas && *materialize_blank_cells,
                    "S1 公式复制要求同时复制值和公式，并覆盖映射区域中的空白单元格"
                );
                if !source_workbooks.contains_key(source_workbook) {
                    source_workbooks.insert(
                        source_workbook.clone(),
                        read_workbook(source_workbook).with_context(|| {
                            format!("读取 S1 公式复制来源工作簿 {}", source_workbook.display())
                        })?,
                    );
                }
                let import_source = source_workbooks
                    .get(source_workbook)
                    .context("S1 公式复制来源工作簿缓存丢失")?;
                let import_sheet = import_source
                    .sheets
                    .iter()
                    .find(|sheet| sheet.name == *source_sheet)
                    .with_context(|| format!("S1 公式复制来源工作簿缺少工作表 {source_sheet:?}"))?;
                let (source_min_row, source_min_col, source_max_row, source_max_col) =
                    parse_a1_range(source_range)?;
                let (target_min_row, target_min_col) = parse_a1_cell(target_start)?;
                let target_max_row = target_min_row
                    .checked_add(source_max_row - source_min_row)
                    .context("S1 公式复制目标行超出坐标范围")?;
                let target_max_col = target_min_col
                    .checked_add(source_max_col - source_min_col)
                    .context("S1 公式复制目标列超出坐标范围")?;
                imported_ranges
                    .entry(target_sheet.clone())
                    .or_default()
                    .push((
                        target_min_row,
                        target_min_col,
                        target_max_row,
                        target_max_col,
                    ));

                let formulas = expected_formulas
                    .get_mut(target_sheet)
                    .with_context(|| format!("S1 复制目标工作表不存在：{target_sheet:?}"))?;
                let values = expected_values
                    .get_mut(target_sheet)
                    .with_context(|| format!("S1 复制目标工作表不存在：{target_sheet:?}"))?;
                let comments = expected_comments.entry(target_sheet.clone()).or_default();
                remove_cells_in_range(
                    formulas,
                    target_min_row,
                    target_min_col,
                    target_max_row,
                    target_max_col,
                )?;
                remove_cells_in_range(
                    values,
                    target_min_row,
                    target_min_col,
                    target_max_row,
                    target_max_col,
                )?;

                if *copy_comments {
                    remove_comments_in_range(
                        comments,
                        target_min_row,
                        target_min_col,
                        target_max_row,
                        target_max_col,
                    )?;
                    if !source_comments.contains_key(source_workbook) {
                        source_comments.insert(
                            source_workbook.clone(),
                            excel_reader::read_ooxml_comments(source_workbook).with_context(
                                || {
                                    format!(
                                        "读取 S1 批注复制来源工作簿 {}",
                                        source_workbook.display()
                                    )
                                },
                            )?,
                        );
                    }
                    let import_comments = source_comments
                        .get(source_workbook)
                        .context("S1 批注复制来源缓存丢失")?;
                    for (address, comment) in import_comments
                        .get(source_sheet)
                        .into_iter()
                        .flat_map(|comments| comments.iter())
                    {
                        let (row, col) = parse_a1_cell(address)?;
                        if row < source_min_row
                            || row > source_max_row
                            || col < source_min_col
                            || col > source_max_col
                        {
                            continue;
                        }
                        let target_row = target_min_row + row - source_min_row;
                        let target_col = target_min_col + col - source_min_col;
                        comments.insert(a1_cell(target_row, target_col), comment.clone());
                    }
                }

                let source_formulas = normalized_formula_map(import_sheet)?;
                for (address, formula) in &source_formulas {
                    let (row, col) = parse_a1_cell(address)?;
                    if row < source_min_row
                        || row > source_max_row
                        || col < source_min_col
                        || col > source_max_col
                    {
                        continue;
                    }
                    let target_row = target_min_row + row - source_min_row;
                    let target_col = target_min_col + col - source_min_col;
                    formulas.insert(a1_cell(target_row, target_col), formula.clone());
                }

                let source_values = normalized_literal_value_map(import_sheet)?;
                for (address, value) in source_values {
                    let (row, col) = parse_a1_cell(&address)?;
                    if row < source_min_row
                        || row > source_max_row
                        || col < source_min_col
                        || col > source_max_col
                    {
                        continue;
                    }
                    let target_row = target_min_row + row - source_min_row;
                    let target_col = target_min_col + col - source_min_col;
                    values.insert(a1_cell(target_row, target_col), value);
                }
            }
            EditOperation::SetWorkbookCalculationProperties { .. } => {}
            other => anyhow::bail!("S1 公式复制校验遇到非预期编辑操作：{:?}", other.kind()),
        }
    }

    let actual_sheet_order = candidate
        .sheets
        .iter()
        .map(|sheet| sheet.name.as_str())
        .collect::<Vec<_>>();
    let expected_names = expected_sheet_order
        .iter()
        .map(String::as_str)
        .collect::<Vec<_>>();
    ensure!(
        actual_sheet_order == expected_names,
        "S1 复制后的工作表名称或顺序与计划不一致：预期 {} 张，实际 {} 张",
        expected_names.len(),
        actual_sheet_order.len()
    );

    for (expected_name, actual_sheet) in expected_sheet_order.iter().zip(&candidate.sheets) {
        let actual_formulas = normalized_formula_map(actual_sheet)?;
        let expected = expected_formulas
            .get(expected_name)
            .with_context(|| format!("S1 公式校验缺少预期工作表 {expected_name:?}"))?;
        let differences = differing_cell_addresses(expected, &actual_formulas);
        let ranges = imported_ranges
            .get(expected_name)
            .map(Vec::as_slice)
            .unwrap_or_default();
        let mut imported_formula_differences = Vec::new();
        for address in differences {
            if address_in_imported_ranges(&address, ranges)? {
                imported_formula_differences.push(address);
            }
        }
        let mut formula_mismatch_details = Vec::new();
        if !imported_formula_differences.is_empty() {
            formula_mismatch_details.push(format!(
                "编辑计划导入区域公式与预期不一致（{}!{}）",
                expected_name,
                imported_formula_differences
                    .iter()
                    .take(8)
                    .cloned()
                    .collect::<Vec<_>>()
                    .join(", ")
            ));
        }
        ensure!(
            formula_mismatch_details.is_empty(),
            "S1 审核副本完整性校验失败：{}；该检查是复制保护校验，不是业务审核异常",
            formula_mismatch_details.join("；")
        );

        let actual_values = normalized_literal_value_map(actual_sheet)?;
        let expected = expected_values
            .get(expected_name)
            .with_context(|| format!("S1 值校验缺少预期工作表 {expected_name:?}"))?;
        let differences = differing_import_value_addresses(expected, &actual_values);
        let mut imported_value_differences = Vec::new();
        let mut untouched_value_differences = Vec::new();
        for address in differences {
            if address_in_imported_ranges(&address, ranges)? {
                imported_value_differences.push(address);
            } else {
                untouched_value_differences.push(address);
            }
        }
        let mut value_mismatch_details = Vec::new();
        if !imported_value_differences.is_empty() {
            value_mismatch_details.push(format!(
                "编辑计划导入区域文本或数值与预期不一致（{}!{}）",
                expected_name,
                imported_value_differences
                    .iter()
                    .take(8)
                    .cloned()
                    .collect::<Vec<_>>()
                    .join(", ")
            ));
        }
        if !untouched_value_differences.is_empty() {
            value_mismatch_details.push(format!(
                "源工作簿未导入区域文本或数值被后端改写（{}!{}）",
                expected_name,
                untouched_value_differences
                    .iter()
                    .take(8)
                    .cloned()
                    .collect::<Vec<_>>()
                    .join(", ")
            ));
        }
        ensure!(
            value_mismatch_details.is_empty(),
            "S1 审核副本完整性校验失败：{}；该检查是复制保护校验，不是业务审核异常",
            value_mismatch_details.join("；")
        );
    }

    let actual_comments = excel_reader::read_ooxml_comments(candidate_path)
        .context("读取 S1 审核副本批注以验证复制结果")?;
    for sheet_name in &expected_sheet_order {
        let expected = expected_comments
            .get(sheet_name)
            .cloned()
            .unwrap_or_default();
        let actual = actual_comments.get(sheet_name).cloned().unwrap_or_default();
        let addresses = expected
            .keys()
            .chain(actual.keys())
            .cloned()
            .collect::<BTreeSet<_>>();
        let differences = addresses
            .into_iter()
            .filter(|address| expected.get(address) != actual.get(address))
            .collect::<Vec<_>>();
        ensure!(
            differences.is_empty(),
            "S1 审核副本完整性校验失败：批注作者、正文或位置与复制计划不一致（{}!{}）；该检查是复制保护校验，不是业务审核异常",
            sheet_name,
            differences
                .into_iter()
                .take(8)
                .collect::<Vec<_>>()
                .join(", ")
        );
    }
    Ok(())
}

fn remove_comments_in_range(
    comments: &mut BTreeMap<String, excel_reader::WorkbookComment>,
    min_row: u32,
    min_col: u32,
    max_row: u32,
    max_col: u32,
) -> Result<()> {
    let addresses = comments.keys().cloned().collect::<Vec<_>>();
    for address in addresses {
        let (row, col) = parse_a1_cell(&address)?;
        if row >= min_row && row <= max_row && col >= min_col && col <= max_col {
            comments.remove(&address);
        }
    }
    Ok(())
}

fn normalized_formula_map(sheet: &SheetDump) -> Result<BTreeMap<String, String>> {
    sheet
        .formulas
        .iter()
        .map(|(address, formula)| {
            let (row, col) = parse_a1_cell(address)
                .with_context(|| format!("解析公式单元格坐标 {}!{address}", sheet.name))?;
            Ok((a1_cell(row, col), formula.clone()))
        })
        .collect()
}

fn normalized_literal_value_map(sheet: &SheetDump) -> Result<BTreeMap<String, serde_json::Value>> {
    let formulas = normalized_formula_map(sheet)?;
    let mut values = BTreeMap::new();
    for (address, value) in &sheet.values {
        let (row, col) = parse_a1_cell(address)
            .with_context(|| format!("解析值单元格坐标 {}!{address}", sheet.name))?;
        let normalized = a1_cell(row, col);
        if !formulas.contains_key(&normalized) {
            values.insert(normalized, value.clone());
        }
    }
    Ok(values)
}

fn parse_a1_range(address: &str) -> Result<(u32, u32, u32, u32)> {
    let range = range_without_sheet(address).replace('$', "");
    let mut parts = range.split(':');
    let start = parts.next().unwrap_or_default().trim();
    let end = parts.next().unwrap_or(start).trim();
    ensure!(parts.next().is_none(), "S1 复制区域不是有效 A1 范围");
    let (min_row, min_col) = parse_a1_cell(start)?;
    let (max_row, max_col) = parse_a1_cell(end)?;
    ensure!(
        min_row <= max_row && min_col <= max_col,
        "S1 复制区域起止坐标顺序无效"
    );
    Ok((min_row, min_col, max_row, max_col))
}

fn remove_cells_in_range<T>(
    cells: &mut BTreeMap<String, T>,
    min_row: u32,
    min_col: u32,
    max_row: u32,
    max_col: u32,
) -> Result<()> {
    let mut to_remove = Vec::new();
    for address in cells.keys() {
        let (row, col) = parse_a1_cell(address)?;
        if (min_row..=max_row).contains(&row) && (min_col..=max_col).contains(&col) {
            to_remove.push(address.clone());
        }
    }
    for address in to_remove {
        cells.remove(&address);
    }
    Ok(())
}

fn differing_cell_addresses<T: PartialEq>(
    expected: &BTreeMap<String, T>,
    actual: &BTreeMap<String, T>,
) -> Vec<String> {
    expected
        .keys()
        .chain(actual.keys())
        .collect::<BTreeSet<_>>()
        .into_iter()
        .filter(|address| expected.get(*address) != actual.get(*address))
        .cloned()
        .collect()
}

fn differing_import_value_addresses(
    expected: &BTreeMap<String, serde_json::Value>,
    actual: &BTreeMap<String, serde_json::Value>,
) -> Vec<String> {
    expected
        .keys()
        .chain(actual.keys())
        .collect::<BTreeSet<_>>()
        .into_iter()
        .filter(|address| {
            let expected_value = expected.get(*address);
            let actual_value = actual.get(*address);
            !equivalent_import_value(expected_value, actual_value)
        })
        .cloned()
        .collect()
}

fn address_in_imported_ranges(address: &str, ranges: &[(u32, u32, u32, u32)]) -> Result<bool> {
    let (row, column) = parse_a1_cell(address)?;
    Ok(ranges
        .iter()
        .any(|(min_row, min_column, max_row, max_column)| {
            (*min_row..=*max_row).contains(&row) && (*min_column..=*max_column).contains(&column)
        }))
}

fn equivalent_import_value(
    expected: Option<&serde_json::Value>,
    actual: Option<&serde_json::Value>,
) -> bool {
    if expected == actual {
        return true;
    }
    match (expected, actual) {
        (Some(expected), None) => is_empty_text_value(expected),
        (None, Some(actual)) => is_empty_text_value(actual),
        (Some(expected), Some(actual)) => {
            if is_empty_text_value(expected) || is_empty_text_value(actual) {
                return false;
            }
            match (excel_numeric_value(expected), excel_numeric_value(actual)) {
                (Some(expected), Some(actual)) => expected == actual,
                _ => false,
            }
        }
        _ => false,
    }
}

/// Calamine may classify an unchanged Excel serial as a date when the source
/// cell has a date number format and as a plain number when an imported range
/// intentionally omits styles. Excel stores both as the same numeric value;
/// compare that underlying value without requiring presentation formatting.
fn excel_numeric_value(value: &serde_json::Value) -> Option<f64> {
    matches!(
        value.get("type").and_then(serde_json::Value::as_str),
        Some("number" | "integer" | "datetime_serial")
    )
    .then(|| value.get("value").and_then(serde_json::Value::as_f64))
    .flatten()
}

fn is_empty_text_value(value: &serde_json::Value) -> bool {
    value.get("type").and_then(serde_json::Value::as_str) == Some("string")
        && value.get("value").and_then(serde_json::Value::as_str) == Some("")
}

fn select_external_sheets<'a>(
    formulas: impl Iterator<Item = &'a str>,
    external_sheets: &[SheetDump],
) -> Result<(Vec<String>, String)> {
    let reference = Regex::new(r#"(?:'(?P<quoted>(?:[^']|'')+)'|(?P<plain>[\p{L}\p{N}_]+))!"#)
        .context("编译 S1 外部工作表公式引用匹配器")?;
    let available = external_sheets
        .iter()
        .map(|sheet| (sheet.name.to_lowercase(), sheet.name.as_str()))
        .collect::<BTreeMap<_, _>>();
    let formulas = formulas.collect::<Vec<_>>();
    let all_sheets = || {
        external_sheets
            .iter()
            .map(|sheet| sheet.name.clone())
            .collect::<Vec<_>>()
    };
    let dynamic_reference = |formula: &str| {
        let upper = formula.to_ascii_uppercase();
        upper.contains("INDIRECT(") || upper.contains("OFFSET(")
    };
    if formulas.iter().any(|formula| dynamic_reference(formula)) {
        return Ok((all_sheets(), "all_sheets_fallback".to_owned()));
    }
    let mut selected = Vec::<String>::new();
    for formula in formulas {
        for captures in reference.captures_iter(formula) {
            let requested = captures
                .name("quoted")
                .or_else(|| captures.name("plain"))
                .map(|name| name.as_str().replace("''", "'"))
                .unwrap_or_default();
            if let Some(actual) = available.get(&requested.to_lowercase()).copied() {
                if !selected.iter().any(|name| name.as_str() == actual) {
                    selected.push(actual.to_owned());
                }
            }
        }
    }
    if !selected.is_empty() {
        let initially_selected = selected.len();
        let mut index = 0;
        while index < selected.len() {
            let sheet = external_sheets
                .iter()
                .find(|sheet| sheet.name == selected[index])
                .context("外部工作表依赖不存在")?;
            for formula in sheet.formulas.values() {
                if dynamic_reference(formula) {
                    return Ok((all_sheets(), "all_sheets_fallback".to_owned()));
                }
                for captures in reference.captures_iter(formula) {
                    let requested = captures
                        .name("quoted")
                        .or_else(|| captures.name("plain"))
                        .map(|name| name.as_str().replace("''", "'"))
                        .unwrap_or_default();
                    if let Some(actual) = available.get(&requested.to_lowercase()).copied() {
                        if !selected.iter().any(|name| name.as_str() == actual) {
                            selected.push(actual.to_owned());
                        }
                    }
                }
            }
            index += 1;
        }
        let reason = if selected.len() == initially_selected {
            "formula_references"
        } else {
            "formula_dependencies"
        };
        return Ok((selected, reason.to_owned()));
    }
    Ok((all_sheets(), "all_sheets_fallback".to_owned()))
}

fn formulas_in_ranges(workbook: &WorkbookDump, ranges: &[CopyRange]) -> Result<Vec<String>> {
    let sheets = workbook
        .sheets
        .iter()
        .map(|sheet| (sheet.name.as_str(), sheet))
        .collect::<BTreeMap<_, _>>();
    let mut formulas = Vec::new();
    for range in ranges {
        let sheet = sheets
            .get(range.sheet_name.as_str())
            .with_context(|| format!("S1 校验区域对应的模板工作表不存在：{}", range.sheet_name))?;
        let (min_row, min_col, max_row, max_col) =
            parse_a1_range(&range.address).with_context(|| {
                format!(
                    "S1 校验区域地址无效：{}!{}",
                    range.sheet_name, range.address
                )
            })?;
        for (address, formula) in &sheet.formulas {
            let (row, column) = parse_a1_cell(address).with_context(|| {
                format!("S1 模板公式单元格地址无效：{}!{address}", range.sheet_name)
            })?;
            if row >= min_row
                && row <= max_row
                && column >= min_col
                && column <= max_col
                && !formula.trim().is_empty()
            {
                formulas.push(formula.clone());
            }
        }
    }
    Ok(formulas)
}

fn used_range(sheet: &SheetDump) -> Result<String> {
    let mut bounds: Option<(u32, u32, u32, u32)> = None;
    for address in sheet.values.keys().chain(sheet.formulas.keys()) {
        let (row, column) = parse_a1_cell(address)
            .with_context(|| format!("解析 S1 外部工作表 {} 的单元格 {address}", sheet.name))?;
        bounds = Some(match bounds {
            Some((min_row, min_column, max_row, max_column)) => (
                min_row.min(row),
                min_column.min(column),
                max_row.max(row),
                max_column.max(column),
            ),
            None => (row, column, row, column),
        });
    }
    let Some((min_row, min_column, max_row, max_column)) = bounds else {
        return Ok("A1".to_owned());
    };
    // Excel writes a single used cell as `B3` in the sheet dimension, not
    // `B3:B3`; keep the minimal form for a one-cell range.
    if min_row == max_row && min_column == max_column {
        return Ok(a1_cell(min_row, min_column));
    }
    Ok(format!(
        "{}:{}",
        a1_cell(min_row, min_column),
        a1_cell(max_row, max_column)
    ))
}

fn xlsx_extension(path: &Path) -> bool {
    path.extension()
        .and_then(|extension| extension.to_str())
        .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx"))
}

#[cfg(windows)]
fn same_path(left: &Path, right: &Path) -> bool {
    left.to_string_lossy()
        .eq_ignore_ascii_case(&right.to_string_lossy())
}

#[cfg(not(windows))]
fn same_path(left: &Path, right: &Path) -> bool {
    left == right
}

fn range_target_start(address: &str) -> Result<String> {
    let range = range_without_sheet(address).replace('$', "");
    let first_cell = range.split(':').next().unwrap_or(&range).trim();
    let (row, column) = parse_a1_cell(first_cell).context("区域没有有效的起始单元格")?;
    Ok(a1_cell(row, column))
}

fn range_without_sheet(address: &str) -> &str {
    address.rsplit('!').next().unwrap_or(address).trim()
}

fn parse_a1_cell(address: &str) -> Result<(u32, u32)> {
    let mut row = 0u32;
    let mut column = 0u32;
    let mut reading_row = false;
    for character in address.trim().chars() {
        if character == '$' {
            continue;
        }
        if character.is_ascii_alphabetic() && !reading_row {
            column = column
                .checked_mul(26)
                .and_then(|value| {
                    value.checked_add(character.to_ascii_uppercase() as u32 - 'A' as u32 + 1)
                })
                .context("单元格列坐标溢出")?;
        } else if character.is_ascii_digit() {
            reading_row = true;
            row = row
                .checked_mul(10)
                .and_then(|value| value.checked_add(character as u32 - '0' as u32))
                .context("单元格行坐标溢出")?;
        } else {
            anyhow::bail!("无效 A1 单元格坐标");
        }
    }
    ensure!(reading_row && row > 0 && column > 0, "无效 A1 单元格坐标");
    Ok((row, column))
}

fn a1_cell(row: u32, mut column: u32) -> String {
    let mut letters = Vec::new();
    while column > 0 {
        let digit = (column - 1) % 26;
        letters.push((b'A' + digit as u8) as char);
        column = (column - 1) / 26;
    }
    format!("{}{}", letters.into_iter().rev().collect::<String>(), row)
}

#[cfg(test)]
mod tests {
    use super::*;
    use excel_editor::{
        BackendCapability, CapabilityEvidence, CapabilityStatus, EditBackend, EditOperationKind,
        EngineId, OoxmlImportBackend, OpenpyxlWorkerConfig, OpenpyxlXlsxBackend,
        OpenpyxlXlsxValidator, OutputValidator, UmyaXlsxBackend, UmyaXlsxValidator, WorkbookFormat,
    };
    use excel_writer::Workbook;
    use std::{
        fs,
        path::PathBuf,
        sync::{
            Arc,
            atomic::{AtomicU64, Ordering},
        },
        time::{SystemTime, UNIX_EPOCH},
    };

    static TEMP_DIR_COUNTER: AtomicU64 = AtomicU64::new(0);

    struct NoopValidator;

    impl OutputValidator for NoopValidator {
        fn validate(&self, _plan: &ExcelEditPlan, _candidate_output: &Path) -> Result<()> {
            Ok(())
        }
    }

    fn passed_umya_modifier_for(plan: &ExcelEditPlan) -> ExcelModifier {
        let platform = "s1-formula-copy-test".to_owned();
        let capability = BackendCapability {
            engine: EngineId::UmyaSpreadsheet,
            status: CapabilityStatus::Passed,
            format: plan.format,
            operations: plan.operations.iter().map(EditOperation::kind).collect(),
            supported_features: plan.workbook.features_present.clone(),
            preserved: plan.preserve.clone(),
            platform: platform.clone(),
            evidence: Some(CapabilityEvidence {
                sample_id: "s1-formula-copy-synthetic".to_owned(),
                report_path: "in-memory test capability".to_owned(),
                adapter_version: "excel-editor test adapter".to_owned(),
                engine_version: "umya-spreadsheet test engine".to_owned(),
                platform: platform.clone(),
            }),
        };
        let backend =
            Arc::new(UmyaXlsxBackend::new(capability).expect("construct test Umya backend"));
        ExcelModifier::new(platform, vec![backend], Arc::new(UmyaXlsxValidator))
            .expect("construct test Excel modifier")
    }

    fn temporary_directory() -> PathBuf {
        for _ in 0..100 {
            let nonce = SystemTime::now()
                .duration_since(UNIX_EPOCH)
                .expect("system clock after Unix epoch")
                .as_nanos();
            let sequence = TEMP_DIR_COUNTER.fetch_add(1, Ordering::Relaxed);
            let path = std::env::temp_dir().join(format!(
                "audit_s1_formula_copy_plan_{}_{}_{}",
                std::process::id(),
                nonce,
                sequence
            ));
            match fs::create_dir(&path) {
                Ok(()) => return path,
                Err(error) if error.kind() == std::io::ErrorKind::AlreadyExists => continue,
                Err(error) => panic!("create formula-copy test directory: {error}"),
            }
        }
        panic!("could not allocate a unique formula-copy test directory")
    }

    fn write_template(path: &Path) {
        write_template_with_formula(path, "=IF(TRUE,\"错误|字段|指标|说明\",\"通过\")");
    }

    fn write_template_with_formula(path: &Path, formula: &str) {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet.set_name("检查表").expect("name formula sheet");
        sheet.write_formula(0, 0, formula).expect("write formula");
        workbook
            .define_name("校验区域", "='检查表'!$A$1:$B$2")
            .expect("define formula-copy range");
        workbook.save(path).expect("save formula template");
    }

    fn write_external_workbook(path: &Path) {
        let mut workbook = Workbook::new();
        let lookup = workbook.add_worksheet();
        lookup
            .set_name("外部维表")
            .expect("name external lookup sheet");
        lookup
            .write_string(0, 0, "编码")
            .expect("write external lookup header");
        lookup
            .write_string(0, 1, "名称")
            .expect("write external lookup header");
        lookup
            .write_string(1, 0, "A-01")
            .expect("write external lookup key");
        lookup
            .write_string(1, 1, "样例名称")
            .expect("write external lookup value");
        let unused = workbook.add_worksheet();
        unused.set_name("未引用表").expect("name unused sheet");
        unused
            .write_string(0, 0, "未引用")
            .expect("write unused external value");
        workbook.save(path).expect("save external workbook");
    }

    fn write_structured_template_with_duplicate_ranges(path: &Path) {
        let mut workbook = Workbook::new();
        {
            let rules = workbook.add_worksheet();
            rules
                .set_name("审核规则")
                .expect("name structured rule worksheet");
            for (column, header) in crate::template::REQUIRED_HEADERS.iter().enumerate() {
                rules
                    .write_string(0, column as u16, *header)
                    .expect("write structured rule header");
            }
            for (row, rule_id) in [(1, "R-01"), (2, "R-02")] {
                for (column, value) in [
                    rule_id,
                    "是",
                    "报表甲",
                    "检查表",
                    "A1",
                    "B1",
                    "错误",
                    "核对错误项",
                ]
                .iter()
                .enumerate()
                {
                    rules
                        .write_string(row, column as u16, *value)
                        .expect("write structured rule row");
                }
            }
        }
        let formulas = workbook.add_worksheet();
        formulas.set_name("检查表").expect("name formula worksheet");
        formulas
            .write_formula(0, 0, "=IF(TRUE,\"错误|字段|指标|说明\",\"通过\")")
            .expect("write structured formula");
        workbook
            .save(path)
            .expect("save duplicate-range structured template");
        let mut workbook = umya_spreadsheet::reader::xlsx::read(path)
            .expect("reopen structured template to add rule-sheet comment");
        let rules = workbook
            .get_sheet_by_name_mut("审核规则")
            .expect("reopen structured rule sheet");
        let mut comment = umya_spreadsheet::Comment::default();
        comment.new_comment("A1");
        comment.set_author("测试");
        comment.set_text_string("规则表说明");
        rules.add_comments(comment);
        umya_spreadsheet::writer::xlsx::write(&workbook, path)
            .expect("save structured template with rule-sheet comment");
    }

    fn write_target(path: &Path) {
        write_target_with_sheet(path, "检查表");
    }

    fn add_comment(path: &Path, sheet_name: &str, cell: &str, author: &str, text: &str) {
        let mut workbook = umya_spreadsheet::reader::xlsx::read(path)
            .expect("reopen workbook to add a legacy cell comment");
        let sheet = workbook
            .get_sheet_by_name_mut(sheet_name)
            .expect("find comment worksheet");
        let mut comment = umya_spreadsheet::Comment::default();
        comment.new_comment(cell);
        comment.set_author(author);
        comment.set_text_string(text);
        sheet.add_comments(comment);
        umya_spreadsheet::writer::xlsx::write(&workbook, path)
            .expect("save workbook with legacy cell comment");
    }

    fn write_target_with_rule_sheet(path: &Path) {
        let mut workbook = Workbook::new();
        let data = workbook.add_worksheet();
        data.set_name("检查表").expect("name target data sheet");
        data.write_string(0, 0, "target value")
            .expect("write target data value");
        let rules = workbook.add_worksheet();
        rules
            .set_name("审核规则")
            .expect("name existing target rule sheet");
        rules
            .write_string(0, 0, "obsolete target rule")
            .expect("write existing target rule value");
        workbook
            .save(path)
            .expect("save target workbook with rule sheet");
    }

    fn write_formula_target(path: &Path) {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet.set_name("检查表").expect("name target formula sheet");
        sheet
            .write_formula(0, 0, "=1+1")
            .expect("write target formula");
        workbook.save(path).expect("save target formula workbook");
    }

    fn write_target_with_formula_outside_copy_range(path: &Path) {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet.set_name("检查表").expect("name target sheet");
        sheet
            .write_string(0, 0, "target value")
            .expect("write target value");
        sheet
            .write_formula(0, 2, "=SUM(A1:B1)")
            .expect("write ordinary source formula outside validation range");
        workbook.save(path).expect("save target workbook");
    }

    fn write_target_with_sheet(path: &Path, sheet_name: &str) {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet.set_name(sheet_name).expect("name target sheet");
        sheet
            .write_string(0, 0, "target value")
            .expect("write target value");
        workbook.save(path).expect("save target workbook");
    }

    #[test]
    fn builds_one_protected_workbook_plan_for_template_ranges() {
        let directory = temporary_directory();
        let target = directory.join("target.xlsx");
        let template = directory.join("template.xlsx");
        let output = directory.join("output.xlsx");
        write_target(&target);
        write_template(&template);

        let prepared = build_s1_formula_copy_edit_plan(&target, &template, &output, None, None)
            .expect("build S1 formula-copy plan");

        assert_eq!(prepared.edit_plan.format, WorkbookFormat::Xlsx);
        assert_eq!(prepared.template_summary.copy_ranges, 1);
        assert_eq!(prepared.edit_plan.operations.len(), 2);
        assert!(
            prepared.edit_plan.drop_allowed.is_empty(),
            "S1 formula copy must not allow a backend to drop merged-cell payloads"
        );
        assert_eq!(
            prepared
                .edit_plan
                .operations
                .iter()
                .filter(|operation| operation.kind() == EditOperationKind::ImportRange)
                .count(),
            1
        );
        assert!(matches!(
            &prepared.edit_plan.operations[0],
            EditOperation::ImportRange {
                source_range,
                target_start,
                copy_values: true,
                copy_formulas: true,
                copy_comments: true,
                materialize_blank_cells: true,
                style_mode: RangeStyleMode::None,
                ..
            } if source_range == "A1:B2" && target_start == "A1"
        ));
        assert!(matches!(
            prepared.edit_plan.operations.last(),
            Some(EditOperation::SetWorkbookCalculationProperties {
                full_calc_on_load: Some(true),
                force_full_calc: Some(true),
            })
        ));
        for requirement in [
            PreservationRequirement::Styles,
            PreservationRequirement::RowAndColumnDimensions,
            PreservationRequirement::PageSettings,
            PreservationRequirement::HeaderFooterText,
            PreservationRequirement::UnmodifiedPackageParts,
        ] {
            assert!(
                !prepared.edit_plan.preserve.contains(&requirement),
                "ordinary presentation/package normalization must not gate S1 content copy: {requirement:?}"
            );
        }
        for requirement in [
            PreservationRequirement::CellValuesAndTypes,
            PreservationRequirement::FormulaText,
            PreservationRequirement::CommentsAndPositions,
            PreservationRequirement::MergedRanges,
            PreservationRequirement::ConditionalFormatting,
            PreservationRequirement::DefinedNames,
            PreservationRequirement::Relationships,
            PreservationRequirement::SheetIdentityAndOrder,
        ] {
            assert!(
                prepared.edit_plan.preserve.contains(&requirement),
                "business content/structure must remain protected: {requirement:?}"
            );
        }
        assert!(!output.exists(), "plan construction must not write output");

        fs::remove_dir_all(directory).expect("remove formula-copy test directory");
    }

    #[test]
    fn prepares_a_separate_cacheless_formula_copy_without_changing_the_source() {
        let directory = temporary_directory();
        let source = directory.join("source.xlsx");
        let cacheless = directory.join("cacheless.xlsx");
        write_formula_target(&source);
        let original_bytes = fs::read(&source).expect("read original source bytes");
        let original_formulas = read_workbook(&source)
            .expect("read original formulas")
            .sheets[0]
            .formulas
            .clone();

        let report = prepare_s1_recalculation_source_copy(&source, &cacheless)
            .expect("prepare independent formula-cache-free copy")
            .expect("formula source should create an intermediate copy");

        assert_eq!(report.formula_cells, 1);
        assert_eq!(
            fs::read(&source).expect("reread original source"),
            original_bytes
        );
        assert_eq!(
            read_workbook(&cacheless)
                .expect("read cacheless copy formulas")
                .sheets[0]
                .formulas,
            original_formulas
        );
        assert!(cacheless.is_file());
        fs::remove_dir_all(directory).expect("remove formula-copy test directory");
    }

    #[test]
    fn editor_with_no_passed_route_fails_before_publishing_or_changing_source() {
        let directory = temporary_directory();
        let target = directory.join("target.xlsx");
        let template = directory.join("template.xlsx");
        let output = directory.join("result.xlsx");
        write_formula_target(&target);
        write_template(&template);
        let original_bytes = fs::read(&target).expect("read source workbook bytes");
        let modifier = ExcelModifier::new("test-platform", Vec::new(), Arc::new(NoopValidator))
            .expect("create editor with no registered capabilities");

        let error = execute_s1_formula_copy(
            &target,
            &template,
            &output,
            None,
            None,
            directory.join("missing-sidecar.exe"),
            &modifier,
            "auto",
        )
        .expect_err("no Passed editor route must stop the copy stage");

        assert!(format!("{error:#}").contains("no qualified backend"));
        assert_eq!(
            fs::read(&target).expect("reread source workbook"),
            original_bytes
        );
        assert!(
            !output.exists(),
            "a rejected edit must not publish final output"
        );
        assert_eq!(
            fs::read_dir(&directory)
                .expect("read output directory")
                .filter_map(std::result::Result::ok)
                .filter(|entry| entry
                    .file_name()
                    .to_string_lossy()
                    .starts_with(".audit-s1-f1-"))
                .count(),
            0,
            "formula-copy staging directory must be removed on rejection"
        );
        fs::remove_dir_all(directory).expect("remove formula-copy test directory");
    }

    #[test]
    fn manual_recalc_copy_path_runs_runtime_formula_and_value_validation() {
        let directory = temporary_directory();
        let target = directory.join("target.xlsx");
        let template = directory.join("template.xlsx");
        let plan_output = directory.join("capability-plan.xlsx");
        let output = directory.join("audit-copy.xlsx");
        write_target_with_sheet(&target, "检查表");
        write_template(&template);
        let target_before = fs::read(&target).expect("read source before S1 copy");
        let template_before = fs::read(&template).expect("read template before S1 copy");
        let plan = build_s1_formula_copy_edit_plan(&target, &template, &plan_output, None, None)
            .expect("build synthetic S1 capability plan");
        let modifier = passed_umya_modifier_for(&plan.edit_plan);

        let outcome = execute_s1_formula_copy_with_manual_recovery(
            &target, &template, &output, None, None, None, &modifier, "auto",
        )
        .expect("execute and preserve validated cacheless copy");
        let S1FormulaCopyOutcome::WaitingManualRecalc(report) = outcome else {
            panic!("missing sidecar must remain a waiting-manual-recalculation result");
        };
        assert_eq!(report.audit_copy, output);
        assert_eq!(report.formula_cells_prepared, 1);
        assert!(report.reason.contains("未找到公式重算 sidecar"));
        let copied = read_workbook(&output).expect("reopen preserved formula copy");
        assert_eq!(
            copied.sheets[0].formulas.get("A1").unwrap(),
            &read_workbook(&template)
                .expect("read template formulas")
                .sheets[0]
                .formulas["A1"]
        );
        assert_eq!(fs::read(&target).unwrap(), target_before);
        assert_eq!(fs::read(&template).unwrap(), template_before);
        assert!(
            fs::read_dir(&directory)
                .expect("read output directory")
                .filter_map(std::result::Result::ok)
                .all(|entry| !entry
                    .file_name()
                    .to_string_lossy()
                    .starts_with(".audit-s1-f1-")),
            "S1 staging directory should be removed after publishing the copy"
        );
        fs::remove_dir_all(directory).expect("remove formula-copy test directory");
    }

    #[test]
    fn umya_formula_copy_candidate_copies_template_formula_and_preserves_source() {
        let directory = temporary_directory();
        let target = directory.join("target.xlsx");
        let template = directory.join("template.xlsx");
        let output = directory.join("candidate.xlsx");
        write_target_with_sheet(&target, "检查表");
        write_template(&template);
        let source_bytes = fs::read(&target).expect("read original source workbook");
        let template_formula = read_workbook(&template)
            .expect("read formula template")
            .sheets[0]
            .formulas["A1"]
            .clone();
        let prepared = build_s1_formula_copy_edit_plan(&target, &template, &output, None, None)
            .expect("build exact S1 formula-copy batch");
        let capability = BackendCapability {
            engine: EngineId::UmyaSpreadsheet,
            status: CapabilityStatus::Unverified,
            format: prepared.edit_plan.format,
            operations: prepared
                .edit_plan
                .operations
                .iter()
                .map(EditOperation::kind)
                .collect(),
            supported_features: prepared.edit_plan.workbook.features_present.clone(),
            preserved: prepared.edit_plan.preserve.clone(),
            platform: std::env::consts::OS.to_owned(),
            evidence: None,
        };
        let backend = UmyaXlsxBackend::new(capability).expect("construct candidate backend");
        backend
            .apply(&prepared.edit_plan)
            .unwrap_or_else(|error| panic!("formula-copy candidate failed: {}", error.message));
        validate_s1_formula_copy_candidate(&prepared, &output)
            .expect("validate copied formula and literal inputs");

        let candidate = read_workbook(&output).expect("read formula-copy candidate");
        assert_eq!(
            candidate.sheets[0].formulas.get("A1"),
            Some(&template_formula)
        );
        assert_eq!(
            fs::read(&target).expect("reread original source"),
            source_bytes
        );
        assert!(output.is_file());
        fs::remove_dir_all(directory).expect("remove formula-copy candidate directory");
    }

    #[test]
    fn ooxml_s1_copy_replaces_comments_inside_validation_range_and_preserves_outside_comments() {
        let directory = temporary_directory();
        let target = directory.join("target.xlsx");
        let template = directory.join("template.xlsx");
        let output = directory.join("candidate.xlsx");
        write_target(&target);
        write_template(&template);
        add_comment(&target, "检查表", "A1", "Source User", "old source note");
        add_comment(
            &target,
            "检查表",
            "C1",
            "Source User",
            "outside source note",
        );
        add_comment(
            &template,
            "检查表",
            "A1",
            "Template Author",
            "validation note",
        );
        add_comment(
            &template,
            "检查表",
            "C1",
            "Template Author",
            "outside template note must not copy",
        );
        let target_before = fs::read(&target).expect("read source workbook before copy");
        let prepared = build_s1_formula_copy_edit_plan(&target, &template, &output, None, None)
            .expect("build S1 copy plan with comments");
        let capability = BackendCapability {
            engine: EngineId::OoxmlPatch,
            status: CapabilityStatus::Unverified,
            format: prepared.edit_plan.format,
            operations: [
                EditOperationKind::CreateSheet,
                EditOperationKind::DeleteSheet,
                EditOperationKind::DeleteColumn,
                EditOperationKind::ImportRange,
                EditOperationKind::ImportRangeWithFormulaTranslation,
                EditOperationKind::SetWorkbookCalculationProperties,
            ]
            .into_iter()
            .collect(),
            supported_features: prepared.edit_plan.workbook.features_present.clone(),
            preserved: prepared.edit_plan.preserve.clone(),
            platform: std::env::consts::OS.to_owned(),
            evidence: None,
        };
        OoxmlImportBackend::new(capability)
            .expect("construct OOXML candidate backend")
            .apply(&prepared.edit_plan)
            .unwrap_or_else(|error| panic!("S1 OOXML copy failed: {}", error.message));
        validate_s1_formula_copy_candidate(&prepared, &output)
            .expect("validate formula, values, comments, authors and positions");

        let copied = excel_reader::read_ooxml_comments(&output).expect("read output comments");
        let comments = &copied["检查表"];
        assert_eq!(
            comments["A1"],
            excel_reader::WorkbookComment {
                text: "validation note".to_owned(),
                author: Some("Template Author".to_owned()),
            }
        );
        assert_eq!(
            comments["C1"],
            excel_reader::WorkbookComment {
                text: "outside source note".to_owned(),
                author: Some("Source User".to_owned()),
            }
        );
        assert!(!comments.contains_key("B1"));
        assert_eq!(
            fs::read(&target).expect("reread original source workbook"),
            target_before
        );
        fs::remove_dir_all(directory).expect("remove S1 comment-copy test directory");
    }

    #[test]
    #[ignore = "requires RUST_WORKSPACE_OPENPYXL_PYTHON pointing to an openpyxl test runtime"]
    fn openpyxl_s1_candidate_copies_validation_comments_and_passes_output_validation() {
        let python = std::env::var_os("RUST_WORKSPACE_OPENPYXL_PYTHON")
            .map(PathBuf::from)
            .expect("set RUST_WORKSPACE_OPENPYXL_PYTHON to an openpyxl test runtime");
        let directory = temporary_directory();
        let target = directory.join("target.xlsx");
        let template = directory.join("template.xlsx");
        let output = directory.join("candidate.xlsx");
        write_target(&target);
        write_template(&template);
        add_comment(&target, "检查表", "A1", "Source User", "old source note");
        add_comment(
            &target,
            "检查表",
            "C1",
            "Source User",
            "outside source note",
        );
        add_comment(
            &template,
            "检查表",
            "A1",
            "Template Author",
            "validation note",
        );
        let target_before = fs::read(&target).expect("read source workbook before copy");
        let prepared = build_s1_formula_copy_edit_plan(&target, &template, &output, None, None)
            .expect("build S1 openpyxl candidate plan");
        let capability = BackendCapability {
            engine: EngineId::OpenpyxlWorker,
            status: CapabilityStatus::Unverified,
            format: prepared.edit_plan.format,
            operations: prepared
                .edit_plan
                .operations
                .iter()
                .map(EditOperation::kind)
                .collect(),
            supported_features: crate::excel_route::production_supported_features(),
            preserved: prepared.edit_plan.preserve.clone(),
            platform: std::env::consts::OS.to_owned(),
            evidence: None,
        };
        let worker = Path::new(env!("CARGO_MANIFEST_DIR"))
            .parent()
            .expect("audit-core workspace parent")
            .join("crates/excel-editor/python/openpyxl_worker.py");
        let backend = OpenpyxlXlsxBackend::new(
            capability,
            OpenpyxlWorkerConfig::new(python, worker).expect("construct openpyxl worker config"),
        )
        .expect("construct openpyxl candidate backend");
        backend
            .apply(&prepared.edit_plan)
            .unwrap_or_else(|error| panic!("openpyxl S1 candidate failed: {}", error.message));
        OpenpyxlXlsxValidator
            .validate(&prepared.edit_plan, &output)
            .expect("validate openpyxl candidate preservation");
        validate_s1_formula_copy_candidate(&prepared, &output)
            .expect("validate S1 formulas, values and copied comments");

        let comments = excel_reader::read_ooxml_comments(&output).expect("read candidate comments");
        assert_eq!(comments["检查表"]["A1"].text, "validation note");
        assert_eq!(
            comments["检查表"]["A1"].author.as_deref(),
            Some("Template Author")
        );
        assert_eq!(comments["检查表"]["C1"].text, "outside source note");
        assert_eq!(
            fs::read(&target).expect("reread source workbook"),
            target_before
        );
        fs::remove_dir_all(directory).expect("remove openpyxl candidate test directory");
    }

    #[test]
    #[ignore = "requires local S1 source/template paths; optional openpyxl test runtime"]
    fn real_s1_formula_copy_candidate_is_validated_before_recalculation() {
        let source = std::env::var_os("AUDIT_S1_COPY_REAL_SOURCE")
            .map(PathBuf::from)
            .expect("set AUDIT_S1_COPY_REAL_SOURCE");
        let template = std::env::var_os("AUDIT_S1_COPY_REAL_TEMPLATE")
            .map(PathBuf::from)
            .expect("set AUDIT_S1_COPY_REAL_TEMPLATE");
        let external = std::env::var_os("AUDIT_S1_COPY_REAL_EXTERNAL").map(PathBuf::from);
        let engine = std::env::var("AUDIT_S1_COPY_REAL_ENGINE").unwrap_or_else(|_| "ooxml".into());
        assert!(
            source.is_file(),
            "missing real S1 source: {}",
            source.display()
        );
        assert!(
            template.is_file(),
            "missing real S1 template: {}",
            template.display()
        );
        if let Some(external) = &external {
            assert!(
                external.is_file(),
                "missing external workbook: {}",
                external.display()
            );
        }

        let source_before = fs::read(&source).expect("read source before candidate copy");
        let template_before = fs::read(&template).expect("read template before candidate copy");
        let external_before = external
            .as_ref()
            .map(|path| fs::read(path).expect("read external workbook before candidate copy"));
        let directory = temporary_directory();
        let output = directory.join(format!("{engine}-candidate.xlsx"));
        let editor_input = if engine == "openpyxl" {
            let cacheless_source = directory.join("formula-source.xlsx");
            prepare_s1_recalculation_source_copy(&source, &cacheless_source)
                .expect("prepare isolated cacheless editor input");
            cacheless_source
        } else {
            source.clone()
        };
        let prepared = build_s1_formula_copy_edit_plan(
            &editor_input,
            &template,
            &output,
            None,
            external.as_deref(),
        )
        .expect("build real S1 formula-copy plan");

        match engine.as_str() {
            "ooxml" => {
                OoxmlImportBackend::new(crate::excel_route::ooxml_import_capability(
                    &std::env::consts::OS,
                ))
                .expect("construct OOXML candidate backend")
                .apply(&prepared.edit_plan)
                .unwrap_or_else(|error| panic!("real OOXML candidate failed: {}", error.message));
            }
            "openpyxl" => {
                let python = std::env::var_os("RUST_WORKSPACE_OPENPYXL_PYTHON")
                    .map(PathBuf::from)
                    .expect("set RUST_WORKSPACE_OPENPYXL_PYTHON for openpyxl candidate");
                let worker = Path::new(env!("CARGO_MANIFEST_DIR"))
                    .parent()
                    .expect("audit-core workspace parent")
                    .join("crates/excel-editor/python/openpyxl_worker.py");
                let capability = BackendCapability {
                    engine: EngineId::OpenpyxlWorker,
                    status: CapabilityStatus::Unverified,
                    format: prepared.edit_plan.format,
                    operations: prepared
                        .edit_plan
                        .operations
                        .iter()
                        .map(EditOperation::kind)
                        .collect(),
                    supported_features: crate::excel_route::production_supported_features(),
                    preserved: prepared.edit_plan.preserve.clone(),
                    platform: std::env::consts::OS.to_owned(),
                    evidence: None,
                };
                OpenpyxlXlsxBackend::new(
                    capability,
                    OpenpyxlWorkerConfig::new(python, worker)
                        .expect("construct openpyxl candidate runtime"),
                )
                .expect("construct openpyxl candidate backend")
                .apply(&prepared.edit_plan)
                .unwrap_or_else(|error| {
                    panic!("real openpyxl candidate failed: {}", error.message)
                });
            }
            other => panic!("unsupported test candidate engine {other:?}"),
        }

        OpenpyxlXlsxValidator
            .validate(&prepared.edit_plan, &output)
            .expect("validate real formula-copy candidate preservation");
        validate_s1_formula_copy_candidate(&prepared, &output)
            .expect("validate real template formulas, values and copied comments");
        let dump = read_workbook(&output).expect("read real candidate output");
        let copied_comments = excel_reader::read_ooxml_comments(&output)
            .expect("read comments from real candidate output")
            .values()
            .map(BTreeMap::len)
            .sum::<usize>();
        println!(
            "real_s1_candidate engine={engine} operations={} formulas={} copied_comments={} output={}",
            prepared.edit_plan.operations.len(),
            dump.sheets
                .iter()
                .map(|sheet| sheet.formulas.len())
                .sum::<usize>(),
            copied_comments,
            output.display()
        );
        assert_eq!(
            fs::read(&source).expect("reread real source"),
            source_before
        );
        assert_eq!(
            fs::read(&template).expect("reread real template"),
            template_before
        );
        if let (Some(external), Some(external_before)) = (&external, external_before) {
            assert_eq!(
                fs::read(external).expect("reread external workbook"),
                external_before
            );
        }
        fs::remove_dir_all(directory).expect("remove real candidate output directory");
    }

    #[test]
    fn external_formula_sheets_and_their_dependencies_are_copied_without_rejecting_formulas() {
        let directory = temporary_directory();
        let target = directory.join("target.xlsx");
        let template = directory.join("template.xlsx");
        let external = directory.join("external.xlsx");
        let output = directory.join("candidate.xlsx");
        write_target_with_sheet(&target, "检查表");
        write_template_with_formula(&template, "='参照表'!C4");
        let mut book = Workbook::new();
        book.add_worksheet()
            .set_name("参照表")
            .unwrap()
            .write_formula(3, 2, "='支持表'!B1*2")
            .unwrap();
        book.add_worksheet()
            .set_name("支持表")
            .unwrap()
            .write_number(0, 1, 21.0)
            .unwrap();
        book.save(&external).unwrap();
        let target_before = fs::read(&target).unwrap();
        let external_before = fs::read(&external).unwrap();
        let prepared =
            build_s1_formula_copy_edit_plan(&target, &template, &output, None, Some(&external))
                .unwrap();
        let external_plan = prepared.external_plan.as_ref().unwrap();
        assert_eq!(external_plan.selection_reason, "formula_dependencies");
        assert_eq!(
            external_plan
                .sheets
                .iter()
                .map(|sheet| sheet.sheet_name.as_str())
                .collect::<Vec<_>>(),
            vec!["参照表", "支持表"]
        );
        let capability = BackendCapability {
            engine: EngineId::UmyaSpreadsheet,
            status: CapabilityStatus::Unverified,
            format: prepared.edit_plan.format,
            operations: prepared
                .edit_plan
                .operations
                .iter()
                .map(EditOperation::kind)
                .collect(),
            supported_features: prepared.edit_plan.workbook.features_present.clone(),
            preserved: prepared.edit_plan.preserve.clone(),
            platform: std::env::consts::OS.to_owned(),
            evidence: None,
        };
        UmyaXlsxBackend::new(capability)
            .unwrap()
            .apply(&prepared.edit_plan)
            .unwrap();
        validate_s1_formula_copy_candidate(&prepared, &output)
            .expect("validate imported external formulas and values");
        let copied = read_workbook(&output).unwrap();
        let lookup = copied
            .sheets
            .iter()
            .find(|sheet| sheet.name == "参照表")
            .unwrap();
        assert_eq!(
            lookup.formulas["C4"].trim_start_matches('='),
            "'支持表'!B1*2"
        );
        let supporting = copied
            .sheets
            .iter()
            .find(|sheet| sheet.name == "支持表")
            .unwrap();
        assert_eq!(supporting.values["B1"]["value"].as_f64(), Some(21.0));
        assert_eq!(fs::read(&target).unwrap(), target_before);
        assert_eq!(fs::read(&external).unwrap(), external_before);
        fs::remove_dir_all(directory).unwrap();
    }

    #[test]
    fn umya_copies_210_vlookup_formulas_with_both_referenced_external_sheets() {
        let directory = temporary_directory();
        let target = directory.join("target.xlsx");
        let template = directory.join("template.xlsx");
        let external = directory.join("external.xlsx");
        let output = directory.join("candidate.xlsx");
        write_target_with_sheet(&target, "检查表");
        write_template_with_formula(&template, "='参照表'!D2");

        let mut external_book = Workbook::new();
        let reference = external_book
            .add_worksheet()
            .set_name("参照表")
            .expect("name reference sheet");
        reference.write_string(0, 0, "机构代码").unwrap();
        reference.write_string(0, 3, "机构名称").unwrap();
        for index in 0..210_u32 {
            let row = index + 1;
            let row_number = row + 1;
            reference
                .write_string(row, 0, &format!("ORG-{index:03}"))
                .unwrap();
            let formula = format!("VLOOKUP(A{row_number},'机构参照表'!$A:$B,2,FALSE)");
            reference.write_formula(row, 3, formula.as_str()).unwrap();
        }
        let institution = external_book
            .add_worksheet()
            .set_name("机构参照表")
            .expect("name institution reference sheet");
        institution.write_string(0, 0, "机构代码").unwrap();
        institution.write_string(0, 1, "机构名称").unwrap();
        institution.write_string(1, 0, "ORG-000").unwrap();
        institution.write_string(1, 1, "示例机构").unwrap();
        external_book.save(&external).unwrap();

        let target_before = fs::read(&target).unwrap();
        let external_before = fs::read(&external).unwrap();
        let prepared =
            build_s1_formula_copy_edit_plan(&target, &template, &output, None, Some(&external))
                .expect("build reference-table import plan");
        let planned_sheets = prepared
            .external_plan
            .as_ref()
            .expect("external plan")
            .sheets
            .iter()
            .map(|sheet| sheet.sheet_name.as_str())
            .collect::<Vec<_>>();
        assert_eq!(planned_sheets, vec!["参照表", "机构参照表"]);

        let capability = BackendCapability {
            engine: EngineId::UmyaSpreadsheet,
            status: CapabilityStatus::Unverified,
            format: prepared.edit_plan.format,
            operations: prepared
                .edit_plan
                .operations
                .iter()
                .map(EditOperation::kind)
                .collect(),
            supported_features: prepared.edit_plan.workbook.features_present.clone(),
            preserved: prepared.edit_plan.preserve.clone(),
            platform: std::env::consts::OS.to_owned(),
            evidence: None,
        };
        UmyaXlsxBackend::new(capability)
            .unwrap()
            .apply(&prepared.edit_plan)
            .unwrap();
        validate_s1_formula_copy_candidate(&prepared, &output)
            .expect("validate all copied VLOOKUP formulas and reference values");

        let corrupted_values = directory.join("corrupted-reference-value.xlsx");
        fs::copy(&output, &corrupted_values)
            .expect("copy candidate for imported-value integrity probe");
        let mut corrupted_workbook = excel_editor::open_xlsx(&corrupted_values)
            .expect("open candidate for controlled reference-value change");
        corrupted_workbook
            .sheet_by_name_mut("机构参照表")
            .expect("find copied institution lookup sheet")
            .cell_mut("B2")
            .set_value("wrong institution");
        excel_editor::save_xlsx(&corrupted_workbook, &corrupted_values)
            .expect("save controlled reference-value change");
        let error = validate_s1_formula_copy_candidate(&prepared, &corrupted_values)
            .expect_err("changed lookup values must stop the S1 copy path");
        assert!(
            error.to_string().contains("导入区域文本或数值与预期不一致"),
            "unexpected imported-value validation error: {error:#}"
        );

        let reopened = read_workbook(&output).expect("reopen Umya output through Calamine");
        let references = reopened
            .sheets
            .iter()
            .find(|sheet| sheet.name == "参照表")
            .expect("copied reference worksheet");
        for index in 0..210_u32 {
            let row = index + 2;
            let address = format!("D{row}");
            assert_eq!(
                references.formulas.get(&address),
                Some(&format!("VLOOKUP(A{row},'机构参照表'!$A:$B,2,FALSE)")),
                "formula reference must remain unchanged at 参照表!{address}"
            );
        }
        let institution = reopened
            .sheets
            .iter()
            .find(|sheet| sheet.name == "机构参照表")
            .expect("copied institution lookup worksheet");
        assert_eq!(institution.values["B2"]["value"].as_str(), Some("示例机构"));
        assert_eq!(
            reopened
                .sheets
                .iter()
                .find(|sheet| sheet.name == "检查表")
                .unwrap()
                .formulas
                .get("A1")
                .map(String::as_str),
            Some("'参照表'!D2")
        );
        assert_eq!(fs::read(&target).unwrap(), target_before);
        assert_eq!(fs::read(&external).unwrap(), external_before);
        assert!(output.is_file());
        fs::remove_dir_all(directory).unwrap();
    }

    #[test]
    fn s1_copy_validation_checks_imported_formulas_but_ignores_other_formula_text() {
        let directory = temporary_directory();
        let target = directory.join("target.xlsx");
        let template = directory.join("template.xlsx");
        let output = directory.join("candidate.xlsx");
        write_target_with_formula_outside_copy_range(&target);
        write_template(&template);
        let prepared = build_s1_formula_copy_edit_plan(&target, &template, &output, None, None)
            .expect("build S1 formula-copy plan");
        let capability = BackendCapability {
            engine: EngineId::UmyaSpreadsheet,
            status: CapabilityStatus::Unverified,
            format: prepared.edit_plan.format,
            operations: prepared
                .edit_plan
                .operations
                .iter()
                .map(EditOperation::kind)
                .collect(),
            supported_features: prepared.edit_plan.workbook.features_present.clone(),
            preserved: prepared.edit_plan.preserve.clone(),
            platform: std::env::consts::OS.to_owned(),
            evidence: None,
        };
        UmyaXlsxBackend::new(capability)
            .expect("construct candidate backend")
            .apply(&prepared.edit_plan)
            .expect("copy formulas to the candidate");
        validate_s1_formula_copy_candidate(&prepared, &output)
            .expect("valid candidate should pass formula-copy checks");

        let corrupted_formula = directory.join("corrupted-formula.xlsx");
        fs::copy(&output, &corrupted_formula).expect("copy candidate for formula-loss probe");
        let mut workbook = excel_editor::open_xlsx(&corrupted_formula)
            .expect("open formula candidate for controlled corruption");
        workbook
            .sheet_by_name_mut("检查表")
            .expect("find formula target sheet")
            .cell_mut("A1")
            .set_value("formula lost");
        excel_editor::save_xlsx(&workbook, &corrupted_formula)
            .expect("save controlled formula corruption");
        let error = validate_s1_formula_copy_candidate(&prepared, &corrupted_formula)
            .expect_err("formula loss must stop the S1 copy path");
        assert!(error
            .to_string()
            .contains("编辑计划导入区域公式与预期不一致"));

        let other_formula_changed = directory.join("other-formula-changed.xlsx");
        fs::copy(&output, &other_formula_changed)
            .expect("copy candidate for formula-outside-validation-area probe");
        let mut workbook = excel_editor::open_xlsx(&other_formula_changed)
            .expect("open candidate for formula change outside validation area");
        workbook
            .sheet_by_name_mut("检查表")
            .expect("find target data sheet")
            .cell_mut("C1")
            .set_formula("SUM(A1:B1)+1");
        excel_editor::save_xlsx(&workbook, &other_formula_changed)
            .expect("save formula change outside validation area");
        validate_s1_formula_copy_candidate(&prepared, &other_formula_changed)
            .expect("formula text outside the named validation range is outside this copy check");

        let corrupted_structure = directory.join("corrupted-structure.xlsx");
        fs::copy(&output, &corrupted_structure)
            .expect("copy candidate for worksheet-structure probe");
        let mut workbook = excel_editor::open_xlsx(&corrupted_structure)
            .expect("open candidate for controlled worksheet change");
        workbook
            .new_sheet("Unexpected")
            .expect("add unexpected worksheet");
        excel_editor::save_xlsx(&workbook, &corrupted_structure)
            .expect("save controlled worksheet change");
        let error = validate_s1_formula_copy_candidate(&prepared, &corrupted_structure)
            .expect_err("unexpected worksheet changes must stop the S1 copy path");
        assert!(error.to_string().contains("工作表名称或顺序与计划不一致"));

        fs::remove_dir_all(directory).unwrap();
    }

    #[test]
    fn s1_import_value_comparison_compares_excel_numeric_serials_across_formatting() {
        let empty = serde_json::json!({"type":"string", "value":""});
        let nonempty = serde_json::json!({"type":"string", "value":"x"});
        let date_serial = serde_json::json!({"type":"datetime_serial", "value":46295.0});
        let number = serde_json::json!({"type":"number", "value":46295.0});
        let integer = serde_json::json!({"type":"integer", "value":46295});
        let different_number = serde_json::json!({"type":"number", "value":46296.0});
        assert!(equivalent_import_value(Some(&empty), None));
        assert!(equivalent_import_value(None, Some(&empty)));
        assert!(equivalent_import_value(Some(&date_serial), Some(&number)));
        assert!(equivalent_import_value(Some(&date_serial), Some(&integer)));
        assert!(!equivalent_import_value(Some(&empty), Some(&nonempty)));
        assert!(!equivalent_import_value(Some(&nonempty), None));
        assert!(!equivalent_import_value(
            Some(&date_serial),
            Some(&different_number)
        ));
    }

    #[test]
    fn refuses_formula_copy_when_target_lacks_a_mapped_template_sheet() {
        let directory = temporary_directory();
        let target = directory.join("target.xlsx");
        let template = directory.join("template.xlsx");
        let output = directory.join("output.xlsx");
        write_target_with_sheet(&target, "其他工作表");
        write_template(&template);

        let error = build_s1_formula_copy_edit_plan(&target, &template, &output, None, None)
            .expect_err("missing target sheet must stop plan construction");

        assert!(
            format!("{error:#}").contains("目标工作簿缺少模板映射工作表"),
            "{error:#}"
        );
        assert!(
            !output.exists(),
            "preflight must not create candidate output"
        );
        fs::remove_dir_all(directory).expect("remove formula-copy test directory");
    }

    #[test]
    fn deduplicates_identical_structured_formula_copy_ranges() {
        let directory = temporary_directory();
        let target = directory.join("target.xlsx");
        let template = directory.join("template.xlsx");
        let output = directory.join("output.xlsx");
        write_target(&target);
        write_structured_template_with_duplicate_ranges(&template);

        let prepared = build_s1_formula_copy_edit_plan(&target, &template, &output, None, None)
            .expect("build plan with duplicate structured ranges");
        let import_range_count = prepared
            .edit_plan
            .operations
            .iter()
            .filter(|operation| operation.kind() == EditOperationKind::ImportRange)
            .count();

        assert_eq!(prepared.template_summary.copy_ranges, 2);
        assert_eq!(import_range_count, 2);
        assert_eq!(prepared.edit_plan.operations.len(), 4);
        assert!(matches!(
            &prepared.edit_plan.operations[1],
            EditOperation::CreateSheet { name } if name == "审核规则"
        ));
        assert!(matches!(
            &prepared.edit_plan.operations[2],
            EditOperation::ImportRange {
                source_sheet,
                source_range,
                target_sheet,
                copy_values: true,
                copy_formulas: true,
                copy_comments: true,
                ..
            } if source_sheet == "审核规则"
                && source_range == "A1:H3"
                && target_sheet == "审核规则"
        ));
        assert!(!output.exists(), "plan construction must not write output");
        fs::remove_dir_all(directory).expect("remove formula-copy test directory");
    }

    #[test]
    fn umya_copies_structured_rule_sheet_comments_into_the_audit_copy() {
        let directory = temporary_directory();
        let target = directory.join("target.xlsx");
        let template = directory.join("template.xlsx");
        let output = directory.join("candidate.xlsx");
        write_target(&target);
        write_structured_template_with_duplicate_ranges(&template);
        let target_before = fs::read(&target).expect("read original source workbook");
        let prepared = build_s1_formula_copy_edit_plan(&target, &template, &output, None, None)
            .expect("build structured rule sheet copy plan");
        let capability = BackendCapability {
            engine: EngineId::UmyaSpreadsheet,
            status: CapabilityStatus::Unverified,
            format: prepared.edit_plan.format,
            operations: prepared
                .edit_plan
                .operations
                .iter()
                .map(EditOperation::kind)
                .collect(),
            supported_features: prepared.edit_plan.workbook.features_present.clone(),
            preserved: prepared.edit_plan.preserve.clone(),
            platform: std::env::consts::OS.to_owned(),
            evidence: None,
        };
        UmyaXlsxBackend::new(capability)
            .expect("construct Umya candidate backend")
            .apply(&prepared.edit_plan)
            .unwrap_or_else(|error| panic!("structured rule sheet copy failed: {}", error.message));
        validate_s1_formula_copy_candidate(&prepared, &output)
            .expect("validate structured rule sheet values and formulas");

        let candidate = read_workbook(&output).expect("read structured rule sheet candidate");
        let rule_sheet = candidate
            .sheets
            .iter()
            .find(|sheet| sheet.name == "审核规则")
            .expect("copied structured rule sheet");
        assert_eq!(rule_sheet.comments["A1"], "规则表说明");
        let candidate_umya = umya_spreadsheet::reader::xlsx::read(&output)
            .expect("read copied rule sheet comment author");
        let rule_sheet_umya = candidate_umya
            .get_sheet_by_name("审核规则")
            .expect("copied structured rule sheet in Umya");
        let comment = rule_sheet_umya
            .comments()
            .iter()
            .find(|comment| comment.coordinate().to_string() == "A1")
            .expect("copied rule sheet comment");
        assert_eq!(comment.author(), "测试");
        assert_eq!(
            fs::read(&target).expect("reread original source"),
            target_before
        );
        fs::remove_dir_all(directory).expect("remove structured rule sheet test directory");
    }

    #[test]
    fn replaces_an_existing_audit_rule_sheet_with_the_structured_template_sheet() {
        let directory = temporary_directory();
        let target = directory.join("target.xlsx");
        let template = directory.join("template.xlsx");
        let output = directory.join("output.xlsx");
        write_target_with_rule_sheet(&target);
        write_structured_template_with_duplicate_ranges(&template);

        let prepared = build_s1_formula_copy_edit_plan(&target, &template, &output, None, None)
            .expect("build plan replacing existing rule sheet");

        assert!(matches!(
            &prepared.edit_plan.operations[1],
            EditOperation::DeleteSheet { name } if name == "审核规则"
        ));
        assert!(matches!(
            &prepared.edit_plan.operations[2],
            EditOperation::CreateSheet { name } if name == "审核规则"
        ));
        assert!(prepared.edit_plan.drop_allowed.contains_key("审核规则"));
        assert_eq!(
            prepared
                .edit_plan
                .operations
                .iter()
                .filter(|operation| operation.kind() == EditOperationKind::DeleteSheet)
                .count(),
            1
        );
        assert!(!output.exists(), "plan construction must not write output");
        fs::remove_dir_all(directory).expect("remove formula-copy test directory");
    }

    #[test]
    fn refuses_to_overwrite_an_existing_formula_copy_output() {
        let directory = temporary_directory();
        let target = directory.join("target.xlsx");
        let template = directory.join("template.xlsx");
        let output = directory.join("output.xlsx");
        write_target(&target);
        write_template(&template);
        fs::write(&output, b"keep existing output").expect("create existing output sentinel");

        let error = build_s1_formula_copy_edit_plan(&target, &template, &output, None, None)
            .expect_err("existing output must not be replaced");

        assert!(format!("{error:#}").contains("拒绝覆盖"));
        assert_eq!(
            fs::read(&output).expect("read existing output sentinel"),
            b"keep existing output"
        );
        fs::remove_dir_all(directory).expect("remove formula-copy test directory");
    }

    #[test]
    fn plans_only_statically_referenced_external_sheets_before_template_ranges() {
        let directory = temporary_directory();
        let target = directory.join("target.xlsx");
        let template = directory.join("template.xlsx");
        let external = directory.join("external.xlsx");
        let output = directory.join("output.xlsx");
        write_target(&target);
        let mut template_workbook = Workbook::new();
        let template_sheet = template_workbook.add_worksheet();
        template_sheet
            .set_name("检查表")
            .expect("name formula sheet");
        template_sheet
            .write_formula(0, 0, "=IFERROR(VLOOKUP(A1,'外部维表'!$A:$B,2,FALSE),\"\")")
            .expect("write in-range audit formula");
        template_sheet
            .write_formula(0, 2, "='未引用表'!A1")
            .expect("write unrelated formula outside audit range");
        template_workbook
            .define_name("校验区域", "='检查表'!$A$1:$B$2")
            .expect("define audit validation range");
        template_workbook
            .save(&template)
            .expect("save formula template");
        write_external_workbook(&external);

        let prepared =
            build_s1_formula_copy_edit_plan(&target, &template, &output, None, Some(&external))
                .expect("build S1 plan with external workbook");

        let external_plan = prepared.external_plan.expect("external sheet plan");
        assert_eq!(external_plan.selection_reason, "formula_references");
        assert_eq!(external_plan.sheets.len(), 1);
        assert_eq!(external_plan.sheets[0].sheet_name, "外部维表");
        assert_eq!(external_plan.sheets[0].source_range, "A1:B2");
        assert!(prepared.external_profile.is_some());
        assert!(matches!(
            &prepared.edit_plan.operations[0],
            EditOperation::CreateSheet { name } if name == "外部维表"
        ));
        assert!(matches!(
            &prepared.edit_plan.operations[1],
            EditOperation::ImportRange {
                source_sheet,
                source_range,
                target_sheet,
                copy_values: true,
                copy_formulas: true,
                style_mode: RangeStyleMode::None,
                materialize_blank_cells: true,
                ..
            } if source_sheet == "外部维表"
                && source_range == "A1:B2"
                && target_sheet == "外部维表"
        ));
        assert!(!output.exists(), "plan construction must not write output");
        fs::remove_dir_all(directory).expect("remove formula-copy test directory");
    }

    #[test]
    fn falls_back_to_all_external_sheets_when_template_references_are_dynamic() {
        let external = vec![
            SheetDump {
                name: "表一".to_owned(),
                values: BTreeMap::from([(
                    "A1".to_owned(),
                    serde_json::json!({"type":"string", "value":"值"}),
                )]),
                formulas: BTreeMap::new(),
                formula_status: "not_requested".to_owned(),
                merges: Vec::new(),
                merge_status: "not_requested".to_owned(),
                comments: BTreeMap::new(),
                comments_status: "not_requested".to_owned(),
                conditional_formatting: Vec::new(),
                conditional_formatting_status: "not_requested".to_owned(),
                data_validations: Vec::new(),
                data_validations_status: "not_requested".to_owned(),
            },
            SheetDump {
                name: "表二".to_owned(),
                values: BTreeMap::from([(
                    "B3".to_owned(),
                    serde_json::json!({"type":"string", "value":"值"}),
                )]),
                formulas: BTreeMap::new(),
                formula_status: "not_requested".to_owned(),
                merges: Vec::new(),
                merge_status: "not_requested".to_owned(),
                comments: BTreeMap::new(),
                comments_status: "not_requested".to_owned(),
                conditional_formatting: Vec::new(),
                conditional_formatting_status: "not_requested".to_owned(),
                data_validations: Vec::new(),
                data_validations_status: "not_requested".to_owned(),
            },
        ];

        let (selected, reason) = select_external_sheets(["=INDIRECT(A1)"].into_iter(), &external)
            .expect("select external sheets for dynamic reference");

        assert_eq!(reason, "all_sheets_fallback");
        assert_eq!(selected, vec!["表一", "表二"]);
        assert_eq!(used_range(&external[1]).expect("used range"), "B3");
    }

    #[test]
    fn refuses_external_sheet_name_conflicting_with_source_workbook() {
        let directory = temporary_directory();
        let target = directory.join("target.xlsx");
        let template = directory.join("template.xlsx");
        let external = directory.join("external.xlsx");
        let output = directory.join("output.xlsx");
        write_target(&target);
        write_template_with_formula(&template, "='检查表'!A1");
        write_external_workbook(&external);
        let mut conflict_workbook = Workbook::new();
        let conflict = conflict_workbook.add_worksheet();
        conflict
            .set_name("检查表")
            .expect("name conflicting external sheet");
        conflict
            .write_string(0, 0, "conflict")
            .expect("write conflicting external value");
        conflict_workbook
            .save(&external)
            .expect("save conflicting external workbook");

        // The template's static lookup resolves to a same-named source sheet;
        // the importer must fail closed before generating an edit plan.
        let error =
            build_s1_formula_copy_edit_plan(&target, &template, &output, None, Some(&external))
                .expect_err("external name conflict must reject the plan");

        assert!(format!("{error:#}").contains("与源报送工作簿重名"));
        assert!(!output.exists(), "rejected plan must not write output");
        fs::remove_dir_all(directory).expect("remove formula-copy test directory");
    }
}
