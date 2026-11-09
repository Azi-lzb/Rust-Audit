//! Persistent S1-F01 manual-recalculation recovery tasks.
//!
//! The store contains only workflow state and extracted audit issues. Audit
//! copies remain in the run's output directory and are never rewritten by the
//! recovery store.

use crate::models::Issue;
use anyhow::{bail, ensure, Context, Result};
use serde::{Deserialize, Serialize};
use serde_json::Value;
use sha2::{Digest, Sha256};
use std::{
    env, fs,
    io::{Read, Write},
    path::{Path, PathBuf},
    time::{SystemTime, UNIX_EPOCH},
};

const RECOVERY_SCHEMA_VERSION: u32 = 1;

fn default_s1_conditional_format_source() -> String {
    "source".to_owned()
}

#[derive(Clone, Copy, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "SCREAMING_SNAKE_CASE")]
pub enum S1RecoveryStatus {
    Copying,
    WaitingManualRecalc,
    ReadyToResume,
    ResumeFailed,
    Completed,
    Partial,
    Failed,
}

impl S1RecoveryStatus {
    pub fn is_pending(self) -> bool {
        matches!(
            self,
            Self::Copying | Self::WaitingManualRecalc | Self::ReadyToResume | Self::ResumeFailed
        )
    }

    pub fn as_str(self) -> &'static str {
        match self {
            Self::Copying => "COPYING",
            Self::WaitingManualRecalc => "WAITING_MANUAL_RECALC",
            Self::ReadyToResume => "READY_TO_RESUME",
            Self::ResumeFailed => "RESUME_FAILED",
            Self::Completed => "COMPLETED",
            Self::Partial => "PARTIAL",
            Self::Failed => "FAILED",
        }
    }
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct S1RecoveryManifest {
    pub schema_version: u32,
    pub run_id: String,
    pub feature_id: String,
    pub created_at: String,
    pub updated_at: String,
    pub output_directory: PathBuf,
    pub run_directory: PathBuf,
    pub result_path: Option<PathBuf>,
    /// Deterministic S1-F3 result path for crash-safe publication. The result
    /// is staged and hashed before it is moved to this path.
    #[serde(default)]
    pub recovery_result_path: Option<PathBuf>,
    /// Candidate workbook and expected digest persisted before publication so
    /// a restarted worker can finish the same output without creating a second.
    #[serde(default)]
    pub recovery_result_candidate_path: Option<PathBuf>,
    #[serde(default)]
    pub recovery_result_input_sha256: Option<String>,
    #[serde(default)]
    pub recovery_result_sha256: Option<String>,
    pub external_data: Option<PathBuf>,
    pub period: String,
    pub batch_id: String,
    pub audit_time: String,
    pub mapping_config: Option<PathBuf>,
    pub mapping_config_sha256: Option<String>,
    /// The S1-F01 conditional-format value source must remain stable when a
    /// pending manual-recalculation task resumes through S1-F3.
    #[serde(default = "default_s1_conditional_format_source")]
    pub conditional_format_source: String,
    pub history_issues: Vec<Issue>,
    pub skipped_files: Vec<Value>,
    pub files: Vec<S1RecoveryFile>,
}

impl S1RecoveryManifest {
    pub fn new(
        run_id: impl Into<String>,
        output_directory: PathBuf,
        run_directory: PathBuf,
        period: impl Into<String>,
        batch_id: impl Into<String>,
        audit_time: impl Into<String>,
        mapping_config: Option<PathBuf>,
        mapping_config_sha256: Option<String>,
        history_issues: Vec<Issue>,
        skipped_files: Vec<Value>,
        files: Vec<S1RecoveryFile>,
    ) -> Self {
        let now = timestamp();
        Self {
            schema_version: RECOVERY_SCHEMA_VERSION,
            run_id: run_id.into(),
            feature_id: "S1-F01".to_owned(),
            created_at: now.clone(),
            updated_at: now,
            output_directory,
            run_directory,
            result_path: None,
            recovery_result_path: None,
            recovery_result_candidate_path: None,
            recovery_result_input_sha256: None,
            recovery_result_sha256: None,
            external_data: None,
            period: period.into(),
            batch_id: batch_id.into(),
            audit_time: audit_time.into(),
            mapping_config,
            mapping_config_sha256,
            conditional_format_source: default_s1_conditional_format_source(),
            history_issues,
            skipped_files,
            files,
        }
    }
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct S1RecoveryFile {
    pub file_id: String,
    pub source_file: PathBuf,
    pub template_file: PathBuf,
    pub template_sha256: String,
    #[serde(default)]
    pub formula_fingerprint: Option<String>,
    /// Fingerprint canonicalized only for LibreOffice's `TRUE()`/`FALSE()`
    /// serialization. Kept beside the exact fingerprint for older task files.
    #[serde(default)]
    pub formula_fingerprint_normalized: Option<String>,
    /// Fingerprint of audit inputs and rule metadata before manual
    /// recalculation. Formula result cells are excluded because recalculation
    /// is expected to change them. Optional for older manifests.
    #[serde(default)]
    pub audit_copy_semantic_fingerprint: Option<String>,
    pub audit_copy: Option<PathBuf>,
    pub status: S1RecoveryStatus,
    pub reason: Option<String>,
    pub recalculation_engine: Option<String>,
    #[serde(default)]
    pub issues_extracted: bool,
    #[serde(default)]
    pub audit_partial: bool,
    pub formula_issues: Vec<Issue>,
    pub conditional_issues: Vec<Issue>,
    pub conditional_diagnostics: Vec<Value>,
    pub updated_at: String,
}

impl S1RecoveryFile {
    pub fn new(
        source_file: PathBuf,
        template_file: PathBuf,
        template_sha256: String,
        audit_copy: Option<PathBuf>,
        status: S1RecoveryStatus,
    ) -> Self {
        let now = timestamp();
        Self {
            file_id: stable_file_id(&source_file),
            source_file,
            template_file,
            template_sha256,
            formula_fingerprint: None,
            formula_fingerprint_normalized: None,
            audit_copy_semantic_fingerprint: None,
            audit_copy,
            status,
            reason: None,
            recalculation_engine: None,
            issues_extracted: false,
            audit_partial: false,
            formula_issues: Vec::new(),
            conditional_issues: Vec::new(),
            conditional_diagnostics: Vec::new(),
            updated_at: now,
        }
    }
}

#[derive(Clone, Debug, Serialize)]
pub struct S1RecoveryTaskSummary {
    pub run_id: String,
    pub file_id: String,
    pub source_file: PathBuf,
    pub audit_copy: Option<PathBuf>,
    pub status: &'static str,
    pub reason: Option<String>,
    pub updated_at: String,
}

#[derive(Clone, Debug)]
pub struct S1RecoveryStore {
    directory: PathBuf,
}

impl S1RecoveryStore {
    pub fn new(directory: impl Into<PathBuf>) -> Self {
        Self {
            directory: directory.into(),
        }
    }

    pub fn user_default() -> Result<Self> {
        if let Some(directory) = env::var_os("AUDIT_STATE_DIR").filter(|value| !value.is_empty()) {
            return Ok(Self::new(PathBuf::from(directory).join("s1-f3-recovery")));
        }
        #[cfg(windows)]
        let base = env::var_os("LOCALAPPDATA")
            .map(PathBuf::from)
            .context("Windows 用户状态目录不可用：LOCALAPPDATA 未设置")?;
        #[cfg(not(windows))]
        let base = if let Some(path) = env::var_os("XDG_STATE_HOME").filter(|v| !v.is_empty()) {
            PathBuf::from(path)
        } else {
            env::var_os("HOME")
                .map(PathBuf::from)
                .context("用户状态目录不可用：XDG_STATE_HOME 和 HOME 均未设置")?
                .join(".local")
                .join("state")
        };
        Ok(Self::new(
            base.join(crate::app_identity::PRODUCT_DIR)
                .join("s1-f3-recovery"),
        ))
    }

    pub fn directory(&self) -> &Path {
        &self.directory
    }

    pub fn save(&self, manifest: &mut S1RecoveryManifest) -> Result<PathBuf> {
        validate_run_id(&manifest.run_id)?;
        ensure!(
            manifest.schema_version == RECOVERY_SCHEMA_VERSION,
            "不支持的 S1 恢复清单版本：{}",
            manifest.schema_version
        );
        ensure!(manifest.feature_id == "S1-F01", "恢复清单功能编号不正确");
        fs::create_dir_all(&self.directory)
            .with_context(|| format!("创建 S1 恢复状态目录 {}", self.directory.display()))?;
        self.recover_interrupted_replacements()?;
        manifest.updated_at = timestamp();
        let path = self.manifest_path(&manifest.run_id)?;
        let bytes = serde_json::to_vec_pretty(manifest).context("序列化 S1 恢复清单")?;
        atomic_replace(&path, &bytes)?;
        Ok(path)
    }

    pub fn load(&self, run_id: &str) -> Result<S1RecoveryManifest> {
        validate_run_id(run_id)?;
        self.recover_interrupted_replacements()?;
        let path = self.manifest_path(run_id)?;
        let bytes =
            fs::read(&path).with_context(|| format!("读取 S1 恢复清单 {}", path.display()))?;
        let manifest: S1RecoveryManifest =
            serde_json::from_slice(&bytes).context("解析 S1 恢复清单")?;
        ensure!(
            manifest.run_id == run_id,
            "S1 恢复清单 run_id 与文件名不一致"
        );
        ensure!(manifest.feature_id == "S1-F01", "仅支持恢复 S1-F01 任务");
        ensure!(
            manifest.schema_version == RECOVERY_SCHEMA_VERSION,
            "不支持的 S1 恢复清单版本：{}",
            manifest.schema_version
        );
        Ok(manifest)
    }

    pub fn pending_tasks(&self) -> Result<Vec<S1RecoveryTaskSummary>> {
        if !self.directory.exists() {
            return Ok(Vec::new());
        }
        self.recover_interrupted_replacements()?;
        self.reconcile_interrupted_copies()?;
        let mut tasks = Vec::new();
        for entry in fs::read_dir(&self.directory)
            .with_context(|| format!("扫描 S1 恢复状态目录 {}", self.directory.display()))?
        {
            let entry = entry?;
            let path = entry.path();
            if !path.is_file() || path.extension().and_then(|ext| ext.to_str()) != Some("json") {
                continue;
            }
            let bytes = fs::read(&path)
                .with_context(|| format!("读取 S1 恢复任务清单 {}", path.display()))?;
            let manifest: S1RecoveryManifest = serde_json::from_slice(&bytes)
                .with_context(|| format!("解析 S1 恢复任务清单 {}", path.display()))?;
            ensure!(
                manifest.schema_version == RECOVERY_SCHEMA_VERSION
                    && manifest.feature_id == "S1-F01",
                "S1 恢复状态目录中存在不兼容清单：{}",
                path.display()
            );
            for file in manifest.files.iter().filter(|file| {
                file.status.is_pending() && file.audit_copy.as_deref().is_some_and(Path::is_file)
            }) {
                tasks.push(S1RecoveryTaskSummary {
                    run_id: manifest.run_id.clone(),
                    file_id: file.file_id.clone(),
                    source_file: file.source_file.clone(),
                    audit_copy: file.audit_copy.clone(),
                    status: file.status.as_str(),
                    reason: file.reason.clone(),
                    updated_at: file.updated_at.clone(),
                });
            }
        }
        tasks.sort_by(|left, right| {
            left.updated_at
                .cmp(&right.updated_at)
                .then_with(|| left.source_file.cmp(&right.source_file))
        });
        Ok(tasks)
    }

    /// Return recovery entries whose audit copy disappeared after a pending
    /// task was persisted. These cannot be resumed, but must remain visible in
    /// the UI so the user knows to rerun S1-F1.
    pub fn lost_copy_tasks(&self) -> Result<Vec<S1RecoveryTaskSummary>> {
        if !self.directory.exists() {
            return Ok(Vec::new());
        }
        self.recover_interrupted_replacements()?;
        self.reconcile_interrupted_copies()?;
        let mut tasks = Vec::new();
        for entry in fs::read_dir(&self.directory)
            .with_context(|| format!("扫描 S1 恢复状态目录 {}", self.directory.display()))?
        {
            let entry = entry?;
            let path = entry.path();
            if !path.is_file() || path.extension().and_then(|ext| ext.to_str()) != Some("json") {
                continue;
            }
            let bytes = fs::read(&path)
                .with_context(|| format!("读取 S1 恢复任务清单 {}", path.display()))?;
            let manifest: S1RecoveryManifest = serde_json::from_slice(&bytes)
                .with_context(|| format!("解析 S1 恢复任务清单 {}", path.display()))?;
            ensure!(
                manifest.schema_version == RECOVERY_SCHEMA_VERSION
                    && manifest.feature_id == "S1-F01",
                "S1 恢复状态目录中存在不兼容清单：{}",
                path.display()
            );
            for file in manifest.files.iter().filter(|file| {
                file.status == S1RecoveryStatus::Failed
                    && file.reason.as_deref().is_some_and(|reason| {
                        reason.contains("待恢复任务的审核副本已丢失或路径无效")
                    })
            }) {
                tasks.push(S1RecoveryTaskSummary {
                    run_id: manifest.run_id.clone(),
                    file_id: file.file_id.clone(),
                    source_file: file.source_file.clone(),
                    audit_copy: file.audit_copy.clone(),
                    status: file.status.as_str(),
                    reason: file.reason.clone(),
                    updated_at: file.updated_at.clone(),
                });
            }
        }
        tasks.sort_by(|left, right| {
            left.updated_at
                .cmp(&right.updated_at)
                .then_with(|| left.source_file.cmp(&right.source_file))
        });
        Ok(tasks)
    }

    fn manifest_path(&self, run_id: &str) -> Result<PathBuf> {
        validate_run_id(run_id)?;
        Ok(self.directory.join(format!("{run_id}.json")))
    }

    fn recover_interrupted_replacements(&self) -> Result<()> {
        if !self.directory.exists() {
            return Ok(());
        }
        for entry in fs::read_dir(&self.directory)? {
            let path = entry?.path();
            let name = path
                .file_name()
                .and_then(|value| value.to_str())
                .unwrap_or_default();
            let Some(original_name) = name.strip_suffix(".json.bak") else {
                continue;
            };
            let original = self.directory.join(format!("{original_name}.json"));
            if original.exists() {
                fs::remove_file(&path)
                    .with_context(|| format!("清理旧的 S1 恢复清单备份 {}", path.display()))?;
            } else {
                fs::rename(&path, &original).with_context(|| {
                    format!("恢复中断写入的 S1 恢复清单 {}", original.display())
                })?;
            }
        }
        Ok(())
    }

    fn reconcile_interrupted_copies(&self) -> Result<()> {
        for entry in fs::read_dir(&self.directory)
            .with_context(|| format!("扫描 S1 恢复状态目录 {}", self.directory.display()))?
        {
            let path = entry?.path();
            if !path.is_file() || path.extension().and_then(|ext| ext.to_str()) != Some("json") {
                continue;
            }
            let bytes = fs::read(&path)
                .with_context(|| format!("读取 S1 恢复任务清单 {}", path.display()))?;
            let mut manifest: S1RecoveryManifest = serde_json::from_slice(&bytes)
                .with_context(|| format!("解析 S1 恢复任务清单 {}", path.display()))?;
            let mut changed = false;
            for file in &mut manifest.files {
                if file.status == S1RecoveryStatus::Copying {
                    let copy_path = file.audit_copy.clone();
                    let copy_exists = copy_path.as_deref().is_some_and(Path::is_file);
                    file.status = if let Some(copy_path) = copy_path.filter(|_| copy_exists) {
                        match workbook_recovery_fingerprints(&copy_path) {
                            Ok((
                                fingerprint,
                                normalized_fingerprint,
                                semantic_fingerprint,
                                formula_count,
                            )) if formula_count > 0 => {
                                file.formula_fingerprint = Some(fingerprint);
                                file.formula_fingerprint_normalized = Some(normalized_fingerprint);
                                file.audit_copy_semantic_fingerprint = Some(semantic_fingerprint);
                                file.reason = Some(
                                    "上次运行在副本生成后中断；请确认该副本已由 Excel/WPS/LibreOffice 重新计算并保存，再使用 S1-F3 恢复。".to_owned(),
                                );
                                S1RecoveryStatus::WaitingManualRecalc
                            }
                            Ok(_) => {
                                file.reason = Some(
                                    "上次运行生成的副本中没有可恢复公式；请重新运行 S1-F1。"
                                        .to_owned(),
                                );
                                S1RecoveryStatus::Failed
                            }
                            Err(error) => {
                                file.reason = Some(format!(
                                    "上次运行生成的审核副本无法读取公式：{error:#}；请重新运行 S1-F1。"
                                ));
                                S1RecoveryStatus::Failed
                            }
                        }
                    } else {
                        file.reason = Some(
                            "上次运行在审核副本生成前中断，且没有可恢复副本；请重新运行 S1-F1。"
                                .to_owned(),
                        );
                        S1RecoveryStatus::Failed
                    };
                    file.updated_at = timestamp();
                    changed = true;
                    continue;
                }
                if file.status.is_pending()
                    && !file.audit_copy.as_deref().is_some_and(Path::is_file)
                {
                    file.status = S1RecoveryStatus::Failed;
                    file.reason = Some(
                        "待恢复任务的审核副本已丢失或路径无效；无法安全恢复，请重新运行 S1-F1。"
                            .to_owned(),
                    );
                    file.updated_at = timestamp();
                    changed = true;
                }
            }
            if changed {
                self.save(&mut manifest)?;
            }
        }
        Ok(())
    }
}

/// Hash the ordered worksheet/formula map and return its formula-cell count.
/// Cached values are intentionally excluded because a user must recalculate
/// and save those values before S1-F3 resumes the audit.
pub fn workbook_formula_fingerprint(path: impl AsRef<Path>) -> Result<(String, usize)> {
    let (exact, _, count) = workbook_formula_fingerprints(path)?;
    Ok((exact, count))
}

pub fn workbook_formula_fingerprints(path: impl AsRef<Path>) -> Result<(String, String, usize)> {
    let path = path.as_ref();
    let workbook = excel_reader::read_workbook_with_unknown_error_tokens_as_text(path)
        .with_context(|| format!("读取审核副本公式指纹：{}", path.display()))?
        .workbook;
    Ok(formula_fingerprints(&workbook)?)
}

/// Return formula fingerprints plus a semantic snapshot of the data and rules
/// that must remain unchanged while a user recalculates the audit copy.
/// Formula result cells are excluded from the semantic snapshot so Excel and
/// LibreOffice can populate them without invalidating the manual-recovery task.
pub fn workbook_recovery_fingerprints(
    path: impl AsRef<Path>,
) -> Result<(String, String, String, usize)> {
    let path = path.as_ref();
    let workbook = excel_reader::read_workbook_with_unknown_error_tokens_as_text(path)
        .with_context(|| format!("读取审核副本完整性指纹：{}", path.display()))?
        .workbook;
    let (exact, normalized, formula_count) = formula_fingerprints(&workbook)?;
    let semantic = workbook_semantic_fingerprint(path, &workbook)?;
    Ok((exact, normalized, semantic, formula_count))
}

fn workbook_semantic_fingerprint(
    path: &Path,
    workbook: &excel_reader::WorkbookDump,
) -> Result<String> {
    let snapshot = workbook_semantic_snapshot(path, workbook)?;
    let encoded =
        serde_json::to_vec(&snapshot).context("序列化审核副本非公式数据与审核规则指纹")?;
    Ok(format!("{:x}", Sha256::digest(encoded)))
}

fn workbook_semantic_snapshot(path: &Path, workbook: &excel_reader::WorkbookDump) -> Result<Value> {
    let mut defined_names = excel_reader::read_defined_names(path)
        .with_context(|| format!("读取审核副本命名区域完整性指纹：{}", path.display()))?
        .into_iter()
        .map(|name| {
            let raw_reference = if matches!(
                name.parse_status,
                excel_reader::DefinedNameParseStatus::Areas
            ) {
                Value::Null
            } else {
                Value::String(excel_reader::canonicalize_formula_boolean_literals(
                    &name.raw_reference,
                ))
            };
            serde_json::json!({
                "name": name.name,
                "scope": name.scope,
                "parseStatus": name.parse_status,
                "areas": name.areas,
                "rawReference": raw_reference,
            })
        })
        .collect::<Vec<_>>();
    defined_names.sort_by_key(|name| {
        (
            name["name"].as_str().unwrap_or_default().to_owned(),
            serde_json::to_string(&name["scope"]).unwrap_or_default(),
        )
    });
    let sheets = workbook
        .sheets
        .iter()
        .map(|sheet| {
            let non_formula_values = sheet
                .values
                .iter()
                .filter(|(cell, _)| !sheet.formulas.contains_key(*cell))
                .collect::<std::collections::BTreeMap<_, _>>();
            serde_json::json!({
                "name": &sheet.name,
                "nonFormulaValues": non_formula_values,
                "mergedRanges": &sheet.merges,
                "mergeStatus": &sheet.merge_status,
                "conditionalFormatting": conditional_formatting_snapshot(sheet),
                "conditionalFormattingStatus": &sheet.conditional_formatting_status,
            })
        })
        .collect::<Vec<_>>();
    Ok(serde_json::json!({
        "definedNames": defined_names,
        "sheets": sheets,
    }))
}

fn conditional_formatting_snapshot(sheet: &excel_reader::SheetDump) -> Vec<Value> {
    let mut rules = Vec::<(u32, usize, Value)>::new();
    let mut order = 0usize;
    for formatting in &sheet.conditional_formatting {
        let applies_to = formatting
            .get("attributes")
            .and_then(|attributes| attributes.get("sqref"))
            .and_then(Value::as_str)
            .unwrap_or_default();
        let Some(children) = formatting.get("content").and_then(Value::as_array) else {
            continue;
        };
        for rule in children
            .iter()
            .filter(|child| child.get("name").and_then(Value::as_str) == Some("cfRule"))
        {
            let attributes = rule
                .get("attributes")
                .and_then(Value::as_object)
                .cloned()
                .unwrap_or_default();
            let priority = attributes
                .get("priority")
                .and_then(Value::as_str)
                .and_then(|value| value.parse::<u32>().ok())
                .unwrap_or(order as u32 + 1);
            let mut semantic_attributes = serde_json::Map::new();
            for (key, value) in attributes {
                if matches!(key.as_str(), "dxfId" | "priority") {
                    continue;
                }
                if (matches!(
                    key.as_str(),
                    "aboveAverage" | "bottom" | "equalAverage" | "percent" | "stopIfTrue"
                ) && matches!(value.as_str(), Some("0" | "false")))
                    || (key == "rank" && value.as_str() == Some("0"))
                    || (key == "text" && value.as_str() == Some(""))
                {
                    continue;
                }
                semantic_attributes.insert(key, value);
            }
            let formulas = rule
                .get("content")
                .and_then(Value::as_array)
                .into_iter()
                .flatten()
                .filter(|child| child.get("name").and_then(Value::as_str) == Some("formula"))
                .map(|formula| {
                    excel_reader::canonicalize_formula_boolean_literals(&conditional_node_text(
                        formula,
                    ))
                })
                .collect::<Vec<_>>();
            rules.push((
                priority,
                order,
                serde_json::json!({
                    "appliesTo": applies_to,
                    "attributes": semantic_attributes,
                    "formulas": formulas,
                }),
            ));
            order += 1;
        }
    }
    rules.sort_by_key(|(priority, order, _)| (*priority, *order));
    rules
        .into_iter()
        .enumerate()
        .map(|(priority_order, (_, _, mut rule))| {
            rule["priorityOrder"] = serde_json::json!(priority_order + 1);
            rule
        })
        .collect()
}

fn conditional_node_text(node: &Value) -> String {
    if let Some(text) = node.get("text").and_then(Value::as_str) {
        return text.to_owned();
    }
    node.get("content")
        .and_then(Value::as_array)
        .into_iter()
        .flatten()
        .map(conditional_node_text)
        .collect()
}

fn formula_fingerprints(workbook: &excel_reader::WorkbookDump) -> Result<(String, String, usize)> {
    let formulas = workbook
        .sheets
        .iter()
        .map(|sheet| (&sheet.name, &sheet.formulas))
        .collect::<Vec<_>>();
    let count = formulas.iter().map(|(_, formulas)| formulas.len()).sum();
    let exact = formula_map_sha256(&formulas)?;
    let normalized = workbook
        .sheets
        .iter()
        .map(|sheet| {
            (
                &sheet.name,
                sheet
                    .formulas
                    .iter()
                    .map(|(cell, formula)| {
                        (
                            cell.clone(),
                            excel_reader::canonicalize_formula_boolean_literals(formula),
                        )
                    })
                    .collect::<std::collections::BTreeMap<_, _>>(),
            )
        })
        .collect::<Vec<_>>();
    let normalized = formula_map_sha256(&normalized)?;
    Ok((exact, normalized, count))
}

/// Validate that recalculation preserved the original formulas and non-formula
/// audit inputs. A missing semantic fingerprint identifies an older manifest
/// and retains its historical formula-only validation behavior.
pub fn validate_recovery_fingerprints(
    expected_formula_fingerprint: &str,
    expected_normalized_formula_fingerprint: Option<&str>,
    expected_semantic_fingerprint: Option<&str>,
    current_formula_fingerprint: &str,
    current_normalized_formula_fingerprint: &str,
    current_semantic_fingerprint: &str,
) -> Result<()> {
    ensure!(
        current_formula_fingerprint == expected_formula_fingerprint
            || expected_normalized_formula_fingerprint
                == Some(current_normalized_formula_fingerprint),
        "S1-F3 副本公式文本与 F1 创建的审核副本不一致（仅兼容 LibreOffice TRUE/FALSE 等价规范化），拒绝继续"
    );
    if let Some(expected_semantic_fingerprint) = expected_semantic_fingerprint {
        ensure!(
            current_semantic_fingerprint == expected_semantic_fingerprint,
            "S1-F3 副本中的非公式输入、条件格式、合并区域或命名区域已变化；为避免按修改后的数据审核，请恢复审核副本或重新运行 S1-F1"
        );
    }
    Ok(())
}

fn formula_map_sha256<T: Serialize>(formulas: &T) -> Result<String> {
    let encoded = serde_json::to_vec(formulas).context("序列化审核副本公式指纹")?;
    let digest = Sha256::digest(encoded);
    Ok(format!("{digest:x}"))
}

pub fn file_sha256(path: impl AsRef<Path>) -> Result<String> {
    let path = path.as_ref();
    let mut file =
        fs::File::open(path).with_context(|| format!("读取文件摘要：{}", path.display()))?;
    let mut digest = Sha256::new();
    let mut buffer = [0_u8; 64 * 1024];
    loop {
        let count = file.read(&mut buffer)?;
        if count == 0 {
            break;
        }
        digest.update(&buffer[..count]);
    }
    Ok(format!("{:x}", digest.finalize()))
}

pub fn stable_file_id(path: impl AsRef<Path>) -> String {
    let value = path.as_ref().to_string_lossy().replace('\\', "/");
    let value = if cfg!(windows) {
        value.to_lowercase()
    } else {
        value
    };
    format!("{:x}", Sha256::digest(value.as_bytes()))
}

fn validate_run_id(run_id: &str) -> Result<()> {
    ensure!(!run_id.is_empty(), "S1 恢复 run_id 不能为空");
    ensure!(
        run_id
            .chars()
            .all(|character| character.is_ascii_alphanumeric() || matches!(character, '-' | '_')),
        "S1 恢复 run_id 包含不安全字符"
    );
    Ok(())
}

fn atomic_replace(path: &Path, bytes: &[u8]) -> Result<()> {
    let parent = path.parent().context("S1 恢复清单缺少父目录")?;
    for attempt in 0..64_u32 {
        let nonce = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
        let temp = parent.join(format!(
            ".s1-recovery-{}-{nonce}-{attempt}.tmp",
            std::process::id()
        ));
        let write_result = (|| -> Result<()> {
            let mut file = fs::OpenOptions::new()
                .write(true)
                .create_new(true)
                .open(&temp)?;
            file.write_all(bytes)?;
            file.sync_all()?;
            Ok(())
        })();
        if let Err(error) = write_result {
            let _ = fs::remove_file(&temp);
            if error
                .downcast_ref::<std::io::Error>()
                .is_some_and(|io| io.kind() == std::io::ErrorKind::AlreadyExists)
            {
                continue;
            }
            return Err(error).context("写入 S1 恢复清单临时文件");
        }
        if !path.exists() {
            if let Err(error) = fs::rename(&temp, path) {
                let _ = fs::remove_file(&temp);
                return Err(error).with_context(|| format!("发布 S1 恢复清单 {}", path.display()));
            }
            return Ok(());
        }

        let backup = path.with_file_name(format!(
            "{}.json.bak",
            path.file_stem()
                .and_then(|value| value.to_str())
                .unwrap_or("run")
        ));
        if backup.exists() {
            let _ = fs::remove_file(&backup);
        }
        fs::rename(path, &backup)
            .with_context(|| format!("备份现有 S1 恢复清单 {}", path.display()))?;
        match fs::rename(&temp, path) {
            Ok(()) => {
                let _ = fs::remove_file(backup);
                return Ok(());
            }
            Err(error) => {
                let _ = fs::rename(&backup, path);
                let _ = fs::remove_file(&temp);
                return Err(error)
                    .with_context(|| format!("原子替换 S1 恢复清单 {}", path.display()));
            }
        }
    }
    bail!("无法分配唯一的 S1 恢复清单临时文件")
}

fn timestamp() -> String {
    chrono::Local::now().to_rfc3339()
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    fn temp_dir() -> PathBuf {
        let nonce = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock after epoch")
            .as_nanos();
        let path = env::temp_dir().join(format!(
            "audit_s1_recovery_{}_{}",
            std::process::id(),
            nonce
        ));
        fs::create_dir_all(&path).expect("create recovery test directory");
        path
    }

    #[test]
    fn pending_manual_tasks_survive_store_recreation_and_disappear_when_completed() -> Result<()> {
        let root = temp_dir();
        let store = S1RecoveryStore::new(root.join("state"));
        let audit_copy = root.join("output/run/audit-copy.xlsx");
        fs::create_dir_all(audit_copy.parent().expect("audit copy parent"))?;
        fs::write(&audit_copy, b"audit copy placeholder")?;
        let mut manifest = S1RecoveryManifest::new(
            "S1F1-run_1",
            root.join("output"),
            root.join("output/run"),
            "2026-10",
            "batch-1",
            "2026-10-28 12:00:00",
            None,
            None,
            Vec::new(),
            vec![json!({"sourceFile":"skipped.xlsx"})],
            vec![S1RecoveryFile::new(
                root.join("source.xlsx"),
                root.join("template.xlsx"),
                "template-hash".to_owned(),
                Some(audit_copy),
                S1RecoveryStatus::WaitingManualRecalc,
            )],
        );
        store.save(&mut manifest)?;

        let recreated = S1RecoveryStore::new(root.join("state"));
        let pending = recreated.pending_tasks()?;
        assert_eq!(pending.len(), 1);
        assert_eq!(pending[0].status, "WAITING_MANUAL_RECALC");
        assert_eq!(pending[0].run_id, "S1F1-run_1");

        let mut persisted = recreated.load("S1F1-run_1")?;
        persisted.files[0].status = S1RecoveryStatus::Completed;
        recreated.save(&mut persisted)?;
        assert!(recreated.pending_tasks()?.is_empty());
        fs::remove_dir_all(root)?;
        Ok(())
    }

    #[test]
    fn recovery_store_rejects_path_traversal_run_ids() {
        assert!(validate_run_id("../outside").is_err());
        assert!(validate_run_id("run-01_abc").is_ok());
    }

    #[test]
    fn interrupted_copying_tasks_are_reconciled_without_exposing_missing_copies() -> Result<()> {
        let root = temp_dir();
        let store = S1RecoveryStore::new(root.join("state"));
        let existing_copy = root.join("output/run/ready.xlsx");
        fs::create_dir_all(existing_copy.parent().expect("copy parent"))?;
        let formula_fixture =
            Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/synthetic_probe.xlsx");
        fs::copy(formula_fixture, &existing_copy)?;
        let mut manifest = S1RecoveryManifest::new(
            "recover-interrupted",
            root.join("output"),
            root.join("output/run"),
            "2026-10",
            "batch-2",
            "2026-10-28 12:00:00",
            None,
            None,
            Vec::new(),
            Vec::new(),
            vec![
                S1RecoveryFile::new(
                    root.join("source-ready.xlsx"),
                    root.join("template.xlsx"),
                    "template-hash".to_owned(),
                    Some(existing_copy.clone()),
                    S1RecoveryStatus::Copying,
                ),
                S1RecoveryFile::new(
                    root.join("source-missing.xlsx"),
                    root.join("template.xlsx"),
                    "template-hash".to_owned(),
                    Some(root.join("output/run/missing.xlsx")),
                    S1RecoveryStatus::Copying,
                ),
            ],
        );
        store.save(&mut manifest)?;

        let tasks = S1RecoveryStore::new(root.join("state")).pending_tasks()?;
        assert_eq!(tasks.len(), 1);
        assert_eq!(tasks[0].status, "WAITING_MANUAL_RECALC");
        let reconciled = store.load("recover-interrupted")?;
        assert_eq!(
            reconciled.files[0].status,
            S1RecoveryStatus::WaitingManualRecalc
        );
        assert!(reconciled.files[0]
            .audit_copy_semantic_fingerprint
            .as_deref()
            .is_some_and(|fingerprint| !fingerprint.is_empty()));
        assert_eq!(reconciled.files[1].status, S1RecoveryStatus::Failed);
        assert!(reconciled.files[0]
            .reason
            .as_deref()
            .unwrap()
            .contains("中断"));
        assert!(reconciled.files[1]
            .reason
            .as_deref()
            .unwrap()
            .contains("重新运行 S1-F1"));
        fs::remove_dir_all(root)?;
        Ok(())
    }

    #[test]
    fn persisted_manual_task_with_missing_copy_is_marked_failed_with_recovery_reason() -> Result<()>
    {
        let root = temp_dir();
        let state_dir = root.join("state");
        let store = S1RecoveryStore::new(&state_dir);
        let missing_copy = root.join("output/run/missing-after-restart.xlsx");
        let mut manifest = S1RecoveryManifest::new(
            "recover-missing-copy",
            root.join("output"),
            root.join("output/run"),
            "2026-10",
            "batch-missing-copy",
            "2026-10-29 12:00:00",
            None,
            None,
            Vec::new(),
            Vec::new(),
            vec![S1RecoveryFile::new(
                root.join("source.xlsx"),
                root.join("template.xlsx"),
                "template-hash".to_owned(),
                Some(missing_copy),
                S1RecoveryStatus::WaitingManualRecalc,
            )],
        );
        store.save(&mut manifest)?;

        let restarted_store = S1RecoveryStore::new(&state_dir);
        assert!(restarted_store.pending_tasks()?.is_empty());
        let reconciled = restarted_store.load("recover-missing-copy")?;
        assert_eq!(reconciled.files[0].status, S1RecoveryStatus::Failed);
        assert!(reconciled.files[0]
            .reason
            .as_deref()
            .is_some_and(|reason| reason.contains("审核副本已丢失")));
        let lost = restarted_store.lost_copy_tasks()?;
        assert_eq!(lost.len(), 1);
        assert_eq!(lost[0].status, "FAILED");
        assert!(lost[0]
            .reason
            .as_deref()
            .is_some_and(|reason| reason.contains("重新运行 S1-F1")));

        fs::remove_dir_all(root)?;
        Ok(())
    }

    #[test]
    fn formula_fingerprint_tracks_formula_text_and_cell_count() -> Result<()> {
        let fixture =
            Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/synthetic_probe.xlsx");
        let (fingerprint, formula_count) = workbook_formula_fingerprint(&fixture)?;
        assert!(!fingerprint.is_empty());
        assert!(formula_count > 0);
        assert_eq!(
            (fingerprint.clone(), formula_count),
            workbook_formula_fingerprint(fixture)?
        );
        Ok(())
    }

    #[test]
    fn recovery_semantic_fingerprint_ignores_formula_caches_but_tracks_input_cells() -> Result<()> {
        let root = temp_dir();
        let fixture =
            Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/synthetic_probe.xlsx");
        let cacheless = root.join("cacheless.xlsx");
        let changed_input = root.join("changed-input.xlsx");

        let (_, _, baseline_semantic, formula_count) = workbook_recovery_fingerprints(&fixture)?;
        assert!(formula_count > 0);
        ooxml_patch::clear_formula_caches_for_recalculation(&fixture, &cacheless)?;
        let (_, _, cacheless_semantic, cacheless_formula_count) =
            workbook_recovery_fingerprints(&cacheless)?;
        assert_eq!(formula_count, cacheless_formula_count);
        assert_eq!(
            baseline_semantic, cacheless_semantic,
            "clearing formula caches must not invalidate the non-formula audit snapshot"
        );

        let source_dump = excel_reader::read_workbook(&fixture)?;
        let (sheet_name, cell) = source_dump
            .sheets
            .iter()
            .find_map(|sheet| {
                sheet
                    .values
                    .keys()
                    .find(|cell| !sheet.formulas.contains_key(*cell))
                    .map(|cell| (sheet.name.clone(), cell.clone()))
            })
            .context("合成审核样本缺少非公式输入单元格")?;
        let mut edited = umya_spreadsheet::reader::xlsx::read(&cacheless)?;
        edited
            .sheet_by_name_mut(&sheet_name)?
            .cell_mut(cell.as_str())
            .set_value_string("RECOVERY_INPUT_CHANGED");
        umya_spreadsheet::writer::xlsx::write(&edited, &changed_input)?;
        let (_, _, changed_semantic, changed_formula_count) =
            workbook_recovery_fingerprints(&changed_input)?;
        assert_eq!(cacheless_formula_count, changed_formula_count);
        assert_ne!(
            cacheless_semantic, changed_semantic,
            "manual recovery must reject changed non-formula input values"
        );

        fs::remove_dir_all(root)?;
        Ok(())
    }

    #[test]
    fn recovery_fingerprint_validation_rejects_changed_inputs_and_accepts_legacy_manifests() {
        assert!(validate_recovery_fingerprints(
            "formula-exact",
            Some("formula-normalized"),
            Some("inputs-original"),
            "formula-exact",
            "formula-normalized",
            "inputs-changed",
        )
        .is_err());
        assert!(validate_recovery_fingerprints(
            "formula-exact",
            Some("formula-normalized"),
            Some("inputs-original"),
            "formula-other",
            "formula-normalized",
            "inputs-original",
        )
        .is_ok());
        assert!(validate_recovery_fingerprints(
            "formula-exact",
            None,
            None,
            "formula-exact",
            "formula-normalized",
            "legacy-no-semantic-fingerprint",
        )
        .is_ok());
    }

    #[test]
    fn recovery_semantic_fingerprint_survives_libreoffice_recalculation_when_configured(
    ) -> Result<()> {
        let Some(sidecar) = env::var_os("AUDIT_S1_F3_TEST_SIDECAR") else {
            return Ok(());
        };
        let root = temp_dir();
        let fixture =
            Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/synthetic_probe.xlsx");
        let output = root.join("libreoffice-recalculated.xlsx");
        let before_dump = excel_reader::read_workbook_with_unknown_error_tokens_as_text(&fixture)?;
        let before_snapshot = workbook_semantic_snapshot(&fixture, &before_dump.workbook)?;
        let (before_exact, before_normalized, before_semantic, before_count) =
            workbook_recovery_fingerprints(&fixture)?;
        let result = std::process::Command::new(sidecar)
            .args([
                "recalculate",
                "--input",
                fixture.to_str().context("fixture path is not UTF-8")?,
                "--output",
                output.to_str().context("output path is not UTF-8")?,
                "--engine",
                "lo",
            ])
            .output()
            .context("运行 LibreOffice S1-F3 恢复摘要集成验证")?;
        ensure!(
            result.status.success(),
            "LibreOffice sidecar failed: {}",
            String::from_utf8_lossy(&result.stderr)
        );
        let after_dump = excel_reader::read_workbook_with_unknown_error_tokens_as_text(&output)?;
        let after_snapshot = workbook_semantic_snapshot(&output, &after_dump.workbook)?;
        let (after_exact, after_normalized, after_semantic, after_count) =
            workbook_recovery_fingerprints(&output)?;
        ensure!(
            before_count == after_count,
            "LibreOffice changed the formula-cell count"
        );
        ensure!(
            before_exact == after_exact || before_normalized == after_normalized,
            "LibreOffice changed formula text outside the accepted boolean serialization"
        );
        if before_semantic != after_semantic {
            eprintln!(
                "LibreOffice defined names changed: {}",
                before_snapshot["definedNames"] != after_snapshot["definedNames"]
            );
            if before_snapshot["definedNames"] != after_snapshot["definedNames"] {
                eprintln!(
                    "LibreOffice defined names before: {:#}",
                    before_snapshot["definedNames"]
                );
                eprintln!(
                    "LibreOffice defined names after: {:#}",
                    after_snapshot["definedNames"]
                );
            }
            for key in [
                "nonFormulaValues",
                "mergedRanges",
                "mergeStatus",
                "conditionalFormatting",
                "conditionalFormattingStatus",
                "dataValidations",
                "dataValidationStatus",
            ] {
                let before = &before_snapshot["sheets"][0][key];
                let after = &after_snapshot["sheets"][0][key];
                if before == after {
                    continue;
                }
                if key == "nonFormulaValues" {
                    let before_map = before.as_object().context("before values are not a map")?;
                    let after_map = after.as_object().context("after values are not a map")?;
                    for cell in before_map.keys().chain(after_map.keys()) {
                        if before_map.get(cell) != after_map.get(cell) {
                            eprintln!(
                                "LibreOffice changed non-formula {cell}: {:?} -> {:?}",
                                before_map.get(cell),
                                after_map.get(cell)
                            );
                        }
                    }
                } else {
                    eprintln!("LibreOffice changed {key}: {before:#} -> {after:#}");
                }
            }
            bail!("LibreOffice recalculation changed non-formula audit inputs or rule metadata");
        }
        fs::remove_dir_all(root)?;
        Ok(())
    }

    #[test]
    fn normalized_fingerprint_accepts_only_lo_boolean_serialization() -> Result<()> {
        let fixture =
            Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/synthetic_probe.xlsx");
        let mut workbook = excel_reader::read_workbook(fixture)?;
        workbook.sheets[0].formulas.clear();
        workbook.sheets[0]
            .formulas
            .insert("D2".to_owned(), "VLOOKUP(A1,Ref!A:B,2,FALSE)".to_owned());
        let (before_exact, before_normalized, count) = formula_fingerprints(&workbook)?;
        workbook.sheets[0]
            .formulas
            .insert("D2".to_owned(), "VLOOKUP(A1,Ref!A:B,2,FALSE())".to_owned());
        let (after_exact, after_normalized, after_count) = formula_fingerprints(&workbook)?;

        assert_ne!(before_exact, after_exact);
        assert_eq!(before_normalized, after_normalized);
        assert_eq!(count, 1);
        assert_eq!(after_count, count);

        workbook.sheets[0].formulas.insert(
            "D2".to_owned(),
            "VLOOKUP(A1,Other!A:B,2,FALSE())".to_owned(),
        );
        let (_, changed_reference_fingerprint, changed_count) = formula_fingerprints(&workbook)?;
        assert_ne!(before_normalized, changed_reference_fingerprint);
        assert_eq!(changed_count, count);
        Ok(())
    }
}
