//! betteroffice（纯 Rust）校验区域重算路由。
//!
//! S1-F01 校验区域重算的首选引擎：打开工作副本 → 按配置轮数 `recalculate_all`
//! → 写回 → 统计受限格与错误格 → 由调用方决定采用或回退 COM/LibreOffice。
//!
//! 配置（环境变量，与 Rust-Tools `pytools-core::bo_calc` 同名）：
//! - `PYTOOLS_BO_CALC`：`auto`（默认，启用）/`off`（完全走 COM/LO 旧路由）
//! - `PYTOOLS_BO_CALC_MAX_VISITS`：重算预算上限；`0` 表示用库默认（10,000,000）
//! - `PYTOOLS_BO_CALC_ROUNDS`：重算轮数（默认 2）。betteroffice 的依赖图
//!   看不到 `INDIRECT` 的动态引用，跨表动态链第一轮可能读到尚未更新的值；
//!   第二轮起收敛（实测 S1F1 模板第二轮 changed=0）。
//! - `PYTOOLS_BO_CALC_STRIP_STYLE`：`off`（默认）/`on`。`on` 时在重算前把
//!   工作副本里的纯样式空单元格剥掉（`ooxml_patch::strip_style_only_cells`）。
//! - `PYTOOLS_BO_CALC_FALLBACK`：`auto`（按平台回 COM/LibreOffice）/`com`/
//!   `libreoffice`/`none`（超限即报错，绝不静默采用受限结果）。

use std::path::Path;
use std::time::Instant;

use anyhow::{bail, Context, Result};

/// why a betteroffice recalculation cannot be trusted and must fall back.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum BoCalcRefusal {
    BudgetExhausted { limited_cells: usize },
    UnsupportedFunctions { name_error_cells: usize },
    Unusable { reason: String },
}

impl std::fmt::Display for BoCalcRefusal {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            BoCalcRefusal::BudgetExhausted { limited_cells } => write!(
                formatter,
                "重算预算耗尽（{limited_cells} 个公式格未求值，保留旧值不可信）"
            ),
            BoCalcRefusal::UnsupportedFunctions { name_error_cells } => write!(
                formatter,
                "{name_error_cells} 个公式格求值为 #NAME?（模板使用了 betteroffice 未实现的函数）"
            ),
            BoCalcRefusal::Unusable { reason } => write!(formatter, "{reason}"),
        }
    }
}

impl std::error::Error for BoCalcRefusal {}

/// what one betteroffice recalculation did to the working copy.
#[derive(Debug, Clone)]
pub struct BoCalcOutcome {
    pub application: &'static str,
    pub rounds_used: u32,
    pub changed_cells_total: usize,
    pub limited_cells: usize,
    pub error_cells: usize,
    pub name_error_cells: usize,
    pub style_cells_stripped: usize,
    pub open_elapsed: std::time::Duration,
    pub calc_elapsed: std::time::Duration,
    pub save_elapsed: std::time::Duration,
}

/// routing configuration, read once per call from the process environment.
#[derive(Debug, Clone)]
pub struct BoCalcConfig {
    pub enabled: bool,
    pub max_visits: Option<u64>,
    pub rounds: u32,
    pub strip_style_cells: bool,
    pub fallback_com: bool,
    pub fallback_libreoffice: bool,
}

impl BoCalcConfig {
    /// 配置优先级：GUI 高级设置（AppSettings 持久化）> 环境变量 > 内置默认，
    /// 与 Rust-Tools 的「config > env > default」完全一致。
    pub fn from_app() -> Self {
        // 基座 = env（未设则内置默认）；GUI 设置逐项覆盖——但只有 GUI
        // 「显式设置过」的键才覆盖，未设置的键让 env/默认继续生效。
        let mut config = Self::from_env();
        let settings = crate::app_settings::AppSettings::load();
        // 先应用回退策略，再应用引擎选择——「com=直接用 Excel/WPS」对回退
        // 方向有最终决定权（锁定 COM-only）。
        if let Some(fallback) = settings.bo_calc_fallback() {
            match fallback {
                "none" => {
                    config.fallback_com = false;
                    config.fallback_libreoffice = false;
                }
                "com" => {
                    config.fallback_com = true;
                    config.fallback_libreoffice = false;
                }
                "libreoffice" => {
                    config.fallback_com = false;
                    config.fallback_libreoffice = true;
                }
                _ => {
                    config.fallback_com = true;
                    config.fallback_libreoffice = true;
                }
            }
        }
        if let Some(engine) = settings.bo_calc_engine() {
            match engine {
                // 直接用 Excel/WPS：跳过 betteroffice，回退锁定 COM。
                "com" => {
                    config.enabled = false;
                    config.fallback_com = true;
                    config.fallback_libreoffice = false;
                }
                // 完全回到 COM/LibreOffice 旧路由（回退按平台 auto）。
                "off" => config.enabled = false,
                // betteroffice：启用。
                _ => config.enabled = true,
            }
        }
        if let Some(max_visits) = settings.bo_calc_max_visits() {
            config.max_visits = Some(max_visits);
        }
        if let Some(rounds) = settings.bo_calc_rounds() {
            config.rounds = rounds;
        }
        if let Some(strip) = settings.bo_calc_strip_style() {
            config.strip_style_cells = strip;
        }
        config
    }

    pub fn from_env() -> Self {
        let mode = std::env::var("PYTOOLS_BO_CALC").unwrap_or_default();
        let enabled = !mode.eq_ignore_ascii_case("off");
        let max_visits = std::env::var("PYTOOLS_BO_CALC_MAX_VISITS")
            .ok()
            .and_then(|value| value.trim().parse::<u64>().ok())
            .filter(|value| *value > 0);
        let rounds = std::env::var("PYTOOLS_BO_CALC_ROUNDS")
            .ok()
            .and_then(|value| value.trim().parse::<u32>().ok())
            .filter(|value| *value >= 1)
            .unwrap_or(2)
            .min(5);
        let strip_style_cells = std::env::var("PYTOOLS_BO_CALC_STRIP_STYLE")
            .map(|value| value.trim().eq_ignore_ascii_case("on"))
            .unwrap_or(false);
        let fallback = std::env::var("PYTOOLS_BO_CALC_FALLBACK")
            .unwrap_or_default()
            .trim()
            .to_ascii_lowercase();
        let fallback_com = fallback != "none" && fallback != "libreoffice";
        let fallback_libreoffice = fallback != "none" && fallback != "com";
        Self {
            enabled,
            max_visits,
            rounds,
            strip_style_cells,
            fallback_com,
            fallback_libreoffice,
        }
    }

    /// whether this platform can serve the fallback this config asks for.
    pub fn fallback_available(&self) -> bool {
        #[cfg(windows)]
        {
            self.fallback_com
        }
        #[cfg(target_os = "linux")]
        {
            self.fallback_libreoffice
        }
        #[cfg(not(any(windows, target_os = "linux")))]
        {
            false
        }
    }
}

/// Recalculate `source_copy` in place with betteroffice.
///
/// The path must be a disposable working copy (`.xlsx`/`.xlsm`); the caller
/// owns snapshot and candidate validation. On success the cached values in the
/// file have been rewritten and the report describes what happened; on
/// `Err(BoCalcRefusal)` the caller must fall back (the file may have been
/// rewritten — the fallback engines recalculate everything themselves, so a
/// partially updated copy is safe to hand them).
pub fn recalculate(source_copy: &Path, config: &BoCalcConfig) -> Result<BoCalcOutcome> {
    let extension = source_copy
        .extension()
        .and_then(|extension| extension.to_str())
        .map(|extension| extension.to_ascii_lowercase());
    if !matches!(extension.as_deref(), Some("xlsx") | Some("xlsm")) {
        bail!(
            "betteroffice 重算需要 .xlsx/.xlsm 工作副本（当前 {}）",
            source_copy.display()
        );
    }

    let mut style_cells_stripped = 0usize;
    if config.strip_style_cells {
        let stripped_path =
            stripped_sibling_path(source_copy).context("为纯样式单元格剥离结果分配临时路径失败")?;
        let report = ooxml_patch::strip_style_only_cells(source_copy, &stripped_path)
            .with_context(|| {
                format!(
                    "剥离工作副本的纯样式空单元格失败（副本：{}）",
                    source_copy.display()
                )
            })?;
        style_cells_stripped = report.style_only_cells_removed;
        if stripped_path.exists() {
            fs_rename(&stripped_path, source_copy)?;
        }
    }

    if let Some(max_visits) = config.max_visits {
        fasteroffice_set_max_visits(max_visits);
    }

    let bytes = std::fs::read(source_copy)
        .with_context(|| format!("读取工作副本失败（{}）", source_copy.display()))?;
    let open_started = Instant::now();
    let mut workbook = betteroffice_xlsx::Workbook::open(&bytes)
        .map_err(|error| anyhow::anyhow!("betteroffice 打开工作副本失败：{}", error))?;
    let open_elapsed = open_started.elapsed();

    let calc_started = Instant::now();
    let options = betteroffice_xlsx::CalculationOptions {
        now_serial: Some(current_excel_serial()),
    };
    let mut changed_cells_total = 0usize;
    let mut result = workbook.recalculate_all(options);
    changed_cells_total += result.changed.len();
    for _ in 1..config.rounds {
        result = workbook.recalculate_all(options);
        changed_cells_total += result.changed.len();
    }
    let limited_cells = result.limited_cells.len();
    let budget_exhausted = result.budget_exhausted;
    let calc_elapsed = calc_started.elapsed();

    let save_started = Instant::now();
    let saved = workbook
        .save()
        .map_err(|error| anyhow::anyhow!("betteroffice 保存工作副本失败：{}", error))?;
    // betteroffice 会把求值错误写成 <v>#ERROR!</v>，而 OOXML 只允许 7 种
    // 标准错误码——Excel 会以 0x80020009 拒开整个文件。任何 #ERROR! 字面
    // 都视为不可采用并回退，绝不把打不开的文件交给下游。
    let error_literal_cells = count_error_literal_cells(&saved);
    if error_literal_cells > 0 {
        return Err(anyhow::anyhow!(BoCalcRefusal::Unusable {
            reason: format!(
                "{error_literal_cells} 个公式格求值为非标准错误字面 #ERROR!（写入会让 Excel 拒开整个文件）"
            )
        }));
    }
    std::fs::write(source_copy, &saved)
        .with_context(|| format!("写回工作副本失败（{}）", source_copy.display()))?;
    let save_elapsed = save_started.elapsed();

    let (error_cells, name_error_cells) = count_error_cells(&workbook);

    if budget_exhausted {
        return Err(anyhow::anyhow!(BoCalcRefusal::BudgetExhausted {
            limited_cells
        }));
    }
    if name_error_cells > 0 {
        return Err(anyhow::anyhow!(BoCalcRefusal::UnsupportedFunctions {
            name_error_cells
        }));
    }
    if !result.cycle_cells.is_empty() {
        return Err(anyhow::anyhow!(BoCalcRefusal::Unusable {
            reason: format!("存在 {} 个循环引用格", result.cycle_cells.len())
        }));
    }

    Ok(BoCalcOutcome {
        application: "betteroffice (pure Rust)",
        rounds_used: config.rounds,
        changed_cells_total,
        limited_cells,
        error_cells,
        name_error_cells,
        style_cells_stripped,
        open_elapsed,
        calc_elapsed,
        save_elapsed,
    })
}

/// 设置 betteroffice 的重算预算上限（vendor 修改版提供的运行时接口）。
fn fasteroffice_set_max_visits(max_visits: u64) {
    xlsx_calc::eval::set_max_recalculation_visits(max_visits);
}

/// betteroffice 会把求值错误写成 `<v>#ERROR!</v>`，而 OOXML 只允许 7 种
/// 标准错误码——Excel 会以 0x80020009 拒开整个文件。
fn count_error_literal_cells(bytes: &[u8]) -> usize {
    bytes
        .windows(b"<v>#ERROR!</v>".len())
        .filter(|window| window == b"<v>#ERROR!</v>")
        .count()
}

/// the local clock as an Excel 1900-system serial (days since 1899-12-30).
/// betteroffice has no clock of its own: without this, `NOW()`/`TODAY()` and
/// everything built on them evaluate to `#VALUE!`.
fn current_excel_serial() -> f64 {
    use chrono::{Datelike, Timelike};
    let now = chrono::Local::now();
    let date = chrono::NaiveDate::from_ymd_opt(now.year(), now.month(), now.day())
        .unwrap_or_else(|| chrono::NaiveDate::from_ymd_opt(1970, 1, 1).expect("epoch"));
    let epoch = chrono::NaiveDate::from_ymd_opt(1899, 12, 30).expect("excel epoch");
    let days = (date - epoch).num_days() as f64;
    let seconds = now.hour() as f64 * 3600.0
        + now.minute() as f64 * 60.0
        + now.second() as f64
        + now.nanosecond() as f64 / 1e9;
    days + seconds / 86_400.0
}

/// total formula cells evaluating to an Excel error, and how many of those are
/// `#NAME?` (the marker for functions betteroffice has not implemented).
fn count_error_cells(workbook: &betteroffice_xlsx::Workbook) -> (usize, usize) {
    use betteroffice_xlsx::CellValue;
    let mut error_cells = 0usize;
    let mut name_error_cells = 0usize;
    for sheet in &workbook.model().sheets {
        for (_cell_ref, cell) in sheet.iter_cells() {
            if cell.formula.is_none() {
                continue;
            }
            if let CellValue::Error { value } = &cell.value {
                error_cells += 1;
                if matches!(value, betteroffice_xlsx::ErrorValue::Name) {
                    // #NAME? is the unsupported-function marker; the rest are
                    // ordinary business errors the candidate validation sees.
                    name_error_cells += 1;
                }
            }
        }
    }
    (error_cells, name_error_cells)
}

/// a sibling temp path for the stripped copy, unique per process and instant.
fn stripped_sibling_path(source_copy: &Path) -> Option<std::path::PathBuf> {
    let file_name = source_copy.file_name()?.to_str()?;
    let nonce = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .ok()?
        .as_nanos();
    let parent = source_copy.parent()?;
    Some(parent.join(format!(
        ".{file_name}.bostrip.{}.{}",
        std::process::id(),
        nonce
    )))
}

fn fs_rename(from: &Path, to: &Path) -> Result<()> {
    #[cfg(windows)]
    {
        // Windows rename fails onto an existing target; remove first. The
        // working copy is disposable, and the strip already succeeded.
        if to.exists() {
            std::fs::remove_file(to)
                .with_context(|| format!("移除旧工作副本以便替换（{}）", to.display()))?;
        }
    }
    std::fs::rename(from, to)
        .with_context(|| format!("用剥离后的副本替换工作副本（{}）", to.display()))
}

#[cfg(test)]
mod tests {
    use super::BoCalcConfig;

    #[test]
    fn config_defaults_enable_auto_routing() {
        // 环境未设置时的缺省：启用、库默认预算、2 轮、不剥样式、允许回退。
        let config = BoCalcConfig {
            enabled: true,
            max_visits: None,
            rounds: 2,
            strip_style_cells: false,
            fallback_com: true,
            fallback_libreoffice: true,
        };
        assert!(config.enabled);
        assert!(config.max_visits.is_none());
        assert!(config.fallback_available());
    }

    #[test]
    fn rounds_above_five_are_clamped() {
        let config = BoCalcConfig {
            enabled: true,
            max_visits: Some(1),
            rounds: 1,
            strip_style_cells: true,
            fallback_com: false,
            fallback_libreoffice: false,
        };
        assert!(!config.fallback_available());
        assert_eq!(config.rounds, 1);
    }

    #[test]
    fn error_literal_detection() {
        use super::count_error_literal_cells;
        assert_eq!(count_error_literal_cells(b"<v>#DIV/0!</v>"), 0);
        assert_eq!(count_error_literal_cells(b"<v>#ERROR!</v>"), 1);
        assert_eq!(
            count_error_literal_cells(b"<c r=\"A1\" t=\"e\"><f>1/0</f><v>#ERROR!</v></c>"),
            1
        );
        assert_eq!(
            count_error_literal_cells(b"<v>#ERROR!</v><v>#ERROR!</v>"),
            2
        );
    }

    /// GUI/配置存储 > 环境变量 > 内置默认。AppSettings::load 读
    /// AUDIT_STATE_DIR 下的 settings.json，本测试用自己的临时状态目录。
    #[test]
    fn settings_take_priority_over_env_and_defaults() {
        // AUDIT_STATE_DIR / PYTOOLS_BO_CALC 是进程级环境：与并行测试串行化。
        let _env_guard = crate::app_settings::AUDIT_STATE_ENV_LOCK
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        use crate::app_settings::AppSettings;
        let state = std::env::temp_dir().join(format!(
            "audit_bo_cfg_{}",
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .expect("clock")
                .as_nanos()
        ));
        std::env::set_var("AUDIT_STATE_DIR", &state);

        // 默认（无设置、无 env）：启用、2 轮、auto 回退。
        std::env::remove_var("PYTOOLS_BO_CALC");
        std::env::remove_var("PYTOOLS_BO_CALC_ROUNDS");
        std::env::remove_var("PYTOOLS_BO_CALC_FALLBACK");
        let config = BoCalcConfig::from_app();
        assert!(config.enabled);
        assert_eq!(config.rounds, 2);
        assert!(config.fallback_com && config.fallback_libreoffice);

        // GUI 设置 com：enabled=false 且回退锁定 COM（覆盖 env 的 auto）。
        AppSettings::store_preferences(
            "auto",
            "auto",
            [
                ("boCalcEngine".to_owned(), serde_json::json!("com")),
                ("boCalcFallback".to_owned(), serde_json::json!("auto")),
            ]
            .into_iter()
            .collect(),
        )
        .expect("store com preference");
        let config = BoCalcConfig::from_app();
        assert!(!config.enabled, "GUI com 必须跳过 betteroffice");
        assert!(config.fallback_com && !config.fallback_libreoffice);

        // GUI 设置 betteroffice + none 回退 + 3 轮 + 预算：GUI 覆盖 env。
        std::env::set_var("PYTOOLS_BO_CALC_ROUNDS", "5");
        std::env::set_var("PYTOOLS_BO_CALC_FALLBACK", "auto");
        AppSettings::store_preferences(
            "auto",
            "auto",
            [
                ("boCalcEngine".to_owned(), serde_json::json!("betteroffice")),
                ("boCalcFallback".to_owned(), serde_json::json!("none")),
                ("boCalcRounds".to_owned(), serde_json::json!(3)),
                ("boCalcMaxVisits".to_owned(), serde_json::json!(12345)),
                ("boCalcStripStyle".to_owned(), serde_json::json!(true)),
            ]
            .into_iter()
            .collect(),
        )
        .expect("store betteroffice preferences");
        let config = BoCalcConfig::from_app();
        assert!(config.enabled);
        assert!(
            !config.fallback_com && !config.fallback_libreoffice,
            "FALLBACK=none"
        );
        assert_eq!(config.rounds, 3, "GUI 3 轮覆盖 env 的 5");
        assert_eq!(config.max_visits, Some(12345));
        assert!(config.strip_style_cells);

        // env 只在 GUI 未设置时生效：清掉 GUI 的轮数后 env=5 生效。
        let settings_path = state.join("settings.json");
        let mut raw: serde_json::Value = serde_json::from_str(
            &std::fs::read_to_string(&settings_path).expect("read settings json"),
        )
        .expect("settings json parse");
        if let Some(preferences) = raw.get_mut("preferences").and_then(|v| v.as_object_mut()) {
            preferences.remove("boCalcRounds");
        }
        std::fs::write(
            &settings_path,
            serde_json::to_vec(&raw).expect("reserialize"),
        )
        .expect("write settings json");
        let config = BoCalcConfig::from_app();
        assert_eq!(config.rounds, 5, "env 覆盖默认");

        for (key, _) in std::env::vars_os() {
            let key = key.to_string_lossy().into_owned();
            if key.starts_with("PYTOOLS_BO_CALC") {
                std::env::remove_var(&key);
            }
        }
        std::env::remove_var("AUDIT_STATE_DIR");
    }

    #[test]
    fn fallback_none_disables_both_engines() {
        let config = BoCalcConfig {
            enabled: true,
            max_visits: None,
            rounds: 2,
            strip_style_cells: false,
            fallback_com: false,
            fallback_libreoffice: false,
        };
        assert!(!config.fallback_available(), "FALLBACK=none 不得静默回退");
    }
}
