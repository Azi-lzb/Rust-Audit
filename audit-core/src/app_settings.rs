//! Persistent user-facing engine selection ("高级设置").
//!
//! Two groups of options are exposed in the GUI and persisted here:
//! - 计算引擎 (`calc_engine`): which external calculation engine performs the
//!   S1-F01 formula recalculation. `auto` keeps the platform default order,
//!   `excel`/`wps` force a native COM application on Windows. LibreOffice is
//!   called directly on Windows and through the existing sidecar on Linux.
//! - 修改引擎 (`edit_engine`): which Excel-Core backend executes the
//!   S1-F01/S3-F03/S3-F05 edit batches. `umya` is the verified default;
//!   `openpyxl` is the advanced compatibility path and only routes when its
//!   capability evidence is present.
//!
//! Settings are stored as JSON under the same state base directory the S1
//! recovery store uses (`AUDIT_STATE_DIR` or the per-OS user state location),
//! so test and acceptance runs stay isolated with the same env var. Missing
//! or unreadable files fall back to defaults; saving is atomic.

use anyhow::{Context, Result};
use serde::{Deserialize, Serialize};
use std::{fs, path::PathBuf};

pub const CALC_ENGINE_VALUES: [&str; 4] = ["auto", "excel", "wps", "libreoffice"];
pub const EDIT_ENGINE_VALUES: [&str; 2] = ["umya", "openpyxl"];

pub const CALC_ENGINE_LABELS: [(&str, &str); 4] = [
    ("auto", "自动（Windows=Excel COM，无 Excel 时尝试 WPS；Linux=LibreOffice）"),
    ("excel", "强制 Microsoft Excel（COM）"),
    ("wps", "强制 WPS 表格（COM）"),
    ("libreoffice", "强制 LibreOffice 无头"),
];

pub const EDIT_ENGINE_LABELS: [(&str, &str); 2] = [
    ("umya", "Umya（默认，已验证生产路径）"),
    ("openpyxl", "openpyxl worker（高级：兼容路径）"),
];

#[derive(Clone, Copy, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum CalcEngine {
    Auto,
    Excel,
    Wps,
    LibreOffice,
}

impl CalcEngine {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Auto => "auto",
            Self::Excel => "excel",
            Self::Wps => "wps",
            Self::LibreOffice => "libreoffice",
        }
    }

    pub fn label(self) -> &'static str {
        CALC_ENGINE_LABELS
            .iter()
            .find(|(value, _)| *value == self.as_str())
            .map(|(_, label)| *label)
            .unwrap_or("自动")
    }

    /// The sidecar `--engine` value for the S1-F01 recalculation request.
    pub fn sidecar_engine(self) -> &'static str {
        match self {
            Self::Auto => "auto",
            Self::Excel => "excel",
            Self::Wps => "wps",
            Self::LibreOffice => "lo",
        }
    }

    pub fn parse(value: &str) -> Result<Self> {
        match value {
            "auto" => Ok(Self::Auto),
            "excel" => Ok(Self::Excel),
            "wps" => Ok(Self::Wps),
            "libreoffice" => Ok(Self::LibreOffice),
            other => anyhow::bail!(
                "计算引擎必须是 auto、excel、wps 或 libreoffice：{other}"
            ),
        }
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum EditEngine {
    UmYa,
    Openpyxl,
}

impl EditEngine {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::UmYa => "umya",
            Self::Openpyxl => "openpyxl",
        }
    }

    pub fn label(self) -> &'static str {
        EDIT_ENGINE_LABELS
            .iter()
            .find(|(value, _)| *value == self.as_str())
            .map(|(_, label)| *label)
            .unwrap_or("Umya")
    }

    pub fn parse(value: &str) -> Result<Self> {
        match value {
            "umya" => Ok(Self::UmYa),
            "openpyxl" => Ok(Self::Openpyxl),
            other => anyhow::bail!("修改引擎必须是 umya 或 openpyxl：{other}"),
        }
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq, Serialize, Deserialize)]
pub struct AppSettings {
    pub calc_engine: CalcEngine,
    pub edit_engine: EditEngine,
}

impl Default for AppSettings {
    fn default() -> Self {
        Self {
            calc_engine: CalcEngine::Auto,
            edit_engine: EditEngine::UmYa,
        }
    }
}

impl AppSettings {
    fn settings_path() -> Result<PathBuf> {
        Ok(state_base_dir()?.join("settings.json"))
    }

    /// Loads persisted settings; any missing or unreadable file yields the
    /// defaults so a broken settings file cannot block a run.
    pub fn load() -> Self {
        let Ok(path) = Self::settings_path() else {
            return Self::default();
        };
        let Ok(text) = fs::read_to_string(&path) else {
            return Self::default();
        };
        serde_json::from_str::<AppSettings>(&text).unwrap_or_default()
    }

    /// Validates and saves the new selection atomically; the caller surfaces
    /// the error so invalid GUI input never silently reverts.
    pub fn store(calc_engine: &str, edit_engine: &str) -> Result<Self> {
        let settings = Self {
            calc_engine: CalcEngine::parse(calc_engine)?,
            edit_engine: EditEngine::parse(edit_engine)?,
        };
        let path = Self::settings_path()?;
        if let Some(parent) = path.parent() {
            fs::create_dir_all(parent)
                .with_context(|| format!("创建设置目录：{}", parent.display()))?;
        }
        let temporary = path.with_extension("json.tmp");
        fs::write(&temporary, serde_json::to_vec_pretty(&settings)?)
            .with_context(|| format!("写入设置临时文件：{}", temporary.display()))?;
        fs::rename(&temporary, &path)
            .with_context(|| format!("提交设置文件：{}", path.display()))?;
        Ok(settings)
    }

    /// Where the settings file lives; surfaced in logs and the GUI status.
    pub fn settings_path_for_display() -> String {
        Self::settings_path()
            .map(|path| path.display().to_string())
            .unwrap_or_else(|_| "状态目录不可用".to_owned())
    }
}

/// Mirrors the S1 recovery store's base-directory resolution so acceptance
/// runs can isolate both with one environment variable.
fn state_base_dir() -> Result<PathBuf> {
    if let Some(directory) = std::env::var_os("AUDIT_STATE_DIR").filter(|v| !v.is_empty()) {
        return Ok(PathBuf::from(directory));
    }
    #[cfg(windows)]
    {
        let base = std::env::var_os("LOCALAPPDATA")
            .map(PathBuf::from)
            .context("Windows 用户状态目录不可用：LOCALAPPDATA 未设置")?;
        Ok(base.join("open-data-audit"))
    }
    #[cfg(not(windows))]
    {
        if let Some(path) =
            std::env::var_os("XDG_STATE_HOME").filter(|v| !v.is_empty())
        {
            return Ok(PathBuf::from(path).join("open-data-audit"));
        }
        let home = std::env::var_os("HOME")
            .map(PathBuf::from)
            .context("用户状态目录不可用：XDG_STATE_HOME 和 HOME 均未设置")?;
        Ok(home.join(".local/state/open-data-audit"))
    }
}

/// Whether the openpyxl edit-engine capability evidence is present; the GUI
/// uses this to enable or grey out the advanced compatibility option.
pub fn openpyxl_evidence_status() -> (bool, String) {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("capability_evidence/openpyxl_windows_x86_64_20261029.json");
    let Ok(text) = fs::read_to_string(&path) else {
        return (
            false,
            "尚未生成 openpyxl 能力证据（capability_evidence/openpyxl_windows_x86_64_20261029.json）；             该选项置灰，选择将返回错误"
                .to_owned(),
        );
    };
    let Ok(document) = serde_json::from_str::<serde_json::Value>(&text) else {
        return (
            false,
            "openpyxl 能力证据文件无法解析；该选项置灰".to_owned(),
        );
    };
    match document.get("available").and_then(serde_json::Value::as_bool) {
        Some(true) => (true, path.display().to_string()),
        _ => {
            let conclusion = document
                .get("conclusion")
                .and_then(serde_json::Value::as_str)
                .unwrap_or("能力证据未声明可用");
            (false, format!("openpyxl 修改引擎当前不可用：{conclusion}"))
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn defaults_are_auto_and_umya() {
        let settings = AppSettings::default();
        assert_eq!(settings.calc_engine.as_str(), "auto");
        assert_eq!(settings.edit_engine.as_str(), "umya");
        assert_eq!(settings.calc_engine.sidecar_engine(), "auto");
    }

    #[test]
    fn parse_rejects_unknown_values_with_named_reason() {
        assert!(CalcEngine::parse("pywin32").is_err());
        assert!(EditEngine::parse("com").is_err());
        let error = CalcEngine::parse("bogus").unwrap_err().to_string();
        assert!(error.contains("auto、excel、wps 或 libreoffice"));
    }

    #[test]
    fn libreoffice_maps_to_the_lo_sidecar_channel() {
        assert_eq!(CalcEngine::LibreOffice.sidecar_engine(), "lo");
        assert_eq!(CalcEngine::Wps.sidecar_engine(), "wps");
        assert_eq!(CalcEngine::Excel.sidecar_engine(), "excel");
    }

    #[test]
    fn store_round_trips_and_rejects_invalid_isolated_by_state_dir() {
        let nonce = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let dir = std::env::temp_dir().join(format!(
            "audit_app_settings_{}_{}",
            std::process::id(),
            nonce
        ));
        std::env::set_var("AUDIT_STATE_DIR", &dir);
        let saved = AppSettings::store("wps", "umya").expect("valid selection must save");
        assert_eq!(saved.calc_engine.as_str(), "wps");
        let loaded = AppSettings::load();
        assert_eq!(loaded.calc_engine.as_str(), "wps");
        assert_eq!(loaded.edit_engine.as_str(), "umya");
        assert!(AppSettings::store("pywin32", "umya").is_err());
        // A rejected store must not overwrite the previously saved selection.
        assert_eq!(AppSettings::load().calc_engine.as_str(), "wps");
        std::env::remove_var("AUDIT_STATE_DIR");
        std::fs::remove_dir_all(&dir).unwrap();
    }
}
