//! Persistent user-facing engine selection ("高级设置").
//!
//! Two groups of options are exposed in the GUI and persisted here:
//! - 计算引擎 (`calc_engine`): which external calculation engine performs the
//!   S1-F01 formula recalculation. `auto` keeps the platform default order,
//!   `excel`/`wps` force a native COM application on Windows. LibreOffice is
//!   called directly on Windows and through the existing sidecar on Linux.
//! - 修改引擎 (`edit_engine`): S1-F01 defaults to OOXML Patch, then the
//!   available isolated openpyxl worker, then Umya. Explicit choices force one
//!   backend. S2/S3 keep their existing routes. Every output is validated.
//!
//! Settings are stored as JSON under the same state base directory the S1
//! recovery store uses (`AUDIT_STATE_DIR` or the per-OS user state location),
//! so test and acceptance runs stay isolated with the same env var. Missing
//! or unreadable files fall back to defaults; saving is atomic.

use anyhow::{ensure, Context, Result};
use excel_editor::OpenpyxlWorkerConfig;
use serde::{Deserialize, Serialize};
use std::{collections::BTreeMap, fs, path::PathBuf};

/// AUDIT_STATE_DIR is process-global, so tests that override it or persist
/// preferences must share this lock across modules.
#[cfg(test)]
pub(crate) static AUDIT_STATE_ENV_LOCK: std::sync::Mutex<()> = std::sync::Mutex::new(());

pub const CALC_ENGINE_VALUES: [&str; 4] = ["auto", "excel", "wps", "libreoffice"];
pub const EDIT_ENGINE_VALUES: [&str; 4] = ["auto", "ooxml", "openpyxl", "umya"];
pub const S1_CONDITIONAL_FORMAT_SOURCE_VALUES: [&str; 2] = ["source", "audit_copy"];
/// 条件格式判定方式：PYTHON=OOXML 规则求值（默认，与源生产一致）；
/// COM_DISPLAY=Excel/WPS 实际显示填充作为触发事实（渲染基线）；
/// COM_EVALUATE（公式引擎逐条求值）本产品尚未实现，选择即明确报错。
pub const CONDITIONAL_FORMAT_EVALUATOR_VALUES: [&str; 3] =
    ["PYTHON", "COM_EVALUATE", "COM_DISPLAY"];
/// 大集中核对差异容差默认值（元）：100 元 = 0.01 万元，旧 Audit 同口径。
pub const CENTRAL_DIFF_TOLERANCE_DEFAULT_YUAN: f64 = 100.0;

pub const CALC_ENGINE_LABELS: [(&str, &str); 4] = [
    (
        "auto",
        "自动（Windows=Excel COM，无 Excel 时尝试 WPS；Linux=LibreOffice）",
    ),
    ("excel", "强制 Microsoft Excel（COM）"),
    ("wps", "强制 WPS 表格（COM）"),
    ("libreoffice", "强制 LibreOffice 无头"),
];

pub const EDIT_ENGINE_LABELS: [(&str, &str); 4] = [
    ("auto", "自动：S1-F01 OOXML → 可用的 openpyxl → Umya"),
    ("ooxml", "仅 OOXML Patch（仅 S1-F01）"),
    ("openpyxl", "仅 openpyxl worker（仅 S1-F01，输出仍校验）"),
    ("umya", "仅 Umya（仅 S1-F01）"),
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
            other => anyhow::bail!("计算引擎必须是 auto、excel、wps 或 libreoffice：{other}"),
        }
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum EditEngine {
    /// Legacy persisted `umya` meant the old default route, which already
    /// tried OOXML first. Migrate it to Auto when loading existing settings.
    #[serde(alias = "umya")]
    Auto,
    Ooxml,
    Openpyxl,
    /// The explicit Umya option is serialized distinctly so old `umya`
    /// settings can retain their former automatic behavior.
    #[serde(rename = "umya_only")]
    UmYa,
}

impl EditEngine {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Auto => "auto",
            Self::Ooxml => "ooxml",
            Self::Openpyxl => "openpyxl",
            Self::UmYa => "umya",
        }
    }

    pub fn label(self) -> &'static str {
        EDIT_ENGINE_LABELS
            .iter()
            .find(|(value, _)| *value == self.as_str())
            .map(|(_, label)| *label)
            .unwrap_or("自动")
    }

    pub fn parse(value: &str) -> Result<Self> {
        match value {
            "auto" => Ok(Self::Auto),
            "ooxml" => Ok(Self::Ooxml),
            "openpyxl" => Ok(Self::Openpyxl),
            "umya" => Ok(Self::UmYa),
            other => anyhow::bail!("修改引擎必须是 auto、ooxml、openpyxl 或 umya：{other}"),
        }
    }
}

#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
pub struct AppSettings {
    pub calc_engine: CalcEngine,
    pub edit_engine: EditEngine,
    /// 用户级偏好与最近选择（界面偏好、路径记忆）。按
    /// [`PREFERENCE_SPECS`] 白名单读写：未知键保留（前向兼容）、
    /// 类型不符的键在应用回状态时跳过；失效路径不会覆盖有效启动默认值。
    #[serde(default)]
    pub preferences: BTreeMap<String, serde_json::Value>,
}

impl Default for AppSettings {
    fn default() -> Self {
        Self {
            calc_engine: CalcEngine::Auto,
            edit_engine: EditEngine::Auto,
            preferences: BTreeMap::new(),
        }
    }
}

/// 用户偏好键的白名单与校验类别。路径类字段在启动应用回状态时按
/// 存在性校验：目录必须 is_dir、文件必须 is_file；记忆路径失效时保留有效的
/// 启动默认路径，不影响其他字段与启动。
pub enum PreferenceKind {
    Directory,
    File,
    Boolean,
    Enumeration(&'static [&'static str]),
    IntegerRange {
        min: i64,
        max: i64,
    },
    /// 手动追加待处理工作簿的绝对路径列表（字符串数组）。
    FilePathList,
    /// 浮点数值（如大集中核对容差，单位元）。
    FloatRange {
        min: f64,
        max: f64,
    },
}

pub struct PreferenceSpec {
    pub key: &'static str,
    pub kind: PreferenceKind,
}

macro_rules! preference_specs {
    ($($key:literal => $kind:expr),* $(,)?) => {
        &[$(PreferenceSpec { key: $key, kind: $kind }),*]
    };
}

pub const PREFERENCE_SPECS: &[PreferenceSpec] = preference_specs! {
    "input" => PreferenceKind::Directory,
    "templateDir" => PreferenceKind::Directory,
    "template" => PreferenceKind::File,
    "external" => PreferenceKind::File,
    "output" => PreferenceKind::Directory,
    "outputPinned" => PreferenceKind::Boolean,
    "historyConfig" => PreferenceKind::File,
    "confirmBeforeRun" => PreferenceKind::Boolean,
    "showRunDetailLogs" => PreferenceKind::Boolean,
    "showPerformanceDiagnostics" => PreferenceKind::Boolean,
    "exportRunLogs" => PreferenceKind::Boolean,
    "s1ConditionalFormatSource" => {
        PreferenceKind::Enumeration(&S1_CONDITIONAL_FORMAT_SOURCE_VALUES)
    },
    "conditionalFormatIncludeBlankCells" => PreferenceKind::Boolean,
    "conditionalFormatEvaluator" => {
        PreferenceKind::Enumeration(&CONDITIONAL_FORMAT_EVALUATOR_VALUES)
    },
    "extraFiles" => PreferenceKind::FilePathList,
    "recursiveDepth" => PreferenceKind::IntegerRange { min: -1, max: 10 },
    "uiTheme" => PreferenceKind::Enumeration(&["日间", "日间纯白", "夜间"]),
    "filePickerMode" => {
        PreferenceKind::Enumeration(&["自动", "系统原生", "浏览器内置"])
    },
    "uiAccent" => PreferenceKind::Enumeration(&[
        "默认蓝", "青", "绿", "紫", "金", "橙", "玫", "靛",
    ]),
    "centralCommonConfig" => PreferenceKind::File,
    "centralFormsConfig" => PreferenceKind::File,
    "centralCompareFile" => PreferenceKind::File,
    "centralExplainFile" => PreferenceKind::File,
    "centralComparisonConfig" => PreferenceKind::File,
    "centralCrossConfig" => PreferenceKind::File,
    "centralCrossCur" => PreferenceKind::File,
    "centralCur" => PreferenceKind::File,
    "centralPre" => PreferenceKind::File,
    "centralOutput" => PreferenceKind::Directory,
    "centralFormOutput" => PreferenceKind::Directory,
    "pcCurDir" => PreferenceKind::Directory,
    "pcPreDir" => PreferenceKind::Directory,
    "pcCentral" => PreferenceKind::File,
    "pcConfig" => PreferenceKind::File,
    "pcOutput" => PreferenceKind::Directory,
    "centralTolerance" => PreferenceKind::FloatRange { min: 0.0, max: 1_000_000.0 },
    "centralTrimIndicatorNames" => PreferenceKind::Boolean,
    "centralAutoOpenResult" => PreferenceKind::Boolean,
    "boCalcEngine" => {
        PreferenceKind::Enumeration(&["betteroffice", "com", "off"])
    },
    "boCalcMaxVisits" => PreferenceKind::IntegerRange { min: 0, max: 1_000_000_000 },
    "boCalcRounds" => PreferenceKind::IntegerRange { min: 1, max: 5 },
    "boCalcStripStyle" => PreferenceKind::Boolean,
    "boCalcFallback" => {
        PreferenceKind::Enumeration(&["auto", "com", "libreoffice", "none"])
    },
    "recalcFormulaGuard" => PreferenceKind::Boolean,
};

pub fn preference_spec(key: &str) -> Option<&'static PreferenceSpec> {
    PREFERENCE_SPECS.iter().find(|spec| spec.key == key)
}

/// 校验一个偏好值是否符合其键的类别；返回标准化后的值。
pub fn validate_preference(key: &str, value: &serde_json::Value) -> Result<serde_json::Value> {
    let spec = preference_spec(key).with_context(|| format!("未知的用户设置键：{key}"))?;
    match &spec.kind {
        PreferenceKind::Boolean => {
            ensure!(value.is_boolean(), "用户设置 {key} 必须是布尔值");
            Ok(value.clone())
        }
        PreferenceKind::Directory | PreferenceKind::File => {
            let text = value
                .as_str()
                .with_context(|| format!("用户设置 {key} 必须是字符串路径"))?
                .trim()
                .to_owned();
            Ok(serde_json::Value::String(text))
        }
        PreferenceKind::Enumeration(allowed) => {
            let text = value
                .as_str()
                .with_context(|| format!("用户设置 {key} 必须是字符串"))?;
            ensure!(
                allowed.contains(&text),
                "用户设置 {key} 必须是以下之一：{}",
                allowed.join("、")
            );
            Ok(serde_json::Value::String(text.to_owned()))
        }
        PreferenceKind::IntegerRange { min, max } => {
            let number = value
                .as_i64()
                .with_context(|| format!("用户设置 {key} 必须是整数"))?;
            ensure!(
                (*min..=*max).contains(&number),
                "用户设置 {key} 必须在 {min} 到 {max} 之间"
            );
            Ok(serde_json::Value::Number(number.into()))
        }
        PreferenceKind::FilePathList => {
            let items = value
                .as_array()
                .with_context(|| format!("用户设置 {key} 必须是路径数组"))?;
            let paths = items
                .iter()
                .map(|item| {
                    item.as_str()
                        .map(|text| text.trim().to_owned())
                        .filter(|text| !text.is_empty())
                        .with_context(|| format!("用户设置 {key} 的元素必须是路径字符串"))
                })
                .collect::<Result<Vec<_>>>()?;
            Ok(serde_json::Value::Array(
                paths.into_iter().map(serde_json::Value::String).collect(),
            ))
        }
        PreferenceKind::FloatRange { min, max } => {
            let number = value
                .as_f64()
                .with_context(|| format!("用户设置 {key} 必须是数值"))?;
            ensure!(
                (*min..=*max).contains(&number),
                "用户设置 {key} 必须在 {min} 到 {max} 之间"
            );
            Ok(serde_json::Number::from_f64(number)
                .map(serde_json::Value::Number)
                .with_context(|| format!("用户设置 {key} 不是有限数值"))?)
        }
    }
}

impl AppSettings {
    /// Returns the validated S1-F01 conditional-format value source. Missing
    /// or stale settings default to the repaired recalculated audit copy so
    /// 扫描区域提取、Rust 求值与 COM 渲染都落在同一份副本上，便于对拍；
    /// 原始源文件模式仍可在高级设置中显式选择用于对照。
    pub fn s1_conditional_format_source(&self) -> &str {
        self.preferences
            .get("s1ConditionalFormatSource")
            .and_then(serde_json::Value::as_str)
            .filter(|value| S1_CONDITIONAL_FORMAT_SOURCE_VALUES.contains(value))
            .unwrap_or("audit_copy")
    }

    fn settings_path() -> Result<PathBuf> {
        Ok(state_base_dir()?.join("settings.json"))
    }

    /// 条件格式空白格命中是否计入 S1-F01 结果；默认开启，与 Excel/WPS COM
    /// DisplayFormat 的实际渲染口径一致（空白格实际变红同样算触发）。
    pub fn conditional_format_include_blank_cells(&self) -> bool {
        self.preferences
            .get("conditionalFormatIncludeBlankCells")
            .and_then(serde_json::Value::as_bool)
            .unwrap_or(true)
    }

    /// 手动追加进待处理清单的源目录外工作簿（跨会话保留）。
    pub fn extra_files(&self) -> Vec<String> {
        self.preferences
            .get("extraFiles")
            .and_then(serde_json::Value::as_array)
            .map(|items| {
                items
                    .iter()
                    .filter_map(serde_json::Value::as_str)
                    .map(str::to_owned)
                    .collect()
            })
            .unwrap_or_default()
    }

    /// 条件格式判定方式；缺省 PYTHON（OOXML 规则求值），与源生产默认一致。
    pub fn conditional_format_evaluator(&self) -> &str {
        self.preferences
            .get("conditionalFormatEvaluator")
            .and_then(serde_json::Value::as_str)
            .filter(|value| CONDITIONAL_FORMAT_EVALUATOR_VALUES.contains(value))
            .unwrap_or("PYTHON")
    }

    /// 校验计算引擎：betteroffice（默认，纯 Rust 优先）｜com（直接用
    /// Excel/WPS）｜off（完全回到 COM/LibreOffice 旧路由）。
    /// `None` = GUI 未设置（此时环境变量可以生效）。
    pub fn bo_calc_engine(&self) -> Option<&str> {
        self.preferences
            .get("boCalcEngine")
            .and_then(serde_json::Value::as_str)
            .filter(|value| ["betteroffice", "com", "off"].contains(value))
    }

    /// 校验计算预算上限；0/未设置 = 库默认（10,000,000）。
    pub fn bo_calc_max_visits(&self) -> Option<u64> {
        self.preferences
            .get("boCalcMaxVisits")
            .and_then(serde_json::Value::as_u64)
            .filter(|value| *value > 0)
    }

    /// 校验计算轮数（1~5，默认 2；第 2 轮起跨表动态引用才收敛）。
    pub fn bo_calc_rounds(&self) -> Option<u32> {
        self.preferences
            .get("boCalcRounds")
            .and_then(serde_json::Value::as_u64)
            .filter(|value| (1..=5).contains(value))
            .map(|value| value as u32)
    }

    /// 校验计算剥离样式空格（默认否）。
    pub fn bo_calc_strip_style(&self) -> Option<bool> {
        self.preferences
            .get("boCalcStripStyle")
            .and_then(serde_json::Value::as_bool)
    }

    /// 校验计算回退策略（auto 默认；none=不回退直接报错）。
    pub fn bo_calc_fallback(&self) -> Option<&str> {
        self.preferences
            .get("boCalcFallback")
            .and_then(serde_json::Value::as_str)
            .filter(|value| ["auto", "com", "libreoffice", "none"].contains(value))
    }

    /// 重算公式文本一致性守卫（默认开启）。
    ///
    /// 开启时逐格比对重算前后的公式文本；关闭只跳过这一层比对，
    /// **公式数量、缓存节点完整性、结果可读性三层检查始终保留**——
    /// 那是防"引擎丢公式/丢缓存"的底线。
    pub fn recalc_formula_guard_enabled(&self) -> bool {
        self.preferences
            .get("recalcFormulaGuard")
            .and_then(serde_json::Value::as_bool)
            .unwrap_or(true)
    }

    /// 大集中核对差异容差（元）。保存时已做范围校验；手工改坏的设置文件
    /// 回落默认 100 元（=0.01 万元），与旧 Audit 设置中心口径一致。
    pub fn central_diff_tolerance_yuan(&self) -> f64 {
        self.preferences
            .get("centralTolerance")
            .and_then(serde_json::Value::as_f64)
            .filter(|value| (0.0..=1_000_000.0).contains(value))
            .unwrap_or(CENTRAL_DIFF_TOLERANCE_DEFAULT_YUAN)
    }

    /// 比较结果工作簿的指标名称显示列是否去首尾空格（不影响匹配逻辑）。
    pub fn central_trim_indicator_names(&self) -> bool {
        self.preferences
            .get("centralTrimIndicatorNames")
            .and_then(serde_json::Value::as_bool)
            .unwrap_or(false)
    }

    /// 大集中执行比较完成后是否用系统默认程序自动打开结果工作簿。
    pub fn central_auto_open_result(&self) -> bool {
        self.preferences
            .get("centralAutoOpenResult")
            .and_then(serde_json::Value::as_bool)
            .unwrap_or(false)
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
            preferences: Self::load().preferences,
        };
        Self::write(settings)
    }

    /// 校验并合并一批用户偏好（仅白名单键），与引擎选择一并原子写入。
    pub fn store_preferences(
        calc_engine: &str,
        edit_engine: &str,
        preferences: BTreeMap<String, serde_json::Value>,
    ) -> Result<Self> {
        let mut settings = Self {
            calc_engine: CalcEngine::parse(calc_engine)?,
            edit_engine: EditEngine::parse(edit_engine)?,
            preferences: Self::load().preferences,
        };
        for (key, value) in &preferences {
            let normalized = validate_preference(key, value)?;
            settings.preferences.insert(key.clone(), normalized);
        }
        Self::write(settings)
    }

    fn write(settings: AppSettings) -> Result<AppSettings> {
        let path = Self::settings_path()?;
        if let Some(parent) = path.parent() {
            fs::create_dir_all(parent)
                .with_context(|| format!("创建设置目录：{}", parent.display()))?;
        }
        // 唯一临时名：并发调用方（多线程测试/多请求）共用同一目录时，
        // 固定 tmp 名会让后到的 rename 找不到已被先到者改名的文件。
        let nonce = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .map(|value| value.as_nanos())
            .unwrap_or_default();
        let temporary = path.with_extension(format!("json.tmp-{}-{nonce}", std::process::id()));
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

/// 将持久化的偏好应用回 GUI 状态：
/// - 布尔/枚举：类型与取值合法才覆盖；
/// - 目录/文件：有效的记忆路径覆盖启动默认值；失效记忆不会清空已经存在的
///   启动默认路径；两者都无效时才清空该字段；
/// - 未知键跳过（前向兼容）。
/// 发行包整体移动后，把失效的配置路径按文件名重新定位到
/// `<exe 目录>/config-real/`。找不到时返回 None，由调用方保留启动默认。
fn relocate_bundled_config(trimmed: &str) -> Option<String> {
    let file_name = std::path::Path::new(trimmed).file_name()?;
    let executable = std::env::current_exe().ok()?;
    let directory = executable.parent()?.join("config-real").join(file_name);
    directory
        .is_file()
        .then(|| directory.to_string_lossy().into_owned())
}

pub fn apply_preferences_to_state(state: &mut serde_json::Value) {
    let preferences = AppSettings::load().preferences;
    let Some(state_object) = state.as_object_mut() else {
        return;
    };
    for (key, value) in &preferences {
        let Some(spec) = preference_spec(key) else {
            continue;
        };
        match &spec.kind {
            PreferenceKind::Boolean => {
                if value.is_boolean() {
                    state_object.insert(key.clone(), value.clone());
                }
            }
            PreferenceKind::Directory | PreferenceKind::File => {
                let Some(text) = value.as_str() else {
                    continue;
                };
                let trimmed = text.trim();
                let path = PathBuf::from(trimmed);
                let valid = !trimmed.is_empty()
                    && match &spec.kind {
                        PreferenceKind::Directory => path.is_dir(),
                        PreferenceKind::File => path.is_file(),
                        _ => false,
                    };
                if valid {
                    state_object.insert(key.clone(), serde_json::Value::String(trimmed.to_owned()));
                } else {
                    // 包被整体移动后，旧绝对路径失效：按「同文件名」在
                    // 可执行文件同级的 config-real/ 下重新定位，用户不必重选。
                    if let Some(relocated) = relocate_bundled_config(trimmed) {
                        state_object
                            .insert(key.clone(), serde_json::Value::String(relocated));
                        continue;
                    }
                    let default_is_valid = state_object
                        .get(key)
                        .and_then(serde_json::Value::as_str)
                        .map(str::trim)
                        .is_some_and(|default| {
                            let path = PathBuf::from(default);
                            match &spec.kind {
                                PreferenceKind::Directory => path.is_dir(),
                                PreferenceKind::File => path.is_file(),
                                _ => false,
                            }
                        });
                    if !default_is_valid {
                        state_object.insert(key.clone(), serde_json::Value::String(String::new()));
                    }
                }
            }
            PreferenceKind::Enumeration(allowed) => {
                if let Some(text) = value.as_str() {
                    if allowed.contains(&text) {
                        state_object
                            .insert(key.clone(), serde_json::Value::String(text.to_owned()));
                    }
                }
            }
            PreferenceKind::IntegerRange { min, max } => {
                if value
                    .as_i64()
                    .is_some_and(|number| (*min..=*max).contains(&number))
                {
                    state_object.insert(key.clone(), value.clone());
                }
            }
            PreferenceKind::FilePathList => {
                if value.is_array() {
                    state_object.insert(key.clone(), value.clone());
                }
            }
            PreferenceKind::FloatRange { min, max } => {
                if value
                    .as_f64()
                    .is_some_and(|number| (*min..=*max).contains(&number))
                {
                    state_object.insert(key.clone(), value.clone());
                }
            }
        }
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
        if let Some(path) = std::env::var_os("XDG_STATE_HOME").filter(|v| !v.is_empty()) {
            return Ok(PathBuf::from(path).join("open-data-audit"));
        }
        let home = std::env::var_os("HOME")
            .map(PathBuf::from)
            .context("用户状态目录不可用：XDG_STATE_HOME 和 HOME 均未设置")?;
        Ok(home.join(".local/state/open-data-audit"))
    }
}

/// Whether the historical openpyxl production capability evidence passed.
pub fn openpyxl_evidence_status() -> (bool, String) {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("capability_evidence/openpyxl_windows_x86_64_20261029.json");
    let Ok(text) = fs::read_to_string(&path) else {
        return (
            false,
            "尚未生成 openpyxl 能力证据（capability_evidence/openpyxl_windows_x86_64_20261029.json）"
                .to_owned(),
        );
    };
    let Ok(document) = serde_json::from_str::<serde_json::Value>(&text) else {
        return (false, "openpyxl 能力证据文件无法解析".to_owned());
    };
    match document
        .get("available")
        .and_then(serde_json::Value::as_bool)
    {
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

/// Openpyxl remains selectable for an explicit S1-F01 trial without passing
/// historical production evidence, but only when a Python interpreter and
/// worker script are actually available. Output still passes the normal
/// validator before publication.
pub fn openpyxl_selection_status() -> (bool, String) {
    static STATUS: std::sync::OnceLock<(bool, String)> = std::sync::OnceLock::new();
    STATUS.get_or_init(detect_openpyxl_selection_status).clone()
}

fn detect_openpyxl_selection_status() -> (bool, String) {
    let config = std::env::current_exe()
        .ok()
        .and_then(|binary| OpenpyxlWorkerConfig::for_product_binary(binary, "open-data-audit"));
    let Some(config) = config else {
        return (
            false,
            "无法定位可用的 Python 解释器和 openpyxl worker".to_owned(),
        );
    };
    if !config.is_available() {
        return (
            false,
            format!(
                "openpyxl 备选不可用：Python={}；worker={}。Windows 需在 PATH 中安装 Python，Linux 需部署随包隔离运行时。",
                config.python_executable.display(),
                config.worker_script.display()
            ),
        );
    }
    let import_check = std::process::Command::new(&config.python_executable)
        .args(["-I", "-c", "import openpyxl"])
        .stdin(std::process::Stdio::null())
        .stdout(std::process::Stdio::null())
        .stderr(std::process::Stdio::null())
        .status();
    match import_check {
        Ok(status) if status.success() => {}
        Ok(_) => {
            return (
                false,
                format!(
                    "openpyxl 备选不可用：Python={} 无法导入 openpyxl；自动模式会跳过此后端",
                    config.python_executable.display()
                ),
            );
        }
        Err(error) => {
            return (
                false,
                format!(
                    "openpyxl 备选不可用：无法运行 Python={} 检查 openpyxl：{error}",
                    config.python_executable.display()
                ),
            );
        }
    }
    let (verified, evidence_detail) = openpyxl_evidence_status();
    let runtime_note = if cfg!(windows) {
        format!(
            "Windows 将调用 PATH 中的 Python（{}）；该 Python 环境还需可导入 openpyxl",
            config.python_executable.display()
        )
    } else {
        "已检测到随包隔离 Python/openpyxl 运行时和 worker".to_owned()
    };
    if verified {
        return (
            true,
            format!("openpyxl 已有可用能力证据；{runtime_note}；输出仍会执行独立校验"),
        );
    }
    (
        true,
        format!(
            "允许显式试跑未验证的 openpyxl 候选；历史验证记录：{evidence_detail}；{runtime_note}。只有通过写后校验的候选才会发布"
        ),
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn defaults_are_auto_for_both_engines() {
        let settings = AppSettings::default();
        assert_eq!(settings.calc_engine.as_str(), "auto");
        assert_eq!(settings.edit_engine.as_str(), "auto");
        assert_eq!(settings.calc_engine.sidecar_engine(), "auto");
    }

    #[test]
    fn new_preference_kinds_validate_and_normalize() {
        // 路径数组：合法值保留、非法元素拒绝。
        let list = validate_preference(
            "extraFiles",
            &serde_json::json!(["C:\\a.xlsx", " C:\\b.xlsm "]),
        )
        .expect("路径数组必须可保存");
        assert_eq!(
            list,
            serde_json::json!(["C:\\a.xlsx", "C:\\b.xlsm"]),
            "元素需去首尾空格"
        );
        assert!(validate_preference("extraFiles", &serde_json::json!(["a", 1])).is_err());
        assert!(validate_preference("extraFiles", &serde_json::json!("不是数组")).is_err());
        // 浮点范围：边界与越界。
        let tolerance =
            validate_preference("centralTolerance", &serde_json::json!(150.5)).expect("容差可保存");
        assert_eq!(tolerance, serde_json::json!(150.5));
        assert!(validate_preference("centralTolerance", &serde_json::json!(-1.0)).is_err());
        assert!(validate_preference("centralTolerance", &serde_json::json!(2_000_000.0)).is_err());
        assert!(validate_preference("centralTolerance", &serde_json::json!("100")).is_err());
        // 布尔键。
        assert!(validate_preference(
            "conditionalFormatIncludeBlankCells",
            &serde_json::json!(true)
        )
        .is_ok());
        assert!(validate_preference(
            "conditionalFormatIncludeBlankCells",
            &serde_json::json!("yes")
        )
        .is_err());
    }

    #[test]
    fn accessors_apply_documented_defaults() {
        let settings = AppSettings::default();
        assert!(settings.conditional_format_include_blank_cells());
        assert_eq!(settings.central_diff_tolerance_yuan(), 100.0);
        assert!(!settings.central_trim_indicator_names());
        assert!(!settings.central_auto_open_result());
        assert!(settings.extra_files().is_empty());
        // 手工改坏的设置文件按口径回落默认，不阻断运行。
        let mut broken = AppSettings::default();
        broken
            .preferences
            .insert("centralTolerance".to_owned(), serde_json::json!(-5.0));
        assert_eq!(broken.central_diff_tolerance_yuan(), 100.0);
        broken.preferences.insert(
            "conditionalFormatIncludeBlankCells".to_owned(),
            serde_json::json!(false),
        );
        assert!(!broken.conditional_format_include_blank_cells());
    }

    #[test]
    fn parse_rejects_unknown_values_with_named_reason() {
        assert!(CalcEngine::parse("pywin32").is_err());
        assert!(EditEngine::parse("com").is_err());
        let error = CalcEngine::parse("bogus").unwrap_err().to_string();
        assert!(error.contains("auto、excel、wps 或 libreoffice"));
    }

    #[test]
    fn legacy_umya_setting_migrates_to_the_old_automatic_route() {
        let legacy: AppSettings =
            serde_json::from_str(r#"{"calc_engine":"auto","edit_engine":"umya"}"#)
                .expect("legacy settings remain readable");
        assert_eq!(legacy.edit_engine, EditEngine::Auto);
        assert_eq!(legacy.edit_engine.as_str(), "auto");

        let explicit_umya: AppSettings =
            serde_json::from_str(r#"{"calc_engine":"auto","edit_engine":"umya_only"}"#)
                .expect("new explicit Umya settings remain readable");
        assert_eq!(explicit_umya.edit_engine, EditEngine::UmYa);
        assert_eq!(explicit_umya.edit_engine.as_str(), "umya");
    }

    #[test]
    fn libreoffice_maps_to_the_lo_sidecar_channel() {
        assert_eq!(CalcEngine::LibreOffice.sidecar_engine(), "lo");
        assert_eq!(CalcEngine::Wps.sidecar_engine(), "wps");
        assert_eq!(CalcEngine::Excel.sidecar_engine(), "excel");
    }

    #[test]
    fn store_round_trips_and_rejects_invalid_isolated_by_state_dir() {
        let _settings_env_guard = AUDIT_STATE_ENV_LOCK
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        let nonce = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let dir = std::env::temp_dir().join(format!(
            "audit_app_settings_{}_{}",
            std::process::id(),
            nonce
        ));
        let previous_state_dir = std::env::var("AUDIT_STATE_DIR").ok();
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
        if let Some(previous) = previous_state_dir {
            std::env::set_var("AUDIT_STATE_DIR", previous);
        }
        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn every_edit_engine_choice_is_exposed_and_has_a_stable_label() {
        assert_eq!(EDIT_ENGINE_VALUES, ["auto", "ooxml", "openpyxl", "umya"]);
        for value in EDIT_ENGINE_VALUES {
            let parsed = EditEngine::parse(value).expect("listed engine choice parses");
            assert_eq!(parsed.as_str(), value);
            assert!(!parsed.label().is_empty());
        }
    }

    #[test]
    fn missing_saved_path_preserves_an_existing_startup_default() {
        let _settings_env_guard = AUDIT_STATE_ENV_LOCK
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        let nonce = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let root = std::env::temp_dir().join(format!(
            "audit_app_settings_default_path_{}_{}",
            std::process::id(),
            nonce
        ));
        let state_dir = root.join("state");
        let config_dir = root.join("config-real");
        fs::create_dir_all(&state_dir).expect("create isolated settings directory");
        fs::create_dir_all(&config_dir).expect("create startup config directory");
        let default_config = config_dir.join("3.0大集中通用配置.xlsx");
        fs::write(&default_config, b"available config").expect("write startup config");

        let previous_state_dir = std::env::var("AUDIT_STATE_DIR").ok();
        std::env::set_var("AUDIT_STATE_DIR", &state_dir);
        AppSettings::store_preferences(
            "auto",
            "umya",
            BTreeMap::from([(
                "centralCommonConfig".to_owned(),
                serde_json::Value::String(root.join("removed-config.xlsx").display().to_string()),
            )]),
        )
        .expect("save stale config path");

        let mut state = serde_json::json!({
            "centralCommonConfig": default_config.to_string_lossy()
        });
        apply_preferences_to_state(&mut state);
        assert_eq!(
            state["centralCommonConfig"].as_str(),
            Some(default_config.to_string_lossy().as_ref()),
            "a stale remembered path must not erase the detected packaged default"
        );

        std::env::remove_var("AUDIT_STATE_DIR");
        if let Some(previous) = previous_state_dir {
            std::env::set_var("AUDIT_STATE_DIR", previous);
        }
        fs::remove_dir_all(root).expect("remove settings test directory");
    }
}
