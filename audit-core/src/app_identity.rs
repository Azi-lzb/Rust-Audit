//! 产品身份：名称、发布版本与产物命名。
//!
//! 与旧 audit `base_audit/app_identity.py` 同口径：
//! - 版本清单 `VERSION_MANIFEST.json` 的 `product.id/name/version` 可信；
//! - 版本格式 `YY.MAJOR.MINOR.PATCH`（年份两位），各段不超过 65535；
//! - 界面标题为「名称 V版本」；英文产品名固定为 `AuditTools`。
//!
//! 发行包在安装根放置 `VERSION_MANIFEST.json`（config_update 也读它）；
//! 单文件分发时回退到编译期嵌入的同一份清单，保证 --version 与界面标题
//! 始终可用。

use std::path::PathBuf;
use std::sync::OnceLock;

use anyhow::{ensure, Context, Result};

/// 英文产品名（发布产物、安装路径、进程名使用）。
pub const PRODUCT_NAME: &str = "AuditTools";
/// 产品资源目录名：`libexec/<PRODUCT_DIR>/openpyxl`、`/opt/<PRODUCT_DIR>`。
/// 只允许小写字母、数字与 `-`（openpyxl worker 解析器约束）。
pub const PRODUCT_DIR: &str = "audittools";
/// 编译期兜底清单：发行包内清单缺失时仍能显示名称与版本。
const EMBEDDED_MANIFEST: &str = include_str!("../../VERSION_MANIFEST.json");

/// 当前产品版本：必须与 `VERSION_MANIFEST.json` 的 `product.version` 一致
/// （由单测强制校验）。clap 的 `--version` 需要 `&'static str`，故用常量。
pub const PRODUCT_VERSION: &str = "26.3.0.0";

const DEFAULT_NAME: &str = "审核工具";

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ProductInfo {
    pub id: String,
    pub name: String,
    pub version: String,
}

impl ProductInfo {
    /// 界面标题：「审核工具 V26.3.0.0」。
    pub fn title(&self) -> String {
        format!("{} V{}", self.name, self.version)
    }
}

/// 校验版本号是否符合 `YY.MAJOR.MINOR.PATCH` 且各段不超过 65535。
pub fn validate_version(version: &str) -> Result<()> {
    let parts: Vec<&str> = version.split('.').collect();
    ensure!(
        parts.len() == 4,
        "产品版本必须为 年份.大版本.中版本.小版本（年份两位）：{version}"
    );
    ensure!(
        parts[0].len() == 2 && parts[0].bytes().all(|b| b.is_ascii_digit()),
        "产品版本第一段必须是两位年份：{version}"
    );
    for part in &parts {
        ensure!(
            !part.is_empty() && part.bytes().all(|b| b.is_ascii_digit()),
            "产品版本各段必须是数字：{version}"
        );
        let value: u64 = part.parse().context("解析产品版本段")?;
        ensure!(value <= 65535, "产品版本各段不能超过 65535：{version}");
    }
    Ok(())
}

/// 发行包安装根：设置 `AUDIT_INSTALL_ROOT`/`AUDIT_CORE_ROOT` 时优先，
/// 否则取可执行文件所在目录（开发目录直接双击 dist exe 时取其上级源码树）。
fn manifest_search_paths() -> Vec<PathBuf> {
    let mut paths = Vec::new();
    for key in ["AUDIT_INSTALL_ROOT", "AUDIT_CORE_ROOT"] {
        if let Some(root) = std::env::var_os(key).filter(|value| !value.is_empty()) {
            paths.push(PathBuf::from(root).join("VERSION_MANIFEST.json"));
        }
    }
    if let Ok(executable) = std::env::current_exe() {
        if let Some(directory) = executable.parent() {
            paths.push(directory.join("VERSION_MANIFEST.json"));
            if let Some(parent) = directory.parent() {
                paths.push(parent.join("VERSION_MANIFEST.json"));
            }
        }
    }
    if let Ok(current) = std::env::current_dir() {
        paths.push(current.join("VERSION_MANIFEST.json"));
        paths.push(current.join("audit-core/VERSION_MANIFEST.json"));
    }
    paths
}

fn parse_manifest(text: &str) -> Result<ProductInfo> {
    let value: serde_json::Value =
        serde_json::from_str(text).context("解析产品版本清单 VERSION_MANIFEST.json")?;
    let product = value
        .get("product")
        .context("版本清单缺少 product 段")?
        .as_object()
        .context("版本清单 product 必须是对象")?;
    let read = |key: &str| -> Result<String> {
        product
            .get(key)
            .and_then(|item| item.as_str())
            .map(str::trim)
            .filter(|text| !text.is_empty())
            .map(str::to_owned)
            .with_context(|| format!("版本清单 product.{key} 必须为非空文字"))
    };
    let info = ProductInfo {
        id: read("id")?,
        name: read("name")?,
        version: read("version")?,
    };
    ensure!(
        info.name.chars().all(|ch| !matches!(
            ch,
            '<' | '>' | ':' | '"' | '/' | '\\' | '|' | '?' | '*' | '\u{0}'..='\u{1f}'
        )) && !info.name.ends_with(['.', ' ']),
        "产品名称含文件名不支持的字符：{}",
        info.name
    );
    validate_version(&info.version)?;
    Ok(info)
}

/// 当前产品信息：优先读安装根清单，缺失/损坏时用编译期嵌入清单。
pub fn product_info() -> &'static ProductInfo {
    static CACHED: OnceLock<ProductInfo> = OnceLock::new();
    CACHED.get_or_init(|| {
        let parsed = manifest_search_paths()
            .into_iter()
            .find(|path| path.is_file())
            .and_then(|path| {
                std::fs::read_to_string(&path)
                    .ok()
                    .and_then(|text| parse_manifest(&text).ok())
            });
        parsed
            .or_else(|| parse_manifest(EMBEDDED_MANIFEST).ok())
            .unwrap_or_else(|| ProductInfo {
                id: PRODUCT_NAME.to_owned(),
                name: DEFAULT_NAME.to_owned(),
                version: PRODUCT_VERSION.to_owned(),
            })
    })
}

/// 界面/窗口标题：「审核工具 V26.3.0.0」。
pub fn app_title() -> String {
    product_info().title()
}

/// 发布产物名：`AuditTools_26.3.0.0`。
pub fn executable_stem() -> String {
    let info = product_info();
    format!("{}_{}", PRODUCT_NAME, info.version)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn embedded_manifest_matches_the_version_const() {
        let info = parse_manifest(EMBEDDED_MANIFEST).expect("embedded manifest parses");
        assert_eq!(
            info.version, PRODUCT_VERSION,
            "VERSION_MANIFEST.json 与 PRODUCT_VERSION 必须同步更新"
        );
        assert_eq!(info.name, "审核工具");
        assert_eq!(info.id, PRODUCT_NAME);
        assert_eq!(info.title(), format!("审核工具 V{PRODUCT_VERSION}"));
    }

    #[test]
    fn version_scheme_rejects_bad_values() {
        for good in ["26.3.0.0", "26.2.2.5", "99.65535.0.1"] {
            validate_version(good).unwrap_or_else(|error| panic!("{good}: {error}"));
        }
        for bad in ["0.1.0", "3.0.0.0", "26.3.0", "26.a.0.0", "26.3.0.65536"] {
            assert!(validate_version(bad).is_err(), "{bad} 应被拒绝");
        }
    }

    #[test]
    fn product_dir_is_worker_compatible() {
        assert!(PRODUCT_DIR
            .bytes()
            .all(|byte| byte.is_ascii_lowercase() || byte.is_ascii_digit() || byte == b'-'));
        assert_eq!(PRODUCT_NAME, "AuditTools");
    }
}
