//! 发行包配置更新检查：对比随包配置与用户配置，备份后再更新。
//!
//! 与源生产 `base_audit/config_update.py` 口径一致：
//! - 仅 Linux 安装包适用；Windows 是免安装目录布局，配置随程序目录更换，
//!   该页签不显示（`supported()=false`）。
//! - 随包版本来自安装根的 `VERSION_MANIFEST.json`（configs 组件表），可信；
//!   用户侧版本来自本模块记录的 `.配置版本.json`，首次记录时只在本机文件与
//!   随包内容完全一致时写入，不凭空声明版本。
//! - 两侧内容不同且用户侧无记录 → `unknown`，如实提示"无法确认版本"；
//!   绝不谎称"有新版"。
//! - 更新前逐文件备份到 `备份-配置更新-时间戳` 目录，逐文件校验哈希，
//!   任一失败即停止并报告已完成项；更新完成后由界面提示重启。

use std::collections::BTreeMap;
use std::fs;
use std::path::{Path, PathBuf};

use anyhow::{Context, Result};
use serde::Serialize;
use sha2::{Digest, Sha256};

const RECORD_NAME: &str = ".配置版本.json";
const BACKUP_PREFIX: &str = "备份-配置更新-";

/// Windows 免安装布局不提供该页签；Linux 安装包才检查。
pub fn supported() -> bool {
    !cfg!(windows)
}

pub fn unsupported_reason() -> &'static str {
    "Windows 版为免安装目录布局，配置随程序目录一并更换，不需要单独更新配置。"
}

fn sha256(path: &Path) -> Result<String> {
    let bytes = fs::read(path).with_context(|| format!("读取文件失败：{}", path.display()))?;
    let digest = Sha256::digest(&bytes);
    Ok(hex(&digest))
}

fn hex(digest: &[u8]) -> String {
    digest.iter().map(|byte| format!("{byte:02x}")).collect()
}

/// 随包配置目录：安装根下的 `config`（由包管理器维护、只读）。
/// 冻结包由启动脚本设置 `AUDIT_INSTALL_ROOT` / `AUDIT_CORE_ROOT`；
/// 直接运行时从可执行文件位置推断（exe 同级或上一级的 config）。
fn bundled_config_dir() -> Option<PathBuf> {
    let mut candidates: Vec<PathBuf> = Vec::new();
    for key in ["AUDIT_INSTALL_ROOT", "AUDIT_CORE_ROOT"] {
        if let Some(root) = std::env::var_os(key).filter(|value| !value.is_empty()) {
            let root = PathBuf::from(root);
            candidates.push(root.join("config"));
            if let Some(parent) = root.parent() {
                candidates.push(parent.join("config"));
            }
        }
    }
    if let Ok(exe) = std::env::current_exe() {
        if let Some(root) = exe.parent() {
            candidates.push(root.join("config"));
            if let Some(parent) = root.parent() {
                candidates.push(parent.join("config"));
            }
        }
    }
    candidates.into_iter().find(|candidate| candidate.is_dir())
}

/// 用户配置目录：与随包 config 对应的本机配置（Linux `$DATA_HOME/config`）。
/// Rust 交付使用与恢复任务一致的状态目录约定；安装包环境由
/// `AUDIT_USER_CONFIG_DIR` 或 `AUDIT_STATE_DIR` 指定。
fn user_config_dir() -> Option<PathBuf> {
    if let Some(dir) = std::env::var_os("AUDIT_USER_CONFIG_DIR").filter(|v| !v.is_empty()) {
        let dir = PathBuf::from(dir);
        if dir.is_dir() {
            return Some(dir);
        }
    }
    if let Some(dir) = std::env::var_os("AUDIT_STATE_DIR").filter(|v| !v.is_empty()) {
        let dir = PathBuf::from(dir).join("config");
        if dir.is_dir() {
            return Some(dir);
        }
    }
    let exe = std::env::current_exe().ok()?;
    let root = exe.parent()?;
    let candidate = root.join("config");
    candidate.is_dir().then_some(candidate)
}

/// `VERSION_MANIFEST.json` 的 configs 组件：文件名 → {component,name,version}。
fn manifest_configs() -> Result<BTreeMap<String, ManifestConfig>> {
    let path = version_manifest_path()?;
    let text = fs::read_to_string(&path)
        .with_context(|| format!("读取版本清单失败：{}", path.display()))?;
    let value: serde_json::Value = serde_json::from_str(&text)
        .with_context(|| format!("解析版本清单失败：{}", path.display()))?;
    let mut indexed = BTreeMap::new();
    let Some(configs) = value.get("configs").and_then(|c| c.as_object()) else {
        return Ok(indexed);
    };
    for (component, item) in configs {
        let file = item
            .get("file")
            .and_then(|f| f.as_str())
            .unwrap_or_default()
            .replace('\\', "/");
        let name = Path::new(&file)
            .file_name()
            .map(|n| n.to_string_lossy().into_owned())
            .unwrap_or_default();
        if name.is_empty() {
            continue;
        }
        indexed.insert(
            name.clone(),
            ManifestConfig {
                component: component.clone(),
                name: item
                    .get("name")
                    .and_then(|n| n.as_str())
                    .map(str::to_owned)
                    .unwrap_or_else(|| name.clone()),
                version: item
                    .get("version")
                    .and_then(|v| v.as_str())
                    .unwrap_or_default()
                    .to_owned(),
            },
        );
    }
    Ok(indexed)
}

fn version_manifest_path() -> Result<PathBuf> {
    for key in ["AUDIT_INSTALL_ROOT", "AUDIT_CORE_ROOT"] {
        if let Some(root) = std::env::var_os(key).filter(|v| !v.is_empty()) {
            let candidate = PathBuf::from(&root).join("VERSION_MANIFEST.json");
            if candidate.is_file() {
                return Ok(candidate);
            }
        }
    }
    let exe = std::env::current_exe().context("定位可执行文件失败")?;
    let root = exe.parent().context("可执行文件无上级目录")?;
    let mut candidates = vec![root.join("VERSION_MANIFEST.json")];
    if let Some(parent) = root.parent() {
        candidates.push(parent.join("VERSION_MANIFEST.json"));
    }
    if let Some(found) = candidates.into_iter().find(|c| c.is_file()) {
        return Ok(found);
    }
    anyhow::bail!("未找到 VERSION_MANIFEST.json")
}

#[derive(Clone, Debug, Serialize)]
struct ManifestConfig {
    component: String,
    name: String,
    version: String,
}

fn version_tuple(text: &str) -> Option<Vec<u64>> {
    let parts: Vec<&str> = text.trim().split('.').collect();
    if parts.is_empty() || parts.iter().any(|part| part.parse::<u64>().is_err()) {
        return None;
    }
    Some(
        parts
            .iter()
            .map(|part| part.parse::<u64>().unwrap())
            .collect(),
    )
}

fn is_newer(bundled: &str, local: &str) -> bool {
    let (Some(left), Some(right)) = (version_tuple(bundled), version_tuple(local)) else {
        return false;
    };
    let width = left.len().max(right.len());
    let pad = |mut v: Vec<u64>| {
        v.resize(width, 0);
        v
    };
    pad(left) > pad(right)
}

fn record_path(config_dir: &Path) -> PathBuf {
    config_dir.join(RECORD_NAME)
}

fn load_record(config_dir: &Path) -> BTreeMap<String, String> {
    let Ok(text) = fs::read_to_string(record_path(config_dir)) else {
        return BTreeMap::new();
    };
    let Ok(value) = serde_json::from_str::<serde_json::Value>(&text) else {
        return BTreeMap::new();
    };
    let Some(files) = value.get("files").and_then(|f| f.as_object()) else {
        return BTreeMap::new();
    };
    files
        .iter()
        .filter_map(|(name, version)| {
            version
                .as_str()
                .map(|version| (name.clone(), version.to_owned()))
        })
        .collect()
}

fn save_record(config_dir: &Path, files: &BTreeMap<String, String>) -> Result<()> {
    let payload = serde_json::json!({"version": 1, "files": files});
    fs::write(
        record_path(config_dir),
        serde_json::to_string_pretty(&payload)?,
    )?;
    Ok(())
}

/// 检查随包配置与用户配置的差异；只读，不修改任何文件。
pub fn inspect() -> serde_json::Value {
    let entries: Vec<serde_json::Value> = Vec::new();
    let not_applicable = |reason: &str| serde_json::json!({"applicable": false, "reason": reason, "entries": entries});
    if !supported() {
        return not_applicable(unsupported_reason());
    }
    let Some(bundled_dir) = bundled_config_dir() else {
        return not_applicable("未找到随包配置目录。");
    };
    let Some(config_dir) = user_config_dir() else {
        return not_applicable("未找到本机配置目录。");
    };
    match (bundled_dir.canonicalize(), config_dir.canonicalize()) {
        (Ok(bundled), Ok(local)) if bundled == local => {
            return not_applicable("当前配置目录就是随包目录。");
        }
        _ => {}
    }
    let configs = match manifest_configs() {
        Ok(configs) if !configs.is_empty() => configs,
        _ => return not_applicable("版本清单未登记配置组件。"),
    };
    let recorded = load_record(&config_dir);
    let mut result_entries = Vec::new();
    for (name, meta) in &configs {
        let bundled = bundled_dir.join(name);
        if !bundled.is_file() {
            continue;
        }
        let Ok(bundled_hash) = sha256(&bundled) else {
            continue;
        };
        let local = config_dir.join(name);
        let local_version = recorded.get(name).cloned().unwrap_or_default();
        let (status, available) = if !local.is_file() {
            ("missing", true)
        } else {
            let Ok(local_hash) = sha256(&local) else {
                continue;
            };
            if local_hash == bundled_hash {
                ("same", false)
            } else if !local_version.is_empty() && is_newer(&meta.version, &local_version) {
                ("update", true)
            } else if !local_version.is_empty() {
                ("local_newer", false)
            } else {
                ("unknown", true)
            }
        };
        result_entries.push(serde_json::json!({
            "component": meta.component,
            "name": name,
            "title": meta.name,
            "bundled_version": meta.version,
            "local_version": local_version,
            "bundled_sha256": bundled_hash,
            "status": status,
            "available": available,
        }));
    }
    serde_json::json!({
        "applicable": true,
        "config_dir": config_dir.display().to_string(),
        "bundled_dir": bundled_dir.display().to_string(),
        "entries": result_entries,
    })
}

/// 备份后把选定配置替换为随包版本；任一文件失败即停止并报告已完成项。
pub fn apply_updates(names: &[String]) -> serde_json::Value {
    let report = inspect();
    if report
        .get("applicable")
        .and_then(serde_json::Value::as_bool)
        != Some(true)
    {
        return serde_json::json!({
            "ok": false,
            "message": report
                .get("reason")
                .and_then(serde_json::Value::as_str)
                .unwrap_or("当前环境不支持配置更新。"),
            "updated": [],
        });
    }
    let selected: std::collections::BTreeSet<&String> = names.iter().collect();
    if selected.is_empty() {
        return serde_json::json!({"ok": false, "message": "未选择要更新的配置文件。", "updated": []});
    }
    let config_dir = PathBuf::from(
        report
            .get("config_dir")
            .and_then(serde_json::Value::as_str)
            .unwrap_or_default(),
    );
    let bundled_dir = PathBuf::from(
        report
            .get("bundled_dir")
            .and_then(serde_json::Value::as_str)
            .unwrap_or_default(),
    );
    let entries: Vec<serde_json::Value> = report
        .get("entries")
        .and_then(serde_json::Value::as_array)
        .cloned()
        .unwrap_or_default();
    let by_name: BTreeMap<&str, &serde_json::Value> = entries
        .iter()
        .filter_map(|entry| {
            entry
                .get("name")
                .and_then(serde_json::Value::as_str)
                .map(|name| (name, entry))
        })
        .collect();
    let unknown: Vec<&String> = selected
        .iter()
        .filter(|name| !by_name.contains_key(name.as_str()))
        .copied()
        .collect();
    if !unknown.is_empty() {
        let names = unknown
            .iter()
            .map(|name| name.as_str())
            .collect::<Vec<_>>()
            .join("、");
        return serde_json::json!({"ok": false, "message": format!("未登记的配置文件：{names}"), "updated": []});
    }
    let timestamp = chrono::Local::now().format("%Y%m%d_%H%M%S");
    let mut backup_dir = config_dir.join(format!("{BACKUP_PREFIX}{timestamp}"));
    let mut bucket = 1;
    while backup_dir.exists() {
        bucket += 1;
        backup_dir = config_dir.join(format!("{BACKUP_PREFIX}{timestamp}_{bucket}"));
    }
    let mut updated: Vec<serde_json::Value> = Vec::new();
    let mut files_record = load_record(&config_dir);
    for name in selected {
        let Some(entry) = by_name.get(name.as_str()) else {
            continue;
        };
        let bundled_hash = entry
            .get("bundled_sha256")
            .and_then(serde_json::Value::as_str)
            .unwrap_or_default()
            .to_owned();
        let source = bundled_dir.join(name.as_str());
        let target = config_dir.join(name.as_str());
        let apply = || -> Result<()> {
            if target.is_file() {
                fs::create_dir_all(&backup_dir)?;
                let backup = backup_dir.join(name.as_str());
                fs::copy(&target, &backup)
                    .with_context(|| format!("备份失败：{}", target.display()))?;
                if sha256(&backup)? != sha256(&target)? {
                    anyhow::bail!("备份哈希校验失败：{}", backup.display());
                }
            }
            fs::copy(&source, &target)
                .with_context(|| format!("复制随包配置失败：{}", source.display()))?;
            let applied = sha256(&target)?;
            if applied != bundled_hash {
                anyhow::bail!("更新后哈希校验失败：{}", target.display());
            }
            Ok(())
        };
        match apply() {
            Ok(()) => {
                if let Some(version) = entry
                    .get("bundled_version")
                    .and_then(serde_json::Value::as_str)
                {
                    files_record.insert(name.clone(), version.to_owned());
                }
                updated.push(serde_json::json!({"name": name, "ok": true}));
            }
            Err(error) => {
                let _ = save_record(&config_dir, &files_record);
                return serde_json::json!({
                    "ok": false,
                    "message": format!("更新 {} 失败，已停止：{error:#}", name),
                    "updated": updated,
                });
            }
        }
    }
    if let Err(error) = save_record(&config_dir, &files_record) {
        return serde_json::json!({
            "ok": false,
            "message": format!("配置更新完成，但写入版本记录失败：{error:#}"),
            "updated": updated,
        });
    }
    serde_json::json!({
        "ok": true,
        "message": format!("已更新 {} 个配置文件；备份目录：{}", updated.len(), backup_dir.display()),
        "backup_dir": backup_dir.display().to_string(),
        "updated": updated,
    })
}
