//! Two-period workbook/sheet pairing for the Audit collection frontend.

use crate::{
    discovery::ScanDepth,
    report_source::{report_source_identity, ReportSourceIdentity},
};
use anyhow::{Context, Result};
use excel_reader::read_sheet_names;
use serde::Serialize;
use std::{
    collections::{BTreeMap, HashMap},
    fs,
    path::Path,
};

#[derive(Clone, Debug, PartialEq, Eq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct PeriodPair {
    pub org: String,
    pub form: String,
    pub org_name: String,
    pub social_credit_code: String,
    pub cur_name: String,
    pub cur_path: String,
    pub cur_sheet: String,
    pub cur_date: String,
    pub pre_name: String,
    pub pre_path: String,
    pub pre_sheet: String,
    pub pre_date: String,
    pub matched: bool,
    pub side: String,
}

#[derive(Clone, Debug)]
struct SourceSheet {
    org_key: String,
    org_name: String,
    social_credit_code: String,
    form: String,
    name: String,
    path: String,
    sheet: String,
    date: String,
}

impl SourceSheet {
    fn from_identity(path: &Path, sheet: &str, identity: ReportSourceIdentity) -> Self {
        let org_key = if identity.institution_id.is_empty() {
            identity.institution_name.clone()
        } else {
            identity.institution_id.clone()
        };
        Self {
            org_key,
            org_name: identity.institution_name,
            social_credit_code: identity.institution_id,
            form: identity.form_key,
            name: path
                .file_name()
                .map(|name| name.to_string_lossy().into_owned())
                .unwrap_or_else(|| path.display().to_string()),
            path: path.to_string_lossy().into_owned(),
            sheet: sheet.to_owned(),
            date: identity.data_date,
        }
    }
}

/// Match current and previous period items by institution identity and form.
///
/// Only workbook names and worksheet directories are read; worksheet cell
/// contents are not scanned. Duplicate `(institution, form)` items in a
/// directory follow the original Python behavior: the last lexically sorted
/// file/sheet item wins.
pub fn list_period_pairs(
    current_dir: Option<&Path>,
    previous_dir: Option<&Path>,
) -> Result<Vec<PeriodPair>> {
    list_period_pairs_with_scan_depth(current_dir, previous_dir, ScanDepth::RootOnly)
}

/// Match current and previous period workbooks using the configured directory
/// depth. The original entry point remains root-only for existing CLI callers.
pub fn list_period_pairs_with_scan_depth(
    current_dir: Option<&Path>,
    previous_dir: Option<&Path>,
    depth: ScanDepth,
) -> Result<Vec<PeriodPair>> {
    let current = collect(current_dir, depth)?;
    let previous = collect(previous_dir, depth)?;
    let mut keys = current
        .keys()
        .chain(previous.keys())
        .cloned()
        .collect::<Vec<_>>();
    keys.sort();
    keys.dedup();

    Ok(keys
        .into_iter()
        .map(|key| {
            let cur = current.get(&key);
            let pre = previous.get(&key);
            let example = cur.or(pre);
            let matched = cur.is_some() && pre.is_some();
            PeriodPair {
                org: key.0.clone(),
                form: key.1.clone(),
                org_name: example
                    .map(|item| {
                        if item.org_name.is_empty() {
                            key.0.clone()
                        } else {
                            item.org_name.clone()
                        }
                    })
                    .unwrap_or_else(|| key.0.clone()),
                social_credit_code: example
                    .map(|item| item.social_credit_code.clone())
                    .unwrap_or_default(),
                cur_name: cur.map(|item| item.name.clone()).unwrap_or_default(),
                cur_path: cur.map(|item| item.path.clone()).unwrap_or_default(),
                cur_sheet: cur.map(|item| item.sheet.clone()).unwrap_or_default(),
                cur_date: cur.map(|item| item.date.clone()).unwrap_or_default(),
                pre_name: pre.map(|item| item.name.clone()).unwrap_or_default(),
                pre_path: pre.map(|item| item.path.clone()).unwrap_or_default(),
                pre_sheet: pre.map(|item| item.sheet.clone()).unwrap_or_default(),
                pre_date: pre.map(|item| item.date.clone()).unwrap_or_default(),
                matched,
                side: if matched {
                    "both".to_owned()
                } else if cur.is_some() {
                    "cur".to_owned()
                } else {
                    "pre".to_owned()
                },
            }
        })
        .collect())
}

/// Recreate the source application's banks filename reminder.
pub fn report_source_warning(current_dir: Option<&Path>, previous_dir: Option<&Path>) -> String {
    report_source_warning_with_scan_depth(current_dir, previous_dir, ScanDepth::RootOnly)
}

pub fn report_source_warning_with_scan_depth(
    current_dir: Option<&Path>,
    previous_dir: Option<&Path>,
    depth: ScanDepth,
) -> String {
    let mut found = Vec::new();
    for (label, directory) in [("本期", current_dir), ("上期", previous_dir)] {
        let Some(directory) = directory.filter(|directory| directory.is_dir()) else {
            continue;
        };
        let count = collect_excel_paths(directory, depth)
            .unwrap_or_default()
            .into_iter()
            .filter(|path| {
                let stem = path
                    .file_stem()
                    .map(|value| value.to_string_lossy().into_owned())
                    .unwrap_or_default();
                let parts: Vec<_> = stem.split('#').collect();
                parts.len() >= 4 && parts[0].trim().eq_ignore_ascii_case("banks")
            })
            .count();
        if count > 0 {
            found.push(format!("{label}{count}个"));
        }
    }
    if found.is_empty() {
        return String::new();
    }
    format!(
        "发现 banks 导出文件（{}）。该格式不含机构名称，名称需依赖机构参照映射。建议在来源系统选择“默认单sheet”或“按报表单sheet”重新导出，使用“机构代码#日期#01#表单#机构名称”或“reports#机构代码#日期#01#机构名称”这两类文件名。",
        found.join("、")
    )
}

fn collect(
    directory: Option<&Path>,
    depth: ScanDepth,
) -> Result<BTreeMap<(String, String), SourceSheet>> {
    let Some(directory) = directory.filter(|directory| directory.is_dir()) else {
        return Ok(BTreeMap::new());
    };
    let paths = collect_excel_paths(directory, depth)?;

    let mut items = Vec::new();
    for path in paths {
        let names = read_sheet_names(&path).with_context(|| {
            format!(
                "无法读取报送文件“{}”的工作表列表；请确认文件完整且为 .xls/.xlsx 格式，必要时重新导出。",
                path.file_name()
                    .map(|name| name.to_string_lossy())
                    .unwrap_or_default()
            )
        })?;
        for sheet in names {
            let identity = report_source_identity(&path, &sheet);
            items.push(SourceSheet::from_identity(&path, &sheet, identity));
        }
    }

    let mut names_by_code = HashMap::new();
    for item in &items {
        if !item.org_key.is_empty() && !item.org_name.is_empty() {
            names_by_code.insert(item.org_key.clone(), item.org_name.clone());
        }
    }

    let mut result = BTreeMap::new();
    for mut item in items {
        if item.org_name.is_empty() {
            item.org_name = names_by_code
                .get(&item.org_key)
                .cloned()
                .unwrap_or_else(|| item.org_key.clone());
        }
        result.insert((item.org_key.clone(), item.form.clone()), item);
    }
    Ok(result)
}

fn collect_excel_paths(directory: &Path, max_depth: ScanDepth) -> Result<Vec<std::path::PathBuf>> {
    fn visit(
        directory: &Path,
        directory_depth: usize,
        max_depth: ScanDepth,
        paths: &mut Vec<std::path::PathBuf>,
    ) -> Result<()> {
        for entry in fs::read_dir(directory)
            .with_context(|| format!("读取报送目录失败：{}", directory.display()))?
        {
            let entry = entry?;
            let path = entry.path();
            let kind = entry.file_type()?;
            if kind.is_dir() {
                let should_recurse = match max_depth {
                    ScanDepth::RootOnly => false,
                    ScanDepth::MaxDirectoryDepth(max) => directory_depth < max,
                    ScanDepth::Full => true,
                };
                if should_recurse && entry.file_name() != "执行结果" {
                    visit(&path, directory_depth + 1, max_depth, paths)?;
                }
            } else if kind.is_file()
                && is_excel_file(&path)
                && !path
                    .file_name()
                    .is_some_and(|name| name.to_string_lossy().starts_with("~$"))
            {
                paths.push(path);
            }
        }
        Ok(())
    }

    let mut paths = Vec::new();
    visit(directory, 0, max_depth, &mut paths)?;
    paths.sort();
    Ok(paths)
}

fn is_excel_file(path: &Path) -> bool {
    path.extension()
        .and_then(|extension| extension.to_str())
        .is_some_and(|extension| matches!(extension.to_ascii_lowercase().as_str(), "xls" | "xlsx"))
}

#[cfg(test)]
mod tests {
    use super::{list_period_pairs, report_source_warning};
    use std::{
        fs,
        path::{Path, PathBuf},
        time::{SystemTime, UNIX_EPOCH},
    };

    struct TestDirs(PathBuf);

    impl TestDirs {
        fn new() -> Self {
            let suffix = SystemTime::now()
                .duration_since(UNIX_EPOCH)
                .expect("clock")
                .as_nanos();
            let root = std::env::temp_dir().join(format!("audit-period-pairs-{suffix}"));
            fs::create_dir_all(root.join("current")).expect("create current dir");
            fs::create_dir_all(root.join("previous")).expect("create previous dir");
            Self(root)
        }

        fn copy_fixture(&self, source: &str, directory: &str, name: &str) {
            let fixture = Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("tests/fixtures/s2_import")
                .join(source);
            fs::copy(fixture, self.0.join(directory).join(name)).expect("copy fixture");
        }

        fn path(&self, directory: &str) -> PathBuf {
            self.0.join(directory)
        }
    }

    impl Drop for TestDirs {
        fn drop(&mut self) {
            let _ = fs::remove_dir_all(&self.0);
        }
    }

    #[test]
    fn pairs_by_institution_and_sheet_form_and_keeps_one_sided_items() {
        let dirs = TestDirs::new();
        let fixture = "9144#2026-06-30#x#20202#甲银行.xlsx";
        dirs.copy_fixture(fixture, "current", "9144#2026-06-30#x#20202#甲银行.xlsx");
        dirs.copy_fixture(fixture, "previous", "9144#2026-03-31#x#20202#甲银行.xlsx");
        dirs.copy_fixture(fixture, "current", "banks#2026-06-30#01#20202.xlsx");

        let pairs = list_period_pairs(Some(&dirs.path("current")), Some(&dirs.path("previous")))
            .expect("read pairs");
        assert!(pairs
            .iter()
            .any(|pair| pair.matched && pair.form == "20202"));
        assert!(pairs.iter().any(|pair| !pair.matched && pair.side == "cur"));
        let standard = pairs.iter().find(|pair| pair.org == "9144").unwrap();
        assert_eq!(standard.org_name, "甲银行");
        assert_eq!(standard.social_credit_code, "9144");
        assert_eq!(standard.cur_date, "2026-06-30");
        assert_eq!(standard.pre_date, "2026-03-31");
    }

    #[test]
    fn banks_warning_counts_only_supported_direct_child_exports() {
        let dirs = TestDirs::new();
        dirs.copy_fixture(
            "9144#2026-06-30#x#20202#甲银行.xlsx",
            "current",
            "banks#2026-06-30#01#20202.xlsx",
        );
        let warning = report_source_warning(Some(&dirs.path("current")), None);
        assert!(warning.starts_with("发现 banks 导出文件（本期1个）。"));
        assert!(warning.contains("机构代码#日期#01#表单#机构名称"));
    }
}
