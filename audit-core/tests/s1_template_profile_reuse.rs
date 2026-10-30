use audit_core::s1_template::{
    load_template_profiles, recommend_template_for_workbooks, recommend_template_with_profiles,
};
use serde_json::json;
use std::{
    fs,
    path::{Path, PathBuf},
    time::{Instant, SystemTime, UNIX_EPOCH},
};

struct TempDirectory(PathBuf);

impl TempDirectory {
    fn create() -> anyhow::Result<Self> {
        let suffix = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
        let path = std::env::temp_dir().join(format!("audit_s1_template_profile_reuse_{suffix}"));
        fs::create_dir_all(&path)?;
        Ok(Self(path))
    }
}

impl Drop for TempDirectory {
    fn drop(&mut self) {
        let _ = fs::remove_dir_all(&self.0);
    }
}

fn write_workbook(path: &Path, sheet_names: &[String]) -> anyhow::Result<()> {
    let mut workbook = umya_spreadsheet::new_file();
    anyhow::ensure!(!sheet_names.is_empty(), "fixture needs at least one sheet");
    workbook
        .sheet_by_name_mut("Sheet1")
        .expect("new workbook has its default sheet")
        .set_name(&sheet_names[0]);
    for name in sheet_names.iter().skip(1) {
        workbook.new_sheet(name)?;
    }
    umya_spreadsheet::writer::xlsx::write(&workbook, path)?;
    Ok(())
}

fn sheet_names(prefix: &str, count: usize) -> Vec<String> {
    (1..=count)
        .map(|index| format!("{prefix}{index:02}"))
        .collect()
}

fn assert_same_recommendations(
    template_dir: &Path,
    profiles: &[audit_core::s1_template::TemplateProfile],
    sources: &[PathBuf],
) -> anyhow::Result<()> {
    for source in sources {
        let source_batch = [source.clone()];
        let repeated = recommend_template_for_workbooks(template_dir, &source_batch)?;
        let reused = recommend_template_with_profiles(profiles, &source_batch);
        assert_eq!(repeated, reused, "recommendation changed for {source:?}");
        assert!(reused.matched, "fixture did not match for {source:?}");
        assert!(reused.template_path.as_deref().is_some_and(|path| path
            .file_name()
            .is_some_and(|name| name == "template-00-target.xlsx")));
    }
    Ok(())
}

#[test]
fn s1_template_profile_reuse_preserves_batch_matches() -> anyhow::Result<()> {
    let temp = TempDirectory::create()?;
    let template_dir = temp.0.join("templates");
    let source_dir = temp.0.join("sources");
    fs::create_dir_all(&template_dir)?;
    fs::create_dir_all(&source_dir)?;

    let source_sheets = sheet_names("Data", 8);
    let mut target_sheets = source_sheets.clone();
    target_sheets.push("审核规则".to_owned());
    write_workbook(
        &template_dir.join("template-00-target.xlsx"),
        &target_sheets,
    )?;
    for index in 1..5 {
        let other_sheets = sheet_names(&format!("Other{index}-"), 9);
        write_workbook(
            &template_dir.join(format!("template-{index:02}-other.xlsx")),
            &other_sheets,
        )?;
    }

    let mut sources = Vec::with_capacity(21);
    for index in 0..21 {
        let path = source_dir.join(format!("source-{index:02}.xlsx"));
        write_workbook(&path, &source_sheets)?;
        sources.push(path);
    }

    let mut template_bytes = 0u64;
    for entry in fs::read_dir(&template_dir)? {
        template_bytes += entry?.metadata()?.len();
    }
    let source_bytes = sources
        .iter()
        .map(|path| fs::metadata(path).map(|meta| meta.len()))
        .collect::<Result<Vec<_>, _>>()?
        .into_iter()
        .sum::<u64>();

    let warmups = 1usize;
    let rounds = 5usize;
    let mut repeated_ms = Vec::with_capacity(rounds);
    let mut reused_ms = Vec::with_capacity(rounds);
    for round in 0..(warmups + rounds) {
        let repeated_started = Instant::now();
        for source in &sources {
            let source_batch = [source.clone()];
            let _ = recommend_template_for_workbooks(&template_dir, &source_batch)?;
        }
        let repeated_elapsed = repeated_started.elapsed().as_secs_f64() * 1000.0;

        let reused_started = Instant::now();
        let profiles = load_template_profiles(&template_dir)?;
        for source in &sources {
            let source_batch = [source.clone()];
            let _ = recommend_template_with_profiles(&profiles, &source_batch);
        }
        let reused_elapsed = reused_started.elapsed().as_secs_f64() * 1000.0;

        if round >= warmups {
            repeated_ms.push(repeated_elapsed);
            reused_ms.push(reused_elapsed);
        }
    }

    let profiles = load_template_profiles(&template_dir)?;
    assert_same_recommendations(&template_dir, &profiles, &sources)?;
    repeated_ms.sort_by(f64::total_cmp);
    reused_ms.sort_by(f64::total_cmp);
    eprintln!(
        "S1_TEMPLATE_PROFILE_REUSE_BENCHMARK={}",
        json!({
            "source_files": sources.len(),
            "template_files": profiles.len(),
            "source_sheets": source_sheets.len(),
            "template_sheets_per_workbook": 9,
            "nonempty_cells": 0,
            "source_bytes_total": source_bytes,
            "template_bytes_total": template_bytes,
            "warmups": warmups,
            "rounds": rounds,
            "legacy_reloaded_templates_per_source": true,
            "median_legacy_ms": repeated_ms[rounds / 2],
            "median_profiles_reused_ms": reused_ms[rounds / 2],
            "speedup": repeated_ms[rounds / 2] / reused_ms[rounds / 2],
            "all_21_recommendations_identical": true,
        })
    );
    Ok(())
}
