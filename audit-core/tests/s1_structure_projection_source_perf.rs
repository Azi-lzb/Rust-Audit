use anyhow::{Context, Result};
use audit_core::{
    models::CopyRange,
    s1_structure_compare::{
        compare_s1_source_structure, compare_s1_source_structure_dump, read_s1_structure_template,
    },
};
use excel_reader::read_workbook;
use serde_json::json;
use sha2::{Digest, Sha256};
use std::{env, fs, path::PathBuf, time::Instant};

#[test]
#[ignore = "manual same-input performance comparison for the S1 structure source reader"]
fn profile_source_structure_projection_same_input() -> Result<()> {
    let source = PathBuf::from(env::var("AUDIT_S1_STRUCTURE_PROFILE_SOURCE")?);
    let template_path = PathBuf::from(env::var("AUDIT_S1_STRUCTURE_PROFILE_TEMPLATE")?);
    let rounds = env::var("AUDIT_S1_STRUCTURE_PROFILE_ROUNDS")
        .unwrap_or_else(|_| "5".to_owned())
        .parse::<usize>()?;
    let warmup = env::var("AUDIT_S1_STRUCTURE_PROFILE_WARMUP")
        .unwrap_or_else(|_| "1".to_owned())
        .parse::<usize>()?;
    anyhow::ensure!(rounds > 0, "rounds must be positive");

    let ranges = [CopyRange {
        sheet_name: "S1Conditional".to_owned(),
        address: "A1:H1".to_owned(),
    }];
    let template = read_s1_structure_template(&template_path, &ranges)?;
    let source_before = sha256(&source)?;
    let template_before = sha256(&template_path)?;
    let mut samples = Vec::with_capacity(rounds);

    for index in 0..(warmup + rounds) {
        let projection_first = index % 2 == 1;
        let (full, projection) = if projection_first {
            let projected = timed_projection(&source, &template)?;
            let full = timed_full_dump(&source, &template)?;
            (full, projected)
        } else {
            let full = timed_full_dump(&source, &template)?;
            let projected = timed_projection(&source, &template)?;
            (full, projected)
        };
        ensure_same_match(&full.result, &projection.result)?;
        if index >= warmup {
            samples.push(json!({
                "round": index - warmup + 1,
                "order": if projection_first { "projection_then_full" } else { "full_then_projection" },
                "full_read_us": full.read_us,
                "full_compare_us": full.compare_us,
                "full_total_us": full.read_us + full.compare_us,
                "projection_total_us": projection.total_us,
                "projection_read_ms_reported": projection.result.read_time_ms,
                "projection_compare_ms_reported": projection.result.elapsed_ms,
                "matched_labels": projection.result.matched_labels,
                "checked_labels": projection.result.checked_labels,
                "matched": projection.result.matched
            }));
        }
    }

    let source_after = sha256(&source)?;
    let template_after = sha256(&template_path)?;
    anyhow::ensure!(
        source_before == source_after,
        "source workbook changed during profiling"
    );
    anyhow::ensure!(
        template_before == template_after,
        "template workbook changed during profiling"
    );
    println!(
        "S1_STRUCTURE_SOURCE_PROFILE={}",
        json!({
            "schema_version": 1,
            "source": source,
            "template": template_path,
            "source_bytes": fs::metadata(&source)?.len(),
            "template_bytes": fs::metadata(&template_path)?.len(),
            "source_sha256_before": source_before,
            "source_sha256_after": source_after,
            "template_sha256_before": template_before,
            "template_sha256_after": template_after,
            "warmup": warmup,
            "rounds": samples,
            "comparison": "full WorkbookDump read + compare_s1_source_structure_dump vs compare_s1_source_structure projected file path",
            "inputs_unchanged": true
        })
    );
    Ok(())
}

struct TimedFull {
    read_us: u128,
    compare_us: u128,
    result: audit_core::s1_structure_compare::S1StructureMatch,
}

struct TimedProjection {
    total_us: u128,
    result: audit_core::s1_structure_compare::S1StructureMatch,
}

fn timed_full_dump(
    source: &std::path::Path,
    template: &audit_core::s1_structure_compare::S1StructureTemplate,
) -> Result<TimedFull> {
    let started = Instant::now();
    let workbook = read_workbook(source)?;
    let read_us = started.elapsed().as_micros();
    let started = Instant::now();
    let result = compare_s1_source_structure_dump(&workbook, template)?;
    Ok(TimedFull {
        read_us,
        compare_us: started.elapsed().as_micros(),
        result,
    })
}

fn timed_projection(
    source: &std::path::Path,
    template: &audit_core::s1_structure_compare::S1StructureTemplate,
) -> Result<TimedProjection> {
    let started = Instant::now();
    let result = compare_s1_source_structure(source, template)?;
    Ok(TimedProjection {
        total_us: started.elapsed().as_micros(),
        result,
    })
}

fn ensure_same_match(
    full: &audit_core::s1_structure_compare::S1StructureMatch,
    projection: &audit_core::s1_structure_compare::S1StructureMatch,
) -> Result<()> {
    anyhow::ensure!(full.matched == projection.matched, "match status differs");
    anyhow::ensure!(full.score == projection.score, "match score differs");
    anyhow::ensure!(
        full.matched_labels == projection.matched_labels,
        "matched label count differs"
    );
    anyhow::ensure!(
        full.checked_labels == projection.checked_labels,
        "checked label count differs"
    );
    anyhow::ensure!(
        full.missing_sheets == projection.missing_sheets,
        "missing sheet list differs"
    );
    anyhow::ensure!(
        full.diagnostic == projection.diagnostic,
        "diagnostic differs"
    );
    Ok(())
}

fn sha256(path: &std::path::Path) -> Result<String> {
    let bytes = fs::read(path).with_context(|| format!("read {} for SHA-256", path.display()))?;
    Ok(format!("{:x}", Sha256::digest(bytes)))
}
