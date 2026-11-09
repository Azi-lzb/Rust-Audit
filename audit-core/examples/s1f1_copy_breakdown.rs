//! 拆解 S1-F01「副本准备」各步骤耗时。
//!
//! 用法：s1f1_copy_breakdown <source.xlsx> <template.xlsx> <external.xlsx>
use std::path::PathBuf;
use std::time::Instant;

fn main() -> anyhow::Result<()> {
    let args: Vec<String> = std::env::args().collect();
    let source = PathBuf::from(args.get(1).expect("source"));
    let template = PathBuf::from(args.get(2).expect("template"));
    let external = PathBuf::from(args.get(3).expect("external"));
    let work = std::env::temp_dir().join(format!("s1_copy_bd_{}", std::process::id()));
    std::fs::create_dir_all(&work)?;

    let step = |label: &str, f: &dyn Fn() -> anyhow::Result<String>| {
        let t = Instant::now();
        match f() {
            Ok(detail) => println!("{:>7} ms | {label} | {detail}", t.elapsed().as_millis()),
            Err(e) => println!(
                "{:>7} ms | {label} | FAILED: {e:#}",
                t.elapsed().as_millis()
            ),
        }
    };

    // 1) calamine 全量读源工作簿（判断有无公式 + 后续计划）
    step("读取源工作簿(calamine)", &|| {
        let wb = excel_reader::read_workbook(&source)?;
        let formulas: usize = wb.sheets.iter().map(|s| s.formulas.len()).sum();
        Ok(format!("{} 表 / {formulas} 公式", wb.sheets.len()))
    });

    // 2) 第一次拆缓存（源 -> cacheless）
    let cacheless = work.join("cacheless.xlsx");
    step("拆缓存 #1 (ZIP 重写)", &|| {
        let r =
            audit_core::s1_formula_copy::prepare_s1_recalculation_source_copy(&source, &cacheless)?;
        Ok(match r {
            Some(report) => format!(
                "清除 {} 缓存 / {} 公式",
                report.cache_values_removed, report.formula_cells
            ),
            None => "无公式".to_owned(),
        })
    });

    // 3) 构建编辑计划（读模板 + 外部文件 + 生成操作序列）
    let editor_output = work.join("formula-copy.xlsx");
    let plan_started = Instant::now();
    let plan = audit_core::s1_formula_copy::build_s1_formula_copy_edit_plan(
        &cacheless,
        &template,
        &editor_output,
        None,
        Some(&external),
    );
    println!(
        "{:>7} ms | 构建编辑计划 | {}",
        plan_started.elapsed().as_millis(),
        match &plan {
            Ok(p) => format!("{} 个操作", p.edit_plan.operations.len()),
            Err(e) => format!("FAILED: {e:#}"),
        }
    );

    // 4) 执行编辑（写审核副本）
    if let Ok(plan) = plan {
        // 4)+5) 执行编辑与候选校验都由 modifier.execute 内部完成（含验证）。
        let modifier_result =
            audit_core::excel_route::production_modifier("S1-F01").map(|(modifier, _)| modifier);
        let edit_started = Instant::now();
        let outcome: anyhow::Result<_> = match modifier_result {
            Ok(modifier) => modifier.execute(&plan.edit_plan).map_err(Into::into),
            Err(error) => Err(error),
        };
        match outcome {
            Ok(report) => println!(
                "{:>7} ms | 执行编辑+校验(写副本) | engine={:?}",
                edit_started.elapsed().as_millis(),
                report.engine
            ),
            Err(e) => println!(
                "{:>7} ms | 执行编辑+校验(写副本) | FAILED: {e:#}",
                edit_started.elapsed().as_millis()
            ),
        }

        // 6) 第二次拆缓存（编辑结果 -> cacheless）
        let cacheless2 = work.join("cacheless2.xlsx");
        step("拆缓存 #2 (ZIP 重写)", &|| {
            let r = audit_core::s1_formula_copy::prepare_s1_recalculation_source_copy(
                &editor_output,
                &cacheless2,
            )?;
            Ok(match r {
                Some(report) => format!(
                    "清除 {} 缓存 / {} 公式",
                    report.cache_values_removed, report.formula_cells
                ),
                None => "无公式".to_owned(),
            })
        });
    }

    // 7) OOXML 规范化（repair_audit_copy_for_excel，流程里必跑且原地改文件）
    let repair_target = work.join("repair-target.xlsx");
    if let Ok(Some(_)) =
        audit_core::s1_formula_copy::prepare_s1_recalculation_source_copy(&source, &repair_target)
    {
        let t = Instant::now();
        let report = audit_core::excel_open_repair::repair_audit_copy_for_excel(&repair_target);
        println!(
            "{:>7} ms | OOXML 规范化(repair) | changed={}",
            t.elapsed().as_millis(),
            report.changed
        );
    }

    let _ = std::fs::remove_dir_all(&work);
    Ok(())
}
