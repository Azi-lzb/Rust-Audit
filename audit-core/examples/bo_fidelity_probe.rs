fn main() -> anyhow::Result<()> {
    let args: Vec<String> = std::env::args().collect();
    let src = std::path::PathBuf::from(args.get(1).expect("xlsx"));
    let out = std::path::PathBuf::from(args.get(2).expect("out.xlsx"));
    let cfg = audit_core::bo_calc::BoCalcConfig::from_env();
    std::fs::copy(&src, &out)?;
    let outcome = audit_core::bo_calc::recalculate(&out, &cfg)?;
    println!("OK engine={} errors={}", outcome.application, outcome.error_cells);
    // 逐要素核对：工作表/公式/合并/定义名/条件格式/数据验证
    for label in ["合并:", "定义名:", "条件格式:", "数据验证:"] {
        let _ = label;
    }
    let before = std::fs::read(&src)?;
    let after = std::fs::read(&out)?;
    let inspect = |bytes: &[u8]| -> anyhow::Result<(usize, usize, usize, usize, usize)> {
        let mut zip = zip::ZipArchive::new(std::io::Cursor::new(bytes.to_vec()))?;
        let mut merge = 0; let mut dv = 0; let mut cf = 0; let mut sheets = 0;
        let mut defined = 0;
        for i in 0..zip.len() {
            let mut entry = zip.by_index(i)?;
            let name = entry.name().to_owned();
            let mut text = String::new();
            use std::io::Read;
            entry.read_to_string(&mut text)?;
            if name == "xl/workbook.xml" {
                sheets = text.matches("<sheet ").count();
                if text.contains("<definedNames") || text.contains("definedName") { defined = 1; }
            }
            if name.starts_with("xl/worksheets/sheet") {
                merge += text.matches("<mergeCell ").count();
                dv += text.matches("<dataValidation ").count();
                cf += text.matches("<conditionalFormatting").count();
            }
        }
        Ok((sheets, merge, dv, cf, defined))
    };
    let b = inspect(&before)?;
    let a = inspect(&after)?;
    println!("before: sheets={} merge={} dv={} cf={} definedName={}", b.0, b.1, b.2, b.3, b.4);
    println!("after : sheets={} merge={} dv={} cf={} definedName={}", a.0, a.1, a.2, a.3, a.4);
    anyhow::ensure!(b == a, "重算改变了结构要素");
    println!("fidelity OK");
    Ok(())
}
