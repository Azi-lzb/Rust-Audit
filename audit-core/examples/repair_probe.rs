fn main() -> anyhow::Result<()> {
    let path = std::path::PathBuf::from(std::env::args().nth(1).expect("xlsx path"));
    let report = audit_core::excel_open_repair::repair_audit_copy_for_excel(&path);
    println!(
        "changed={} quoted={} external={} pane={} shared_ref={} reason={}",
        report.changed,
        report.quoted,
        report.external_ref_fixed,
        report.pane_fixed,
        report.shared_formula_ref_fixed,
        report.reason
    );
    Ok(())
}
