use litchi_ods::Spreadsheet;

#[test]
fn reads_libreoffice_standard_conditional_style_maps_when_fixture_is_available() {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/libreoffice-core/sc/qa/unit/data/ods/formats.ods");
    if !path.exists() {
        return;
    }

    let spreadsheet = Spreadsheet::open(path).unwrap();
    let ce35 = spreadsheet.conditional_cell_style("ce35").unwrap();
    assert_eq!(ce35.rules.len(), 3);
    assert_eq!(ce35.rules[0].condition, "cell-content()<1");
    assert_eq!(ce35.rules[0].apply_style_name, "Style1");
    assert_eq!(
        ce35.rules[0].base_cell_address.as_deref(),
        Some("Sheet3.A1")
    );
    assert_eq!(ce35.rules[1].condition, "cell-content-is-between(1,2)");
    assert_eq!(ce35.rules[1].apply_style_name, "Untitled1");
    assert_eq!(ce35.rules[2].condition, "cell-content()>2");
    assert_eq!(ce35.rules[2].apply_style_name, "Untitled2");

    let ce36 = spreadsheet.conditional_cell_style("ce36").unwrap();
    assert_eq!(ce36.rules.len(), 4);
    assert_eq!(ce36.rules[0].condition, "is-true-formula(SUM([.E10:.E13]))");
}
