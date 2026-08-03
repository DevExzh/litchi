use litchi_ooxml::xlsx::outline_properties::{
    WorksheetOutlineProperties, parse_worksheet_outline_properties,
};

#[test]
fn host_reexports_the_canonical_outline_properties_owner() {
    let document = concat!(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
        r#"<sheetPr><outlinePr applyStyles="1" summaryBelow="0" summaryRight="false" showOutlineSymbols="0"/></sheetPr>"#,
        r#"</worksheet>"#,
    );
    let value = parse_worksheet_outline_properties(document.as_bytes())
        .unwrap()
        .unwrap();

    fn accepts_canonical_owner(_: &litchi_xlsx::outline_properties::WorksheetOutlineProperties) {}
    accepts_canonical_owner(&value);

    let _: &WorksheetOutlineProperties = &value;
    assert!(value.apply_styles());
    assert!(!value.summary_below());
    assert!(!value.summary_right());
    assert!(!value.show_outline_symbols());
}
