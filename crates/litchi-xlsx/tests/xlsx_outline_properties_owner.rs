use litchi_xlsx::outline_properties::{OutlineProperties, parse_outline_properties};

#[test]
fn standalone_outline_properties_parse_through_the_owner() {
    let document = concat!(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
        r#"<sheetPr><outlinePr applyStyles="1" summaryBelow="0" summaryRight="false" showOutlineSymbols="0"/></sheetPr>"#,
        r#"</worksheet>"#,
    );
    let value = parse_outline_properties(document.as_bytes())
        .unwrap()
        .unwrap();

    fn accepts_canonical_owner(_: &OutlineProperties) {}
    accepts_canonical_owner(&value);

    let _: &OutlineProperties = &value;
    assert!(value.apply_styles());
    assert!(!value.summary_below());
    assert!(!value.summary_right());
    assert!(!value.show_outline_symbols());
}
