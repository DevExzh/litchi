use litchi_ooxml::xlsx::phonetic_properties::{
    WorksheetPhoneticAlignment, WorksheetPhoneticProperties, WorksheetPhoneticType,
    parse_worksheet_phonetic_properties,
};

#[test]
fn host_reexports_the_canonical_phonetic_properties_owner() {
    let document = concat!(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
        r#"<phoneticPr fontId="3" type="Hiragana" alignment="center"/>"#,
        r#"</worksheet>"#,
    );
    let value = parse_worksheet_phonetic_properties(document.as_bytes())
        .unwrap()
        .unwrap();

    fn accepts_canonical_owner(_: &litchi_xlsx::phonetic_properties::WorksheetPhoneticProperties) {}
    accepts_canonical_owner(&value);

    let _: &WorksheetPhoneticProperties = &value;
    assert_eq!(value.font_id(), 3);
    assert_eq!(value.phonetic_type(), WorksheetPhoneticType::Hiragana);
    assert_eq!(value.alignment(), WorksheetPhoneticAlignment::Center);
}
