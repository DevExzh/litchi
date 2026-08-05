use litchi_ooxml::xlsx::phonetic_properties::{
    PhoneticAlignment, PhoneticProperties, PhoneticType, parse_phonetic_properties,
};

#[test]
fn host_reexports_the_canonical_phonetic_properties_owner() {
    let document = concat!(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
        r#"<phoneticPr fontId="3" type="Hiragana" alignment="center"/>"#,
        r#"</worksheet>"#,
    );
    let value = parse_phonetic_properties(document.as_bytes())
        .unwrap()
        .unwrap();

    fn accepts_canonical_owner(_: &litchi_xlsx::phonetic_properties::PhoneticProperties) {}
    accepts_canonical_owner(&value);

    let _: &PhoneticProperties = &value;
    assert_eq!(value.font_id(), 3);
    assert_eq!(value.phonetic_type(), PhoneticType::Hiragana);
    assert_eq!(value.alignment(), PhoneticAlignment::Center);
}
