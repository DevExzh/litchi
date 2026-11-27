use litchi_xlsx::sheet_protection::{
    Conformance, Metadata, Protection, parse_protection, replace_protection,
};

#[test]
fn worksheet_protection_codec_replaces_only_typed_protection_metadata() {
    let mut protection = Protection::new();
    protection.set_sheet_locked(true);
    let mut metadata = Metadata::new();
    metadata.set_sheet_protection(Some(protection)).unwrap();
    let source = br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/><extLst><ext uri="{other}"/></extLst></worksheet>"#;
    let output = replace_protection(source, &metadata).unwrap();
    let parsed = parse_protection(&output).unwrap();
    assert!(parsed.sheet_protection().unwrap().sheet_locked());
    assert!(std::str::from_utf8(&output).unwrap().contains("{other}"));
    assert!(
        litchi_xlsx::sheet_protection::write_protection(&metadata, Conformance::Strict)
            .unwrap()
            .contains("purl.oclc.org/ooxml/spreadsheetml/main")
    );
}
