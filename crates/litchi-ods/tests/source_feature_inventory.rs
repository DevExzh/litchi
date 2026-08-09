use litchi_ods::{DrawingKind, SourceFeatures};

const XML: &str = r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0"><office:body><office:spreadsheet><table:table table:name="Data"><table:table-row><table:table-cell><text:p><text:a xlink:href="https://example.test/never-contact">external link</text:a></text:p></table:table-cell></table:table-row><calcext:conditional-format/><calcext:sparkline-group/><table:shapes><draw:frame draw:name="chart"><draw:image xlink:href="http://192.0.2.1/tracking-pixel.png"/></draw:frame></table:shapes></table:table></office:spreadsheet></office:body></office:document>"#;

#[test]
fn source_feature_inventory_is_semantic_bounded_and_inert() {
    let features = SourceFeatures::parse(XML).unwrap();
    let sheet = features.sheet("Data").unwrap();
    assert_eq!(sheet.conditional_format_count(), 1);
    assert_eq!(sheet.sparkline_group_count(), 1);
    assert_eq!(
        sheet.hyperlinks()[0].href(),
        "https://example.test/never-contact"
    );
    assert_eq!(sheet.hyperlinks()[0].text(), "external link");
    assert_eq!(sheet.drawings().len(), 2);
    assert_eq!(sheet.drawings()[0].kind(), DrawingKind::Frame);
    assert_eq!(
        sheet.drawings()[1].href(),
        Some("http://192.0.2.1/tracking-pixel.png")
    );
}
