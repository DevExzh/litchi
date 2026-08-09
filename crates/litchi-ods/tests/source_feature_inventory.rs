use litchi_ods::{Builder, DrawingKind, SourceFeatureLimits, SourceFeatures, Spreadsheet};

const XML: &str = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0"><office:body><office:spreadsheet><table:table table:name="Data"><table:table-row><table:table-cell><text:p><text:a xlink:href="https://example.test/never-contact">external link</text:a></text:p></table:table-cell></table:table-row><calcext:conditional-formats><calcext:conditional-format/></calcext:conditional-formats><calcext:sparkline-groups><calcext:sparkline-group/></calcext:sparkline-groups><table:shapes><draw:frame draw:name="chart"><draw:image xlink:href="http://192.0.2.1/tracking-pixel.png"/></draw:frame></table:shapes></table:table></office:spreadsheet></office:body></office:document-content>"#;

#[test]
fn compact_source_feature_inventory_is_semantic_bounded_and_inert() {
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

#[test]
fn package_facade_exposes_the_same_inert_inventory() {
    let bytes = Builder::new().content_xml(XML).build().unwrap();
    let spreadsheet = Spreadsheet::from_bytes(bytes).unwrap();
    let inventory = spreadsheet.source_features().unwrap();
    let sheet = inventory.sheet("Data").unwrap();
    assert_eq!(sheet.hyperlinks()[0].text(), "external link");
    assert_eq!(sheet.drawings()[1].kind(), DrawingKind::Image);
}

#[test]
fn source_feature_boundary_rejects_entities_bad_context_and_limits() {
    assert!(SourceFeatures::parse("<!DOCTYPE x [<!ENTITY e 'x'>]>&e;").is_err());
    assert!(SourceFeatures::parse("<office:document-content/>").is_err());
    assert!(
        SourceFeatures::parse_with(
            XML,
            SourceFeatureLimits::default().with_input_bytes(XML.len() - 1)
        )
        .is_err()
    );
    assert!(
        SourceFeatures::parse_with(XML, SourceFeatureLimits::default().with_events(2)).is_err()
    );
}

#[test]
fn conditional_formats_and_sparklines_are_counted_per_sheet_under_exact_limits()
-> litchi_core::Result<()> {
    let xml = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0"><office:body><office:spreadsheet><table:table table:name="First"><calcext:conditional-formats><calcext:conditional-format/><calcext:conditional-format/></calcext:conditional-formats><calcext:sparkline-groups><calcext:sparkline-group/></calcext:sparkline-groups></table:table><table:table table:name="Second"><calcext:conditional-formats><calcext:conditional-format/></calcext:conditional-formats><calcext:sparkline-groups><calcext:sparkline-group/><calcext:sparkline-group/></calcext:sparkline-groups></table:table></office:spreadsheet></office:body></office:document-content>"#;
    let snapshot =
        SourceFeatures::parse_with(xml, SourceFeatureLimits::default().with_items_per_sheet(2))?;
    let first = snapshot
        .sheet("First")
        .ok_or_else(|| litchi_core::Error::InvalidFormat("missing First sheet".to_string()))?;
    let second = snapshot
        .sheet("Second")
        .ok_or_else(|| litchi_core::Error::InvalidFormat("missing Second sheet".to_string()))?;
    assert_eq!(first.conditional_format_count(), 2);
    assert_eq!(first.sparkline_group_count(), 1);
    assert_eq!(second.conditional_format_count(), 1);
    assert_eq!(second.sparkline_group_count(), 2);
    assert!(
        SourceFeatures::parse_with(xml, SourceFeatureLimits::default().with_items_per_sheet(1))
            .is_err()
    );
    Ok(())
}
