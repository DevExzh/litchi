use litchi_odt::footnote_separator::{Adjustment, Length, LineStyle, Percent, Separator, parse};
use litchi_odt::{Builder, Document};
use std::io::Cursor;

fn sample() -> Separator {
    Separator {
        width: Some(Length::new("0.018cm").unwrap()),
        relative_width: Some(Percent::new("25%").unwrap()),
        color: Some((0, 0, 0)),
        line_style: Some(LineStyle::Solid),
        adjustment: Some(Adjustment::Left),
        distance_before: Some(Length::new("0.101cm").unwrap()),
        distance_after: Some(Length::new("0.101cm").unwrap()),
    }
}

#[test]
fn parses_aliases_and_writes_a_deterministic_self_contained_fragment() {
    let xml = r##"<s:page-layout-properties xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0"><s:footnote-sep s:distance-after-sep="0.101cm" s:color="#000000" s:adjustment="left" s:line-style="solid" s:rel-width="25%" s:width="0.018cm" s:distance-before-sep="0.101cm"/></s:page-layout-properties>"##;
    let parsed = parse(xml).unwrap();
    assert_eq!(parsed, [sample()]);
    assert_eq!(
        parsed[0].to_xml_fragment().unwrap(),
        r##"<style:footnote-sep xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" style:width="0.018cm" style:rel-width="25%" style:color="#000000" style:line-style="solid" style:adjustment="left" style:distance-before-sep="0.101cm" style:distance-after-sep="0.101cm"/>"##
    );
}

#[test]
fn parses_bundled_libreoffice_flat_fixture() {
    let fixture = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/libreoffice-core/sw/qa/core/text/data/footnote-connect.fodt"
    ));
    let document =
        litchi_odt::generic::FlatOpenDocument::from_reader(Cursor::new(fixture)).unwrap();
    let separators = document.style_footnote_separators().unwrap();
    assert!(!separators.is_empty());
    assert!(separators.iter().any(|separator| separator == &sample()));
}

#[test]
fn builder_and_package_accessors_round_trip_typed_state() {
    let mut builder = Builder::new();
    builder
        .add_page_layout_footnote_separator("pm1", sample())
        .unwrap();
    let bytes = builder.build().unwrap();
    let package =
        litchi_odt::generic::OpenDocumentPackage::from_reader(Cursor::new(bytes)).unwrap();
    assert_eq!(package.style_footnote_separators().unwrap(), [sample()]);
}

#[test]
fn page_layout_parser_exposes_the_typed_separator() {
    let mut builder = Builder::new();
    builder
        .add_page_layout_footnote_separator("pm1", sample())
        .unwrap();
    let bytes = builder.build().unwrap();
    let document = Document::from_bytes(bytes).unwrap();
    let mutable = litchi_odt::mutable::MutableDocument::from_document(document).unwrap();
    let layouts = mutable.page_layouts().unwrap();
    assert_eq!(
        layouts[0]
            .properties
            .as_ref()
            .unwrap()
            .footnote_separator
            .as_ref(),
        Some(&sample())
    );
}

#[test]
fn rejects_wrong_scope_namespace_cardinality_content_and_values() {
    const PREFIX: &str = r#"<o:document-styles xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:x="urn:wrong"><o:automatic-styles><s:page-layout s:name="p"><s:page-layout-properties>"#;
    const SUFFIX: &str =
        "</s:page-layout-properties></s:page-layout></o:automatic-styles></o:document-styles>";
    for fragment in [
        r#"<x:footnote-sep/>"#,
        r#"<s:footnote-sep/><s:footnote-sep/>"#,
        r#"<s:footnote-sep>text</s:footnote-sep>"#,
        r#"<s:footnote-sep><s:footnote-sep/></s:footnote-sep>"#,
        r#"<s:footnote-sep x:width="1cm"/>"#,
        r#"<s:footnote-sep s:width="1em"/>"#,
        r#"<s:footnote-sep s:rel-width="25"/>"#,
        r##"<s:footnote-sep s:color="#xyzxyz"/>"##,
        r#"<s:footnote-sep s:line-style="double"/>"#,
        r#"<s:footnote-sep s:adjustment="start"/>"#,
    ] {
        let xml = format!("{PREFIX}{fragment}{SUFFIX}");
        assert!(parse(&xml).is_err(), "accepted {fragment}");
    }
    let misplaced = r#"<s:style xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0"><s:footnote-sep/></s:style>"#;
    assert!(parse(misplaced).is_err());
    let dtd = format!("<!DOCTYPE x>{PREFIX}<s:footnote-sep/>{SUFFIX}");
    assert!(parse(&dtd).is_err());
}
