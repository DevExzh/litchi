use litchi_odt::header_footer_properties::{
    HeaderFooterBorder, HeaderFooterBorderStyle, HeaderFooterColor, HeaderFooterLength,
    HeaderFooterShadow, HeaderFooterStyleProperties, PageHeaderFooterRegion,
    parse_page_layout_header_footer_properties,
};
use litchi_odt::section_properties::BackgroundRepeat;
use litchi_odt::{Builder, Document};

const PREFIX: &str = r#"<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:f="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:v="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:x="http://www.w3.org/1999/xlink"><office:automatic-styles><s:page-layout s:name="pm1">"#;
fn doc(body: &str) -> String {
    format!("{PREFIX}{body}</s:page-layout></office:automatic-styles></office:document-styles>")
}

#[test]
fn parses_all_standard_attributes_and_round_trips() {
    let xml = doc(
        r##"<s:header-style><s:header-footer-properties v:height="1cm" f:min-height="0.5cm" f:margin="0" f:margin-left="-1mm" f:border="1.5pt solid #000000" f:border-bottom="none" s:border-line-width="0.1pt 0.2pt 0.1pt" f:padding="0.01in" f:background-color="#c0c0c0" s:shadow="#808080 0.1cm 0.2cm" s:dynamic-spacing="false"><s:background-image x:href="Pictures/bg.png" s:repeat="no-repeat" x:type="simple" x:actuate="onLoad"/></s:header-footer-properties></s:header-style>"##,
    );
    let parsed = parse_page_layout_header_footer_properties(&xml).unwrap();
    let p = &parsed[0].properties;
    assert_eq!(
        p.background_color,
        Some(HeaderFooterColor::Rgb(192, 192, 192))
    );
    assert!(matches!(
        p.borders.all,
        Some(HeaderFooterBorder::Line {
            style: HeaderFooterBorderStyle::Solid,
            ..
        })
    ));
    assert!(matches!(p.shadow, Some(HeaderFooterShadow::Drop { .. })));
    assert_eq!(
        p.background_image.as_ref().unwrap().repeat,
        Some(BackgroundRepeat::NoRepeat)
    );
    let fragment = p
        .to_region_fragment(PageHeaderFooterRegion::Header)
        .unwrap();
    let reparsed = parse_page_layout_header_footer_properties(&doc(&fragment)).unwrap();
    assert_eq!(parsed[0].properties, reparsed[0].properties);
    assert_eq!(
        fragment,
        reparsed[0]
            .properties
            .to_region_fragment(PageHeaderFooterRegion::Header)
            .unwrap()
    );
}

#[test]
fn parses_odfdo_and_libreoffice_real_fixtures() {
    for path in [
        "../../test-data/odfdo/tests/samples/test_flat_lo.fods",
        "../../test-data/libreoffice-core/xmloff/qa/unit/data/scale-width-redline.fodt",
    ] {
        let source =
            std::fs::read_to_string(std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join(path))
                .unwrap();
        let parsed = parse_page_layout_header_footer_properties(&source).unwrap();
        assert!(!parsed.is_empty(), "{path}");
    }
}

#[test]
fn malformed_namespaces_order_cardinality_and_caps() {
    for body in [
        r#"<s:header-style><s:header-footer-properties s:min-height="1cm"/></s:header-style>"#,
        r#"<s:header-style><s:header-footer-properties s:dynamic-spacing="yes"/></s:header-style>"#,
        r#"<s:header-style><s:header-footer-properties f:border="wide solid #000000"/></s:header-style>"#,
        r#"<s:header-style><s:header-footer-properties f:padding="-1cm"/></s:header-style>"#,
        r#"<s:header-style><s:header-footer-properties><s:background-image/><s:background-image/></s:header-footer-properties></s:header-style>"#,
        r#"<s:header-style><s:header-footer-properties><s:columns/></s:header-footer-properties></s:header-style>"#,
    ] {
        assert!(
            parse_page_layout_header_footer_properties(&doc(body)).is_err(),
            "accepted {body}"
        );
    }
    assert!(parse_page_layout_header_footer_properties("<!DOCTYPE x><x/>").is_err());
}

#[test]
fn public_construction_is_bounded() {
    assert!(HeaderFooterLength::new("1e9cm").is_err());
    let p = HeaderFooterStyleProperties {
        min_height: Some(HeaderFooterLength::new("1cm").unwrap()),
        ..Default::default()
    };
    assert!(
        p.to_xml_fragment()
            .unwrap()
            .contains("fo:min-height=\"1cm\"")
    );
}

#[test]
fn builder_package_and_mutable_paths_preserve_siblings() {
    let header = HeaderFooterStyleProperties {
        min_height: Some(HeaderFooterLength::new("1cm").unwrap()),
        dynamic_spacing: Some(true),
        ..Default::default()
    };
    let footer = HeaderFooterStyleProperties {
        min_height: Some(HeaderFooterLength::new("2cm").unwrap()),
        background_color: Some(HeaderFooterColor::Transparent),
        ..Default::default()
    };
    let mut builder = Builder::new();
    builder.add_paragraph("Body").unwrap();
    builder
        .add_page_layout_header_footer_properties("pm1", PageHeaderFooterRegion::Header, header)
        .unwrap();
    builder
        .add_page_layout_header_footer_properties(
            "pm1",
            PageHeaderFooterRegion::Footer,
            footer.clone(),
        )
        .unwrap();
    let bytes = builder.build().unwrap();
    let package = litchi_odt::generic::OpenDocumentPackage::from_bytes(bytes.clone()).unwrap();
    let entries = package.page_layout_header_footer_properties().unwrap();
    assert_eq!(entries.len(), 2);
    let styles = package.styles_xml().unwrap().unwrap();
    assert_eq!(
        styles
            .matches("<style:page-layout style:name=\"pm1\">")
            .count(),
        1
    );
    assert!(styles.contains("Numbered list style"));
    let document = Document::from_bytes(bytes).unwrap();
    let mut mutable = litchi_odt::mutable::MutableDocument::from_document(document).unwrap();
    let replacement = HeaderFooterStyleProperties {
        min_height: Some(HeaderFooterLength::new("3cm").unwrap()),
        shadow: Some(HeaderFooterShadow::None),
        ..Default::default()
    };
    mutable
        .set_page_layout_header_footer_properties(
            "pm1",
            PageHeaderFooterRegion::Header,
            &replacement,
        )
        .unwrap();
    let output =
        litchi_odt::generic::OpenDocumentPackage::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    let entries = output.page_layout_header_footer_properties().unwrap();
    assert_eq!(
        entries
            .iter()
            .find(|entry| entry.region == PageHeaderFooterRegion::Header)
            .unwrap()
            .properties
            .min_height
            .as_ref()
            .unwrap()
            .as_str(),
        "3cm"
    );
    assert_eq!(
        entries
            .iter()
            .find(|entry| entry.region == PageHeaderFooterRegion::Footer)
            .unwrap()
            .properties,
        footer
    );
    assert!(
        output
            .styles_xml()
            .unwrap()
            .unwrap()
            .contains("Numbered list style")
    );
}
