use litchi_odt::{DocumentBuilder, OpenDocumentPackage};
use litchi_odt::section_properties::{
    BackgroundRepeat, SectionBackgroundColor, SectionBackgroundImage, SectionLength,
    SectionProperties, SectionStyleProperties, SectionWritingMode, parse_section_style_properties,
};

const PREFIX: &str = r#"<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:f="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:x="http://www.w3.org/1999/xlink"><office:styles>"#;
const SUFFIX: &str = "</office:styles></office:document-styles>";
fn document(body: &str) -> Vec<u8> {
    format!("{PREFIX}{body}{SUFFIX}").into_bytes()
}

#[test]
fn parses_all_values_aliases_and_round_trips() {
    let xml = document(
        r##"<s:style s:name="Sect" s:family="section"><s:section-properties f:background-color="#1a2B3c" f:margin-left="-0.5cm" f:margin-right="2in" s:editable="false" s:protect="true" s:writing-mode="tb-rl" t:dont-balance-text-columns="false"><s:background-image x:href="Pictures/a.png" s:repeat="no-repeat" s:position="top left" s:filter-name="soft" d:opacity="75%" x:type="simple" x:show="embed" x:actuate="onLoad"/></s:section-properties></s:style>"##,
    );
    let set = parse_section_style_properties(&xml).unwrap();
    let style = set.get("Sect").unwrap();
    assert_eq!(
        style.properties.background_color,
        Some(SectionBackgroundColor::Rgb(0x1a, 0x2b, 0x3c))
    );
    assert_eq!(
        style.properties.margin_left.as_ref().unwrap().as_str(),
        "-0.5cm"
    );
    assert_eq!(
        style.properties.writing_mode,
        Some(SectionWritingMode::TopToBottomRightToLeft)
    );
    assert_eq!(
        style.properties.background_image.as_ref().unwrap().repeat,
        Some(BackgroundRepeat::NoRepeat)
    );
    let fragment = style.to_xml_fragment().unwrap();
    let reparsed = parse_section_style_properties(&document(&fragment)).unwrap();
    assert_eq!(set, reparsed);
    assert_eq!(
        fragment,
        reparsed.get("Sect").unwrap().to_xml_fragment().unwrap()
    );
}

#[test]
fn builder_and_package_accessor() {
    let properties = SectionProperties {
        background_color: Some(SectionBackgroundColor::Transparent),
        margin_left: Some(SectionLength::new("0cm").unwrap()),
        writing_mode: Some(SectionWritingMode::Page),
        background_image: Some(SectionBackgroundImage {
            href: Some("Pictures/bg.png".into()),
            xlink_type: Some("simple".into()),
            actuate: Some("onLoad".into()),
            ..Default::default()
        }),
        ..Default::default()
    };
    let mut builder = DocumentBuilder::new();
    builder
        .add_section_property_style(SectionStyleProperties::new("S1", properties).unwrap())
        .unwrap();
    let package = builder.build().unwrap();
    let parsed = OpenDocumentPackage::from_bytes(package)
        .unwrap()
        .section_style_properties()
        .unwrap();
    assert!(parsed.get("S1").is_some());
}

#[test]
fn accepts_odfpy_odfdo_and_libreoffice_shapes() {
    let samples = [
        r#"<s:style s:name="odfpy" s:family="section"><s:section-properties f:background-color="transparent" f:margin-left="0cm" s:editable="false" t:dont-balance-text-columns="false"/></s:style>"#,
        r#"<s:style s:name="odfdo" s:family="section"><s:section-properties s:writing-mode="lr-tb"><s:background-image x:href="Pictures/x" x:type="simple" x:actuate="onLoad"/></s:section-properties></s:style>"#,
        r#"<s:style s:name="lo" s:family="section"><s:section-properties f:margin-left="0in" f:margin-right="0in"><s:background-image x:href="Pictures/100.png" x:type="simple" x:show="embed" x:actuate="onLoad" s:position="center center" s:repeat="no-repeat"/></s:section-properties></s:style>"#,
    ];
    for sample in samples {
        assert_eq!(
            parse_section_style_properties(&document(sample))
                .unwrap()
                .styles
                .len(),
            1
        );
    }
}

#[test]
fn enforces_namespace_order_cardinality_and_values() {
    let bad = [
        r#"<s:style s:name="x" s:family="section"><s:section-properties><s:columns/><s:background-image/></s:section-properties></s:style>"#,
        r#"<s:style s:name="x" s:family="section"><s:section-properties><s:background-image/><s:background-image/></s:section-properties></s:style>"#,
        r#"<s:style s:name="x" s:family="section"><s:section-properties><t:background-image/></s:section-properties></s:style>"#,
        r#"<s:style s:name="x" s:family="section"><s:section-properties><s:background-image><s:columns/></s:background-image></s:section-properties></s:style>"#,
        r#"<s:style s:name="x" s:family="section"><s:section-properties s:editable="yes"/></s:style>"#,
        r#"<s:style s:name="x" s:family="section"><s:section-properties f:background-color="red"/></s:style>"#,
        r#"<s:style s:name="x" s:family="section"><s:section-properties s:writing-mode="sideways"/></s:style>"#,
        r#"<s:style s:name="x" s:family="section"><s:section-properties><s:background-image s:repeat="tile"/></s:section-properties></s:style>"#,
    ];
    for sample in bad {
        assert!(
            parse_section_style_properties(&document(sample)).is_err(),
            "accepted {sample}"
        );
    }
}

#[test]
fn validates_existing_notes_and_rejects_dtd() {
    let bad_notes = document(
        r#"<s:style s:name="x" s:family="section"><s:section-properties><t:notes-configuration t:note-class="bogus"/></s:section-properties></s:style>"#,
    );
    assert!(parse_section_style_properties(&bad_notes).is_err());
    assert!(parse_section_style_properties(b"<!DOCTYPE x><x/>").is_err());
}
