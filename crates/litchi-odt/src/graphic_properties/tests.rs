//! Regression fixtures for the complete graphic-property owner.

use super::*;
const HEAD: &str = r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><office:automatic-styles>"#;
fn doc(body: &str) -> String {
    format!("{HEAD}{body}</office:automatic-styles></office:document>")
}
#[test]
fn every_normative_property_kind_round_trips() {
    let mut value = GraphicStyleProperties::default();
    for kind in GraphicPropertyKind::ALL {
        value.set_lexical(*kind, kind.example()).unwrap();
    }
    value.set_child(
        GraphicPropertyChild::new(
            GraphicPropertyChildKind::ListStyle,
            r#"<text:list-style style:name="L"/>"#,
        )
        .unwrap(),
    );
    value.set_child(
        GraphicPropertyChild::new(
            GraphicPropertyChildKind::BackgroundImage,
            r#"<style:background-image xlink:href="Pictures/a&amp;b.png"/>"#,
        )
        .unwrap(),
    );
    value.set_child(
        GraphicPropertyChild::new(
            GraphicPropertyChildKind::Columns,
            r#"<style:columns fo:column-count="2" fo:column-gap="0cm"/>"#,
        )
        .unwrap(),
    );
    let fragment = value.to_xml_fragment().unwrap();
    assert_eq!(
        GraphicStyleProperties::from_xml_fragment(&fragment).unwrap(),
        value
    );
    assert_eq!(value.iter().count(), 174)
}
#[test]
fn parses_real_odfdo_fixture_values() {
    let fixture = include_str!("../../../../test-data/odfdo/tests/samples/images.fodt");
    let begin = fixture
        .find(r#"<style:style style:name="fr1" style:family="graphic""#)
        .unwrap();
    let end = begin + fixture[begin..].find("</style:style>").unwrap() + "</style:style>".len();
    let set = parse_graphic_style_properties(&doc(&fixture[begin..end])).unwrap();
    let value = set.get("fr1").unwrap().properties.as_ref().unwrap();
    assert_eq!(
        value.get(GraphicPropertyKind::DrawGamma).unwrap().lexical(),
        "100%"
    );
    assert_eq!(
        value.get(GraphicPropertyKind::FoClip).unwrap().lexical(),
        "rect(0cm, 0cm, 0cm, 0cm)"
    )
}
#[test]
fn lossless_replace_insert_remove() {
    let original = doc(
        "<!--keep--><style:style style:name=\"a\" style:family=\"graphic\"><x:k xmlns:x=\"urn:k\"/></style:style><style:style style:name=\"b\" style:family=\"graphic\"><style:graphic-properties draw:fill=\"none\"/></style:style>",
    );
    let mut properties = GraphicStyleProperties::default();
    properties
        .set_lexical(GraphicPropertyKind::DrawFill, "solid")
        .unwrap();
    let mut a = GraphicStyleRecord::named("a", Some(properties)).unwrap();
    let inserted = set_graphic_style_properties_xml(&original, &a).unwrap();
    assert!(inserted.contains("<x:k xmlns:x=\"urn:k\"/><style:graphic-properties"));
    a.properties = None;
    let restored = set_graphic_style_properties_xml(&inserted, &a).unwrap();
    assert_eq!(restored, original);
    let removed =
        set_graphic_style_properties_xml(&restored, &GraphicStyleRecord::named("b", None).unwrap())
            .unwrap();
    assert!(!removed.contains("draw:fill=\"none\""));
    assert!(removed.contains("<!--keep-->"))
}
#[test]
fn rejects_malformed_and_capped_input() {
    for body in [
        r#"<style:style style:name="a" style:family="graphic"><style:graphic-properties draw:fill="wrong"/></style:style>"#,
        r#"<style:style style:name="a" style:family="graphic"><style:graphic-properties draw:fill="none" draw:fill="solid"/></style:style>"#,
        r#"<style:style style:name="a" style:family="graphic"><draw:graphic-properties/></style:style>"#,
        r#"<style:style style:name="a" style:family="graphic"><style:graphic-properties><draw:columns/></style:graphic-properties></style:style>"#,
        r#"<style:style style:name="a" style:family="graphic"><style:graphic-properties draw:tile-repeat-offset="101% horizontal"/></style:style>"#,
    ] {
        assert!(
            parse_graphic_style_properties(&doc(body)).is_err(),
            "accepted {body}"
        )
    }
    let huge = "x".repeat(MAX_VALUE + 1);
    assert!(GraphicProperty::new(GraphicPropertyKind::DrawStrokeDash, &huge).is_err());
    let mut attributes = String::new();
    for index in 0..=MAX_ATTRIBUTES {
        attributes.push_str(&format!(" x:a{index}=\"1\""))
    }
    let xml = doc(&format!(
        r#"<style:style style:name="a" style:family="graphic"><style:graphic-properties xmlns:x="urn:x"{attributes}/></style:style>"#
    ));
    assert!(parse_graphic_style_properties(&xml).is_err())
}
