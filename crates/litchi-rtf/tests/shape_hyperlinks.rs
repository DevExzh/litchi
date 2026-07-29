use litchi_rtf::{RtfDocument, RtfWriter, ShapeHyperlink};
use std::borrow::Cow;

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

const SHAPE_WITH_HYPERLINK: &str = concat!(
    r#"{\rtf1 A{\shp{\*\shpinst{\sp{\sn shapeType}{\sv 202}}"#,
    r#"{\sp{\sn hyperlink}{\sv }{\hl {\hlsrc src}{\hlloc http://example.test/x}{\hlfr Click me}}}"#,
    r#"{\shptxt x}}}B}"#,
);

#[test]
fn parses_shape_hyperlink_property_and_round_trips() {
    let document = RtfDocument::parse(SHAPE_WITH_HYPERLINK).unwrap();
    assert_eq!(document.text(), "AB");
    let shape = &document.shapes()[0];
    let property = shape
        .properties
        .iter()
        .find(|property| property.name == "hyperlink")
        .expect("hyperlink property");
    let hyperlink = property.hyperlink.as_ref().expect("hyperlink metadata");
    assert_eq!(hyperlink.location.as_deref(), Some("http://example.test/x"));
    assert_eq!(hyperlink.source.as_deref(), Some("src"));
    assert_eq!(hyperlink.friendly_name.as_deref(), Some("Click me"));

    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    assert!(serialized.contains("{\\hl{\\hlloc http://example.test/x}{\\hlsrc src}{\\hlfr Click me}}"));

    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.text(), document.text());
    let reparsed_property = reparsed.shapes()[0]
        .properties
        .iter()
        .find(|property| property.name == "hyperlink")
        .unwrap();
    assert_eq!(reparsed_property.hyperlink, property.hyperlink);
}

#[test]
fn accepts_partial_hyperlink_strings() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1 A{\shp{\*\shpinst{\sp{\sn shapeType}{\sv 202}}"#,
        r#"{\sp{\sn link}{\sv 0}{\hl {\hlfr only a name}}}{\shptxt x}}}B}"#,
    ))
    .unwrap();
    let property = document.shapes()[0]
        .properties
        .iter()
        .find(|property| property.name == "link")
        .unwrap();
    let hyperlink = property.hyperlink.as_ref().unwrap();
    assert!(hyperlink.location.is_none());
    assert!(hyperlink.source.is_none());
    assert_eq!(hyperlink.friendly_name.as_deref(), Some("only a name"));

    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
    assert_eq!(
        reparsed.shapes()[0].properties[1].hyperlink,
        property.hyperlink
    );
}

#[test]
fn typed_validation_requires_at_least_one_string() {
    assert!(ShapeHyperlink::default().validate().is_err());
    let hyperlink = ShapeHyperlink {
        location: Some(Cow::Borrowed("http://example.test")),
        ..ShapeHyperlink::default()
    };
    assert!(hyperlink.validate().is_ok());
}

#[test]
fn rejects_malformed_shape_hyperlink_groups() {
    let cases = [
        // Empty hl group.
        r#"{\rtf1{\shp{\*\shpinst{\sp{\sn a}{\sv 1}{\hl }}}}}"#,
        // Duplicate string destination.
        r#"{\rtf1{\shp{\*\shpinst{\sp{\sn a}{\sv 1}{\hl {\hlloc x}{\hlloc y}}}}}}"#,
        // Starred hl destination.
        r#"{\rtf1{\shp{\*\shpinst{\sp{\sn a}{\sv 1}{\*\hl {\hlloc x}}}}}}"#,
        // hl before the property name.
        r#"{\rtf1{\shp{\*\shpinst{\sp{\hl {\hlloc x}}{\sn a}{\sv 1}}}}}"#,
        // Second hl group in one property.
        r#"{\rtf1{\shp{\*\shpinst{\sp{\sn a}{\sv 1}{\hl {\hlloc x}}{\hl {\hlsrc y}}}}}}"#,
        // Bare string control outside a group.
        r#"{\rtf1{\shp{\*\shpinst{\sp{\sn a}{\sv 1}\hlloc}}}}"#,
        // String destination outside hl.
        r#"{\rtf1{\shp{\*\shpinst{\sp{\sn a}{\sv 1}{\hlloc x}}}}}"#,
        // Unsupported group inside hl.
        r#"{\rtf1{\shp{\*\shpinst{\sp{\sn a}{\sv 1}{\hl {\sv x}}}}}}"#,
        // Ungrouped text inside hl.
        r#"{\rtf1{\shp{\*\shpinst{\sp{\sn a}{\sv 1}{\hl text}}}}}"#,
        // Unterminated hl group.
        r#"{\rtf1{\shp{\*\shpinst{\sp{\sn a}{\sv 1}{\hl {\hlloc x}}}}"#,
    ];
    for rtf in cases {
        assert!(RtfDocument::parse(rtf).is_err(), "accepted malformed {rtf}");
    }
}
