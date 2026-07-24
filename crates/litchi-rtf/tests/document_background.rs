use std::borrow::Cow;

use litchi_rtf::{RtfDocument, RtfWriter, Shape, ShapeProperty, ShapeType};

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn parses_and_rewrites_real_producer_background_with_provenance() {
    let source = include_bytes!(
        "../../../test-data/libreoffice-core/sw/qa/extras/rtfexport/data/page-background.rtf"
    );
    let document = RtfDocument::parse_bytes(source).unwrap();
    let background = document.background_shape().unwrap();
    assert_eq!(document.shapes().len(), 1);
    assert_eq!(background.shape_type, ShapeType::Rectangle);
    assert_eq!(background.fill.color.raw(), 5_296_274);
    assert_eq!(background.property("bWMode"), Some("9"));
    assert!(background.is_background);

    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    assert!(serialized.contains("{\\*\\background{\\shp{\\*\\shpinst"));
    assert!(serialized.contains("{\\sn bWMode}{\\sv 9}"));
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(
        reparsed.background_shape().unwrap().property("bWMode"),
        Some("9")
    );
    assert_eq!(reparsed.text(), document.text());
}

#[test]
fn typed_api_round_trips_unknown_scalar_properties_and_clears() {
    let mut document = RtfDocument::parse(r#"{\rtf1 Body}"#).unwrap();
    let mut shape = Shape::new(ShapeType::Ellipse);
    shape.geometry.x = 10;
    shape.geometry.y = 20;
    shape.geometry.width = 300;
    shape.geometry.height = 400;
    shape.properties.push(ShapeProperty::new(
        Cow::Borrowed("futureOfficeArtProperty"),
        Cow::Borrowed("opaque {value} \\ untouched"),
    ));
    document.set_background_shape(shape).unwrap();

    let output = write(&document);
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    let background = reparsed.background_shape().unwrap();
    assert_eq!(background.shape_type, ShapeType::Ellipse);
    assert_eq!(background.geometry.width, 300);
    assert_eq!(background.geometry.height, 400);
    assert_eq!(
        background.property("futureOfficeArtProperty"),
        Some("opaque {value} \\ untouched")
    );
    assert_eq!(reparsed.text(), "Body");

    assert!(document.clear_background_shape().is_some());
    assert!(document.background_shape().is_none());
    assert!(
        !String::from_utf8(write(&document))
            .unwrap()
            .contains("\\background")
    );
}

#[test]
fn rejects_malformed_destination_grammar_and_duplicates() {
    for source in [
        r#"{\rtf1{\background{\shp}}}"#,
        r#"{\rtf1{\*\background1{\shp}}}"#,
        r#"{\rtf1{\*\background}}"#,
        r#"{\rtf1{\*\background text{\shp}}}"#,
        r#"{\rtf1{\*\background{\*\shp}}}"#,
        r#"{\rtf1{\*\background{\shp}{\shp}}}"#,
        r#"{\rtf1{\*\background{\shp{\sp{\sn }{\sv 1}}}}}"#,
        r#"{\rtf1{\*\background{\shp}}{\*\background{\shp}}}"#,
        r#"{\rtf1 Body{\*\background{\shp}}}"#,
    ] {
        assert!(
            RtfDocument::parse(source).is_err(),
            "accepted malformed {source}"
        );
    }
}

#[test]
fn enforces_typed_resource_and_geometry_bounds() {
    let mut document = RtfDocument::parse(r#"{\rtf1}"#).unwrap();
    let mut shape = Shape::new(ShapeType::Rectangle);
    shape.geometry.x = i32::MAX;
    shape.geometry.width = 1;
    assert!(document.set_background_shape(shape).is_err());

    let mut shape = Shape::new(ShapeType::Rectangle);
    shape.properties.push(ShapeProperty::new(
        Cow::Borrowed(""),
        Cow::Borrowed("value"),
    ));
    assert!(document.set_background_shape(shape).is_err());
}
