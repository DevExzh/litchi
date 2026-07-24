use std::borrow::Cow;

use litchi_rtf::{RtfDocument, RtfWriter, Shape, ShapeType};

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn parses_real_libreoffice_text_frames_and_round_trips_canonically() {
    let source = include_bytes!(
        "../../../test-data/libreoffice-core/sw/qa/extras/rtfexport/data/relsize.rtf"
    );
    let producer = RtfDocument::parse_bytes(source).unwrap();
    assert_eq!(producer.shapes().len(), 1);
    assert_eq!(producer.shapes()[0].text, "Textbox text.\n");
    assert!(producer.shapes()[0].text_destination_present);

    let mut document = RtfDocument::parse(r#"{\rtf1 Body}"#).unwrap();
    let shape = Shape::text_box(Cow::Borrowed("Line one\nTabbed\tUnicode 你"));
    document.set_background_shape(shape).unwrap();
    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    assert!(serialized.contains("{\\shptxt Line one\\par Tabbed\\tab Unicode "));
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    let text_frame = reparsed.background_shape().unwrap();
    assert_eq!(text_frame.text, "Line one\nTabbed\tUnicode 你");
    assert!(text_frame.text_destination_present);
    assert_eq!(reparsed.text(), "Body");
}

#[test]
fn preserves_empty_presence_and_supports_typed_set_and_clear() {
    let source = r#"{\rtf1{\shp{\*\shpinst{\sp{\sn shapeType}{\sv 202}}{\shptxt}}}}"#;
    let document = RtfDocument::parse(source).unwrap();
    assert!(document.shapes()[0].text.is_empty());
    assert!(document.shapes()[0].text_destination_present);

    let mut shape = Shape::new(ShapeType::TextBox);
    assert!(shape.set_text(Cow::Borrowed("value")).is_empty());
    assert!(shape.text_destination_present);
    assert_eq!(shape.clear_text(), "value");
    assert!(!shape.text_destination_present);
    assert!(shape.text.is_empty());
}

#[test]
fn retains_formatting_group_text_but_skips_nested_destinations_inertly() {
    let source = concat!(
        r#"{\rtf1 A{\shp{\*\shpinst{\sp{\sn shapeType}{\sv 202}}"#,
        r#"{\shptxt before{\b bold}{\*\unknown LEAK}"#,
        r#"{\pict\pngblip 89504e470d0a1a0a}after\par end}}}B}"#,
    );
    let document = RtfDocument::parse(source).unwrap();
    assert_eq!(document.text(), "AB");
    assert_eq!(document.shapes()[0].text, "beforeboldafter\nend");
    assert!(document.pictures().is_empty());
}

#[test]
fn rejects_hostile_shape_text_grammar_and_resource_abuse() {
    for source in [
        r#"{\rtf1{\shptxt x}}"#,
        r#"{\rtf1{\*\shptxt x}}"#,
        r#"{\rtf1{\shp{\*\shpinst{\shptxt1 x}}}}"#,
        r#"{\rtf1{\shp\shptxt x}}"#,
        r#"{\rtf1{\shp{{\shptxt x}}}}"#,
        r#"{\rtf1{\shp{\*\shpinst{\shptxt x}{\sp{\sn shapeType}{\sv 202}}}}}"#,
        r#"{\rtf1{\shp{\*\shpinst{\shptxt one}}{\shptxt two}}}"#,
        r#"{\rtf1{\shp{\*\shprslt fallback}{\shptxt late}}}"#,
        "{\\rtf1{\\shp{\\*\\shpinst{\\shptxt\\bin1 x}}}}",
    ] {
        assert!(
            RtfDocument::parse(source).is_err(),
            "accepted malformed {source}"
        );
    }

    let mut nested = r#"{\rtf1{\shp{\*\shpinst{\shptxt "#.to_string();
    nested.push_str(&"{".repeat(65));
    nested.push('x');
    nested.push_str(&"}".repeat(65));
    nested.push_str("}}}");
    assert!(RtfDocument::parse(&nested).is_err());

    let mut document = RtfDocument::parse(r#"{\rtf1}"#).unwrap();
    let mut shape = Shape::new(ShapeType::TextBox);
    shape.set_text(Cow::Owned("x".repeat(16 * 1_048_576 + 1)));
    assert!(document.set_background_shape(shape).is_err());
}
