use std::borrow::Cow;

use litchi_rtf::{PictureShapeProperties, RtfDocument, RtfWriter, ShapeProperty};

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn parses_real_libreoffice_metro_blob_and_reuses_its_payload() {
    let source = include_bytes!(
        "../../../3rdparty/libreoffice-core/sw/qa/extras/rtfexport/data/tdf167569-2.rtf"
    );
    let source = std::str::from_utf8(source).unwrap();
    let property_start = source.find(r#"{\sp{\sn metroBlob}"#).unwrap();
    let mut depth = 0usize;
    let property_end = source[property_start..]
        .char_indices()
        .find_map(|(offset, character)| {
            match character {
                '{' => depth += 1,
                '}' => {
                    depth -= 1;
                    if depth == 0 {
                        return Some(property_start + offset + 1);
                    }
                },
                _ => {},
            }
            None
        })
        .unwrap();
    let mut isolated = r#"{\rtf1{\shp{\*\shpinst"#.to_string();
    isolated.push_str(&source[property_start..property_end]);
    isolated.push_str("}}}");
    let producer = RtfDocument::parse(&isolated).unwrap();
    let property = producer
        .shapes()
        .iter()
        .flat_map(|shape| shape.properties.iter())
        .find(|property| property.name == "metroBlob")
        .unwrap();
    let payload = property.binary_value.as_deref().unwrap();
    assert!(payload.starts_with(b"PK\x03\x04"));
    assert!(payload.len() > 1_000);

    let mut document = RtfDocument::parse(
        r#"{\rtf1{\*\shppict{\pict\pngblip\picw1\pich1 89504e470d0a1a0a}}Body}"#,
    )
    .unwrap();
    let payload_ptr = payload.as_ptr();
    document
        .set_picture_shape_properties(
            0,
            Some(PictureShapeProperties {
                shape_id: Some(9),
                properties: vec![ShapeProperty::new_binary(
                    Cow::Borrowed("metroBlob"),
                    Cow::Borrowed(payload),
                )],
            }),
        )
        .unwrap();
    assert_eq!(
        document.pictures()[0]
            .shape_properties
            .as_ref()
            .unwrap()
            .properties[0]
            .binary_value
            .as_deref()
            .unwrap()
            .as_ptr(),
        payload_ptr
    );

    let output = write(&document);
    assert_eq!(
        document.pictures()[0]
            .shape_properties
            .as_ref()
            .unwrap()
            .properties[0]
            .binary_value
            .as_deref()
            .unwrap()
            .as_ptr(),
        payload_ptr
    );
    let serialized = String::from_utf8(output.clone()).unwrap();
    assert!(serialized.contains("{\\sp{\\sn metroBlob}{\\sv{\\*\\svb 504b0304"));

    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    let reparsed_property = &reparsed.pictures()[0]
        .shape_properties
        .as_ref()
        .unwrap()
        .properties[0];
    assert_eq!(reparsed_property.name, "metroBlob");
    assert!(reparsed_property.value.is_empty());
    assert_eq!(reparsed_property.binary_value.as_deref(), Some(payload));
    assert_eq!(reparsed.text(), "Body");
}

#[test]
fn round_trips_binary_background_property_canonically() {
    let mut document = RtfDocument::parse(r#"{\rtf1 Body}"#).unwrap();
    let mut shape = litchi_rtf::Shape::new(litchi_rtf::ShapeType::Rectangle);
    shape.properties.push(ShapeProperty::new_binary(
        Cow::Borrowed("pInkData"),
        Cow::Borrowed(&[0x00, 0x7f, 0x80, 0xff]),
    ));
    document.set_background_shape(shape).unwrap();

    let output = write(&document);
    assert!(
        String::from_utf8(output.clone())
            .unwrap()
            .contains("{\\sp{\\sn pInkData}{\\sv{\\*\\svb 007f80ff}}}")
    );
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    let property = reparsed
        .background_shape()
        .unwrap()
        .properties
        .iter()
        .find(|property| property.name == "pInkData")
        .unwrap();
    assert_eq!(
        property.binary_value.as_deref(),
        Some(&[0x00, 0x7f, 0x80, 0xff][..])
    );
}

#[test]
fn rejects_hostile_shape_binary_value_grammar() {
    for source in [
        r#"{\rtf1{\*\svb 00}}"#,
        r#"{\rtf1{\svb 00}}"#,
        r#"{\rtf1{\pict{\*\picprop{\sp{\sn x}{\sv{\svb 00}}}}\pngblip 89504e470d0a1a0a}}"#,
        r#"{\rtf1{\pict{\*\picprop{\sp{\sn x}{\sv{\*\svb1 00}}}}\pngblip 89504e470d0a1a0a}}"#,
        r#"{\rtf1{\pict{\*\picprop{\sp{\sn x}{\sv{\*\svb 00}{\*\svb 01}}}}\pngblip 89504e470d0a1a0a}}"#,
        r#"{\rtf1{\pict{\*\picprop{\sp{\sn x}{\sv scalar{\*\svb 00}}}}\pngblip 89504e470d0a1a0a}}"#,
        r#"{\rtf1{\pict{\*\picprop{\sp{\sn x}{\sv{\*\svb 00}scalar}}}\pngblip 89504e470d0a1a0a}}"#,
        r#"{\rtf1{\pict{\*\picprop{\sp{\sn x}{\sv{\*\svb }}}}\pngblip 89504e470d0a1a0a}}"#,
        r#"{\rtf1{\pict{\*\picprop{\sp{\sn x}{\sv{\*\svb 0}}}}\pngblip 89504e470d0a1a0a}}"#,
        r#"{\rtf1{\pict{\*\picprop{\sp{\sn x}{\sv{\*\svb gg}}}}\pngblip 89504e470d0a1a0a}}"#,
        r#"{\rtf1{\pict{\*\picprop{\sp{\sn x}{\sv{\*\svb{\object}}}}}\pngblip 89504e470d0a1a0a}}"#,
        r#"{\rtf1{\pict{\*\picprop{\sp{\sn{\*\svb 00}}{\sv 1}}}\pngblip 89504e470d0a1a0a}}"#,
        r#"{\rtf1{\pict{\*\picprop{\sp{\sn x}{\sv\svb 00}}}\pngblip 89504e470d0a1a0a}}"#,
    ] {
        assert!(
            RtfDocument::parse(source).is_err(),
            "accepted malformed {source}"
        );
    }
}

#[test]
fn enforces_typed_binary_property_invariants() {
    assert!(
        ShapeProperty::new_binary(Cow::Borrowed("x"), Cow::Borrowed(&[]))
            .validate()
            .is_err()
    );
    assert!(
        ShapeProperty::new_binary(Cow::Borrowed(""), Cow::Borrowed(&[1]))
            .validate()
            .is_err()
    );

    let mut mixed = ShapeProperty::new(Cow::Borrowed("x"), Cow::Borrowed("scalar"));
    mixed.binary_value = Some(Cow::Borrowed(&[1]));
    assert!(mixed.validate().is_err());
}
