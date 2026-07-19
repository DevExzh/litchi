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
fn parses_and_round_trips_real_libreoffice_picture_properties() {
    let source = include_bytes!(
        "../../../3rdparty/libreoffice-core/sw/qa/extras/rtfimport/data/fdo85179.rtf"
    );
    let document = RtfDocument::parse_bytes(source).unwrap();
    assert_eq!(document.pictures().len(), 1);
    let picture = &document.pictures()[0];
    let properties = picture.shape_properties.as_ref().unwrap();
    assert_eq!(properties.shape_id, Some(1025));
    assert_eq!(properties.properties.len(), 4);
    assert_eq!(properties.properties[0].name, "shapeType");
    assert_eq!(properties.properties[0].value, "75");
    assert_eq!(properties.properties[1].name, "lineColor");
    assert_eq!(properties.properties[1].value, "65535");

    let data_ptr = picture.data().as_ptr();
    let output = write(&document);
    assert_eq!(document.pictures()[0].data().as_ptr(), data_ptr);
    let serialized = String::from_utf8(output.clone()).unwrap();
    assert!(serialized.contains("{\\*\\picprop\\shplid1025"));
    assert!(serialized.contains("{\\sp{\\sn shapeType}{\\sv 75}}"));

    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.text(), document.text());
    assert_eq!(
        reparsed.pictures()[0].shape_properties,
        document.pictures()[0].shape_properties
    );
    assert_eq!(reparsed.pictures()[0].data(), document.pictures()[0].data());
}

#[test]
fn typed_mutation_moves_metadata_without_cloning_picture_payload() {
    let mut document = RtfDocument::parse(
        r#"{\rtf1{\*\shppict{\pict\pngblip\picw1\pich1 89504e470d0a1a0a}}Body}"#,
    )
    .unwrap();
    let data_ptr = document.pictures()[0].data().as_ptr();
    let properties = PictureShapeProperties {
        shape_id: Some(-7),
        properties: vec![
            ShapeProperty::new(
                Cow::Borrowed("futureProperty"),
                Cow::Borrowed("opaque value"),
            ),
            ShapeProperty::new(Cow::Borrowed("fFlipH"), Cow::Borrowed("1")),
        ],
    };
    assert!(
        document
            .set_picture_shape_properties(0, Some(properties))
            .unwrap()
            .is_none()
    );
    assert_eq!(document.pictures()[0].data().as_ptr(), data_ptr);

    let output = write(&document);
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    let reparsed_properties = reparsed.pictures()[0].shape_properties.as_ref().unwrap();
    assert_eq!(reparsed_properties.shape_id, Some(-7));
    assert_eq!(reparsed_properties.properties[0].name, "futureProperty");
    assert_eq!(reparsed_properties.properties[0].value, "opaque value");

    let removed = document.set_picture_shape_properties(0, None).unwrap();
    assert!(removed.is_some());
    assert_eq!(document.pictures()[0].data().as_ptr(), data_ptr);
    assert!(
        !String::from_utf8(write(&document))
            .unwrap()
            .contains("\\picprop")
    );
    assert!(document.set_picture_shape_properties(1, None).is_err());
}

#[test]
fn rejects_hostile_picture_property_grammar() {
    let suffix = r#"\pngblip 89504e470d0a1a0a}"#;
    for prefix in [
        r#"{\rtf1{\pict{\picprop{\sp{\sn x}{\sv 1}}}"#,
        r#"{\rtf1{\pict{\*\picprop1{\sp{\sn x}{\sv 1}}}"#,
        r#"{\rtf1{\pict{\*\picprop}"#,
        r#"{\rtf1{\pict{\*\picprop\shplid{\sp{\sn x}{\sv 1}}}"#,
        r#"{\rtf1{\pict{\*\picprop\shplid1\shplid2{\sp{\sn x}{\sv 1}}}"#,
        r#"{\rtf1{\pict{\*\picprop{\sp{\sn x}{\sv 1}}\shplid2}"#,
        r#"{\rtf1{\pict{\*\picprop{\sp{\sv 1}{\sn x}}}"#,
        r#"{\rtf1{\pict{\*\picprop{\sp{\sn x}}}"#,
        r#"{\rtf1{\pict{\*\picprop{\sp1{\sn x}{\sv 1}}}"#,
        r#"{\rtf1{\pict{\*\picprop{\sp{\sn1 x}{\sv 1}}}"#,
        r#"{\rtf1{\pict{\*\picprop{\sp{\sn x{\object}}{\sv 1}}}"#,
        r#"{\rtf1{\pict\pngblip 8950{\*\picprop{\sp{\sn x}{\sv 1}}}"#,
    ] {
        let source = format!("{prefix}{suffix}");
        assert!(
            RtfDocument::parse(&source).is_err(),
            "accepted malformed {source}"
        );
    }

    for source in [
        r#"{\rtf1{\*\picprop{\sp{\sn x}{\sv 1}}}}"#,
        r#"{\rtf1{\picprop{\sp{\sn x}{\sv 1}}}}"#,
        r#"{\rtf1{\pict{\*\picprop{\sp{\sn x}{\sv 1}}}{\*\picprop{\sp{\sn y}{\sv 2}}}\pngblip 89504e470d0a1a0a}}"#,
        r#"{\rtf1{\pict{{\*\picprop{\sp{\sn x}{\sv 1}}}}\pngblip 89504e470d0a1a0a}}"#,
    ] {
        assert!(
            RtfDocument::parse(source).is_err(),
            "accepted malformed {source}"
        );
    }
}

#[test]
fn enforces_typed_property_bounds() {
    let mut document =
        RtfDocument::parse(r#"{\rtf1{\*\shppict{\pict\pngblip 89504e470d0a1a0a}}}"#).unwrap();
    assert!(
        document
            .set_picture_shape_properties(0, Some(PictureShapeProperties::default()))
            .is_err()
    );

    let invalid = PictureShapeProperties {
        shape_id: None,
        properties: vec![ShapeProperty::new(Cow::Borrowed(""), Cow::Borrowed("1"))],
    };
    assert!(
        document
            .set_picture_shape_properties(0, Some(invalid))
            .is_err()
    );
}
