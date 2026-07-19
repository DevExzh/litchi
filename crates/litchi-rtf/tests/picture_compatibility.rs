use litchi_rtf::{
    ImageType, PictureCompatibilityKind, PictureCompatibilityRecord, RtfDocument, RtfWriter,
};

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn parses_real_libreoffice_body_wrapper_and_rewrites_positionally() {
    let source = include_bytes!(
        "../../../3rdparty/libreoffice-core/sw/qa/extras/rtfexport/data/fdo76633.rtf"
    );
    let document = RtfDocument::parse_bytes(source).unwrap();
    let records = document.picture_compatibility_records();
    assert_eq!(records.len(), 1);
    assert_eq!(records[0].kind, PictureCompatibilityKind::ShapePicture);
    assert_eq!(records[0].position, 0);
    let picture = document
        .picture_for_compatibility_record(&records[0])
        .unwrap();
    assert!(std::ptr::eq(
        picture,
        &document.pictures()[records[0].picture_index]
    ));
    assert_eq!(picture.image_type, ImageType::Jpeg);

    let output = write(&document);
    assert!(String::from_utf8_lossy(&output).contains("{\\*\\shppict{\\pict\\jpegblip"));
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.picture_compatibility_records().len(), 1);
    assert_eq!(
        reparsed
            .picture_for_compatibility_record(&reparsed.picture_compatibility_records()[0])
            .unwrap()
            .data(),
        picture.data()
    );
}

#[test]
fn preferred_and_fallback_share_position_and_round_trip_in_order() {
    let source = concat!(
        r#"{\rtf1 A{\*\shppict{\pict\pngblip 89504e470d0a1a0a}}"#,
        r#"{\nonshppict{\pict\jpegblip ffd8ffe0}}B}"#,
    );
    let document = RtfDocument::parse(source).unwrap();
    assert_eq!(document.text(), "AB");
    assert_eq!(document.picture_compatibility_records().len(), 2);
    assert!(
        document
            .picture_compatibility_records()
            .iter()
            .all(|record| record.position == 1)
    );
    assert_eq!(
        document.picture_compatibility_records()[0].kind,
        PictureCompatibilityKind::ShapePicture
    );
    assert_eq!(
        document.picture_compatibility_records()[1].kind,
        PictureCompatibilityKind::NonShapePicture
    );

    let first = write(&document);
    let serialized = String::from_utf8_lossy(&first);
    assert!(serialized.find("\\shppict").unwrap() < serialized.find("\\nonshppict").unwrap());
    let reparsed = RtfDocument::parse_bytes(&first).unwrap();
    assert_eq!(
        reparsed.picture_compatibility_records(),
        document.picture_compatibility_records()
    );
    assert_eq!(write(&reparsed), first);
}

#[test]
fn typed_mutation_references_and_clears_without_deleting_picture() {
    let mut document = RtfDocument::parse(r#"{\rtf1 A{\pict\pngblip 89504e470d0a1a0a}B}"#).unwrap();
    document
        .push_picture_compatibility_record(PictureCompatibilityRecord {
            position: 1,
            kind: PictureCompatibilityKind::ShapePicture,
            picture_index: 0,
        })
        .unwrap();
    assert!(
        document
            .push_picture_compatibility_record(PictureCompatibilityRecord {
                position: 1,
                kind: PictureCompatibilityKind::ShapePicture,
                picture_index: 0,
            })
            .is_err()
    );
    assert!(
        document
            .push_picture_compatibility_record(PictureCompatibilityRecord {
                position: 99,
                kind: PictureCompatibilityKind::NonShapePicture,
                picture_index: 0,
            })
            .is_err()
    );
    assert_eq!(
        RtfDocument::parse_bytes(&write(&document))
            .unwrap()
            .pictures()
            .len(),
        1
    );

    document.clear_picture_compatibility_records();
    assert_eq!(document.pictures().len(), 1);
    assert!(document.picture_compatibility_records().is_empty());
}

#[test]
fn rejects_hostile_wrapper_grammar() {
    for source in [
        r#"{\rtf1{\shppict{\pict\pngblip 00}}}"#,
        r#"{\rtf1{\*\nonshppict{\pict\pngblip 00}}}"#,
        r#"{\rtf1{\*\shppict1{\pict\pngblip 00}}}"#,
        r#"{\rtf1{\nonshppict1{\pict\pngblip 00}}}"#,
        r#"{\rtf1{\*\shppict}}"#,
        r#"{\rtf1{\nonshppict text{\pict\pngblip 00}}}"#,
        r#"{\rtf1{\*\shppict{\pict}}}"#,
        r#"{\rtf1{\*\shppict{\pict\pngblip 00}{\pict\pngblip 00}}}"#,
        r#"{\rtf1{\*\shppict{\pict\pngblip 00}}{\*\shppict{\pict\pngblip 01}}}"#,
    ] {
        assert!(
            RtfDocument::parse(source).is_err(),
            "accepted malformed {source}"
        );
    }
}
