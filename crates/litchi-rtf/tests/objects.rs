use std::borrow::Cow;

use litchi_rtf::{EmbeddedObject, ObjectKind, ObjectResultKind, RtfDocument, RtfWriter};

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn round_trips_real_libreoffice_inline_ole_object() {
    let source = include_bytes!(
        "../../../test-data/libreoffice-core/sw/qa/extras/rtfimport/data/ole-inline.rtf"
    );
    let document = RtfDocument::parse_bytes(source).unwrap();
    assert_eq!(document.objects().len(), 1);
    let object = &document.objects()[0];
    assert_eq!(object.kind, ObjectKind::Embedded);
    assert!(!object.class_name.is_empty());
    assert!(!object.data.is_empty());
    assert!(document.text().get(..object.position).is_some());

    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    assert!(serialized.contains("{\\object\\objemb"));
    assert!(serialized.contains("{\\*\\objclass "));
    assert!(serialized.contains("{\\*\\objdata "));

    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.text(), document.text());
    assert_eq!(reparsed.objects().len(), 1);
    let reparsed_object = &reparsed.objects()[0];
    assert_eq!(reparsed_object.position, object.position);
    assert_eq!(reparsed_object.kind, object.kind);
    assert_eq!(reparsed_object.class_name, object.class_name);
    assert_eq!(reparsed_object.width, object.width);
    assert_eq!(reparsed_object.height, object.height);
    assert_eq!(reparsed_object.data, object.data);
}

#[test]
fn preserves_typed_object_metadata_and_shared_result_picture_indices() {
    let source = r#"{\rtf1\ansi A{\object\objocx\linkself\objlock\objupdate
        {\*\objclass Control.Class}{\*\objname Visible Name}\objsetsize
        \objalign20\objtransy-3\objh40\objw50\objcropt1\objcropb2
        \objcropl3\objcropr4\objscalex125\objscaley80\rsltmerge\rslthtml
        {\*\oleclsid 00112233-4455-6677-8899-AABBCCDDEEFF}
        {\*\objdata 0102a0ff}
        {\result fallback{\pict\pngblip\picw1\pich1 89504e470d0a1a0a}}}B}"#;
    let document = RtfDocument::parse(source).unwrap();
    let object = &document.objects()[0];
    assert_eq!(object.position, 1);
    assert_eq!(object.kind, ObjectKind::OleControl);
    assert!(object.link_self);
    assert_eq!(object.class_id, "00112233-4455-6677-8899-AABBCCDDEEFF");
    assert_eq!(object.alignment, Some(20));
    assert_eq!(object.translation_y, Some(-3));
    assert_eq!(object.crop_top, Some(1));
    assert_eq!(object.crop_bottom, Some(2));
    assert_eq!(object.crop_left, Some(3));
    assert_eq!(object.crop_right, Some(4));
    assert_eq!(object.scale_x, Some(125));
    assert_eq!(object.scale_y, Some(80));
    assert!(object.merge_result);
    assert_eq!(object.result_kind, Some(ObjectResultKind::Html));
    assert_eq!(object.data, [1, 2, 0xa0, 0xff]);
    assert_eq!(object.result_picture_indices.len(), 1);
    let picture_index = object.result_picture_indices[0];
    assert!(std::ptr::eq(
        document.picture_for_object_result(object, 0).unwrap(),
        &document.pictures()[picture_index]
    ));

    let output = write(&document);
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    let reparsed_object = &reparsed.objects()[0];
    assert_eq!(reparsed_object.position, 1);
    assert_eq!(reparsed_object.kind, ObjectKind::OleControl);
    assert_eq!(reparsed_object.class_id, object.class_id);
    assert_eq!(reparsed_object.alignment, Some(20));
    assert_eq!(reparsed_object.result_kind, Some(ObjectResultKind::Html));
    assert_eq!(reparsed_object.data, object.data);
    assert_eq!(reparsed_object.result_text, "fallback");
    assert_eq!(reparsed_object.result_picture_indices.len(), 1);
}

#[test]
fn typed_mutation_validates_positions_and_clear_keeps_shared_pictures() {
    let mut document = RtfDocument::parse(
        r#"{\rtf1 Body{\object\objemb{\*\objdata 00}{\result{\pict\pngblip\picw1\pich1 89504e470d0a1a0a}}}}"#,
    )
    .unwrap();
    assert_eq!(document.pictures().len(), 1);
    document.clear_objects();
    assert_eq!(document.pictures().len(), 1);

    let mut object = EmbeddedObject::new();
    object.position = 2;
    object.kind = ObjectKind::Publisher;
    object.class_name = Cow::Borrowed("Publisher.Class");
    object.data = vec![0x12, 0x34];
    object.result_picture_indices.push(0);
    document.push_object(object.clone()).unwrap();
    assert!(
        document
            .picture_for_object_result(&document.objects()[0], 0)
            .is_some()
    );

    let mut invalid = object;
    invalid.position = 99;
    assert!(document.push_object(invalid).is_err());

    let output = write(&document);
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.objects()[0].position, 2);
    assert_eq!(reparsed.objects()[0].kind, ObjectKind::Publisher);
    assert_eq!(reparsed.objects()[0].result_picture_indices.len(), 1);

    document.clear_objects();
    assert_eq!(document.pictures().len(), 1);
    assert!(
        !String::from_utf8(write(&document))
            .unwrap()
            .contains("\\object")
    );
}

#[test]
fn rejects_hostile_object_destination_grammar() {
    for source in [
        r#"{\rtf1{\*\object\objemb{\*\objdata 00}}}"#,
        r#"{\rtf1{\object1\objemb{\*\objdata 00}}}"#,
        r#"{\rtf1{\object\objemb1{\*\objdata 00}}}"#,
        r#"{\rtf1{\object\objemb\objalign{\*\objdata 00}}}"#,
        r#"{\rtf1{\object\objemb\objalign1\objalign2{\*\objdata 00}}}"#,
        r#"{\rtf1{\object\objemb{\objclass X}{\*\objdata 00}}}"#,
        r#"{\rtf1{\object\objemb{\*\oleclsid1 X}{\*\objdata 00}}}"#,
        r#"{\rtf1{\object\objemb{\*\oleclsid X}{\*\oleclsid Y}{\*\objdata 00}}}"#,
        r#"{\rtf1{\object\objemb{\*\objdata 00}{\*\oleclsid X}}}"#,
        r#"{\rtf1{\object\objemb{\*\objdata 00}{\*\result x}}}"#,
        r#"{\rtf1{\object\objemb{\*\objclass X{\object}}{\*\objdata 00}}}"#,
        r#"{\rtf1\intbl{\object\objemb{\*\objdata 00}}}"#,
    ] {
        assert!(
            RtfDocument::parse(source).is_err(),
            "accepted malformed {source}"
        );
    }
}
