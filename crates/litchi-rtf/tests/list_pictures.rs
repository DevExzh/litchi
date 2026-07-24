use litchi_rtf::{ImageType, RtfDocument, RtfWriter};

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn parses_real_libreoffice_picture_bullet_without_cloning_payload() {
    let source = include_bytes!(
        "../../../test-data/libreoffice-core/sw/qa/extras/rtfexport/data/i120928.rtf"
    );
    let document = RtfDocument::parse_bytes(source).unwrap();
    let indices = document.list_table().picture_bullet_picture_indices();
    assert_eq!(
        indices.len(),
        document.list_table().picture_bullet_count as usize
    );
    let index = indices[0].unwrap();
    let bullet = document.list_picture_bullets().next().flatten().unwrap();
    assert!(std::ptr::eq(bullet, &document.pictures()[index]));
    assert_eq!(bullet.image_type, ImageType::Png);

    let output = write(&document);
    let serialized = String::from_utf8_lossy(&output);
    assert!(serialized.contains("{\\*\\listpicture{\\*\\shppict{\\pict"));
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(
        reparsed.list_table().picture_bullet_count,
        document.list_table().picture_bullet_count
    );
    assert_eq!(
        reparsed
            .list_picture_bullets()
            .next()
            .flatten()
            .unwrap()
            .data(),
        bullet.data()
    );
}

#[test]
fn canonical_roundtrip_preserves_multiple_records_and_empty_legacy_entry() {
    let source = concat!(
        r#"{\rtf1{\*\listtable{\*\listpicture"#,
        r#"{\*\shppict{\pict\pngblip 89504e470d0a1a0a}}"#,
        r#"{\*\shppict{\pict\jpegblip ffd8ffe0}}}}}"#,
    );
    let document = RtfDocument::parse(source).unwrap();
    assert_eq!(document.list_table().picture_bullet_count, 2);
    assert_eq!(document.list_picture_bullets().count(), 2);
    let first = write(&document);
    let reparsed = RtfDocument::parse_bytes(&first).unwrap();
    assert_eq!(
        reparsed.list_table().picture_bullet_picture_indices().len(),
        2
    );
    assert_eq!(reparsed.pictures().len(), 2);
    assert_eq!(write(&reparsed), first);

    let empty = RtfDocument::parse(r#"{\rtf1{\*\listtable{\*\listpicture}}}"#).unwrap();
    assert_eq!(empty.list_table().picture_bullet_count, 1);
    assert!(empty.list_picture_bullets().next().unwrap().is_none());
    let rewritten = write(&empty);
    let reparsed = RtfDocument::parse_bytes(&rewritten).unwrap();
    assert_eq!(reparsed.list_table().picture_bullet_count, 1);
}

#[test]
fn typed_indices_validate_and_clear_only_references() {
    let mut document =
        RtfDocument::parse(r#"{\rtf1{\pict\pngblip 89504e470d0a1a0a}Body}"#).unwrap();
    assert_eq!(document.pictures().len(), 1);
    document
        .set_list_picture_bullet_indices(vec![Some(0), None])
        .unwrap();
    assert_eq!(document.list_picture_bullets().count(), 2);
    assert!(
        document
            .set_list_picture_bullet_indices(vec![Some(1)])
            .is_err()
    );
    let output = write(&document);
    assert_eq!(
        RtfDocument::parse_bytes(&output).unwrap().pictures().len(),
        1
    );

    document.clear_list_picture_bullets().unwrap();
    assert_eq!(document.pictures().len(), 1);
    assert_eq!(document.list_table().picture_bullet_count, 0);
}

#[test]
fn rejects_hostile_list_picture_grammar() {
    for source in [
        r#"{\rtf1{\*\listtable{\listpicture}}}"#,
        r#"{\rtf1{\*\listtable{\*\listpicture1}}}"#,
        r#"{\rtf1{\*\listtable{\*\listpicture text}}}"#,
        r#"{\rtf1{\*\listtable{\*\listpicture{\shppict{\pict\pngblip 00}}}}}"#,
        r#"{\rtf1{\*\listtable{\*\listpicture{\*\shppict1{\pict\pngblip 00}}}}}"#,
        r#"{\rtf1{\*\listtable{\*\listpicture{\*\shppict}}}}"#,
        r#"{\rtf1{\*\listtable{\*\listpicture{\*\shppict{\pict}}}}}"#,
        r#"{\rtf1{\*\listtable{\*\listpicture{\*\shppict{\pict\pngblip 00}{\pict\pngblip 00}}}}}"#,
        r#"{\rtf1{\*\listtable{\*\listpicture}{\*\listpicture}}}"#,
        r#"{\rtf1{\*\listpicture{\*\shppict{\pict\pngblip 00}}}}"#,
    ] {
        assert!(
            RtfDocument::parse(source).is_err(),
            "accepted malformed {source}"
        );
    }
}
