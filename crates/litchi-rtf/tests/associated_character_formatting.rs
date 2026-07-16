use litchi_rtf::{LanguageId, RtfDocument, RtfWriter};

fn associated_for<'a>(document: &'a RtfDocument<'a>, text: &str) -> litchi_rtf::AssociatedCharacterFormatting {
    document
        .blocks()
        .iter()
        .find(|block| block.text == text)
        .unwrap_or_else(|| panic!("missing block {text:?}"))
        .formatting
        .associated
}

#[test]
fn parses_inherits_resets_and_writes_associated_character_formatting() {
    let source = concat!(
        r#"{\rtf1\adeflang1025"#,
        r#"{\af3\afs22\alang1031\ab\ai0 A{\ai B}C}"#,
        r#"{\plain D}}"#,
    );
    let document = RtfDocument::parse(source).unwrap();

    let a = associated_for(&document, "A");
    assert_eq!(a.font_ref, Some(3));
    assert_eq!(a.font_size.unwrap().get(), 22);
    assert_eq!(a.language, Some(LanguageId::new(1031).unwrap()));
    assert_eq!(a.bold, Some(true));
    assert_eq!(a.italic, Some(false));

    let b = associated_for(&document, "B");
    assert_eq!(b.font_ref, Some(3));
    assert_eq!(b.bold, Some(true));
    assert_eq!(b.italic, Some(true));
    assert_eq!(associated_for(&document, "C"), a);

    let reset = associated_for(&document, "D");
    assert_eq!(reset.font_ref, None);
    assert_eq!(reset.font_size, None);
    assert_eq!(reset.bold, None);
    assert_eq!(reset.italic, None);
    assert_eq!(reset.language, Some(LanguageId::new(1025).unwrap()));

    let mut first = Vec::new();
    RtfWriter::new(&mut first).write_document(&document).unwrap();
    let serialized = String::from_utf8(first.clone()).unwrap();
    assert!(serialized.contains(r#"\af3\afs22\alang1031\ab1\ai0"#));
    let reparsed = RtfDocument::parse_bytes(&first).unwrap();
    assert_eq!(associated_for(&reparsed, "A"), a);
    assert_eq!(associated_for(&reparsed, "B"), b);
    assert_eq!(associated_for(&reparsed, "D"), reset);
    let mut second = Vec::new();
    RtfWriter::new(&mut second).write_document(&reparsed).unwrap();
    assert_eq!(first, second);
}

#[test]
fn rejects_unbounded_or_missing_associated_values_but_keeps_field_code_inert() {
    for source in [
        r#"{\rtf1\af X}"#,
        r#"{\rtf1\af-1 X}"#,
        r#"{\rtf1\af65536 X}"#,
        r#"{\rtf1\afs X}"#,
        r#"{\rtf1\afs0 X}"#,
        r#"{\rtf1\afs-1 X}"#,
        r#"{\rtf1\afs65536 X}"#,
        r#"{\rtf1\alang X}"#,
        r#"{\rtf1\alang-1 X}"#,
        r#"{\rtf1\alang65536 X}"#,
        r#"{\rtf1{\stylesheet{\s0\af-1 Bad;}}Body}"#,
    ] {
        assert!(RtfDocument::parse(source).is_err(), "accepted {source}");
    }

    let inert = RtfDocument::parse(
        r#"{\rtf1{\field{\*\fldinst TEST \af-1\afs0\alang-1}{\fldrslt X}}}"#,
    )
    .unwrap();
    assert_eq!(inert.fields().len(), 1);
    assert!(inert
        .blocks()
        .iter()
        .all(|block| block.formatting.associated.font_ref.is_none()));
}

#[test]
fn parses_libreoffice_associated_character_fixture() {
    let source = include_bytes!(
        "../../../3rdparty/libreoffice-core/sw/qa/extras/rtfimport/data/sbkeven.rtf"
    );
    let document = RtfDocument::parse_bytes(source).unwrap();
    assert!(document.blocks().iter().any(|block| {
        let associated = block.formatting.associated;
        associated.font_ref == Some(31507)
            && associated.font_size.is_some_and(|size| size.get() == 22)
            && associated.language == LanguageId::new(1025).ok()
    }));
}
