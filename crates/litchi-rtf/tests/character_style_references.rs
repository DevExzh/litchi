use litchi_rtf::{Formatting, RtfDocument, RtfWriter, StyleType};

fn formatting_for<'a>(document: &'a RtfDocument<'a>, text: &str) -> Formatting {
    document
        .blocks()
        .iter()
        .find(|block| block.text == text)
        .unwrap_or_else(|| panic!("missing block {text:?}"))
        .formatting
}

#[test]
fn retains_scoped_references_without_materializing_style_properties() {
    let source = concat!(
        r#"{\rtf1{\stylesheet"#,
        r#"{\*\cs5\additive\b Emphasis;}"#,
        r#"{\s1\cs5\i Referencing Paragraph;}"#,
        r#"}{\cs5\b A{\cs6\i B}C}{\plain D}}"#,
    );
    let document = RtfDocument::parse(source).unwrap();

    let paragraph_style = document
        .stylesheet()
        .get_typed(StyleType::Paragraph, 1)
        .unwrap();
    assert_eq!(paragraph_style.formatting.character_style, Some(5));
    assert!(paragraph_style.formatting.italic);

    let a = formatting_for(&document, "A");
    assert_eq!(a.character_style, Some(5));
    assert!(a.bold);
    let b = formatting_for(&document, "B");
    assert_eq!(b.character_style, Some(6));
    assert!(b.bold);
    assert!(b.italic);
    assert_eq!(formatting_for(&document, "C"), a);
    assert_eq!(formatting_for(&document, "D").character_style, None);

    let character_style = document
        .stylesheet()
        .get_typed(StyleType::Character, 5)
        .unwrap();
    assert!(character_style.formatting.bold);
    assert_eq!(character_style.formatting.character_style, None);
}

#[test]
fn defaults_mutation_and_canonical_writer_preserve_zero_and_omission() {
    let document =
        RtfDocument::parse(r#"{\rtf1{\*\defchp\cs0\i}{\cs65535\b Maximum}{\plain Reset}}"#)
            .unwrap();
    assert_eq!(
        document
            .default_formatting()
            .character()
            .unwrap()
            .formatting
            .character_style,
        Some(0)
    );
    assert_eq!(
        formatting_for(&document, "Maximum").character_style,
        Some(65_535)
    );
    assert_eq!(formatting_for(&document, "Reset").character_style, None);

    let mut formatting = Formatting::default();
    formatting.set_character_style(Some(0));
    assert_eq!(formatting.character_style, Some(0));
    formatting.set_character_style(None);
    assert_eq!(formatting, Formatting::default());

    let mut first = Vec::new();
    RtfWriter::new(&mut first)
        .write_document(&document)
        .unwrap();
    let serialized = String::from_utf8(first.clone()).unwrap();
    assert!(serialized.contains(r#"{\*\defchp\cs0"#));
    assert!(serialized.contains(r#"\cs65535\fs24\b"#));
    let reparsed = RtfDocument::parse_bytes(&first).unwrap();
    assert_eq!(
        reparsed.default_formatting().character(),
        document.default_formatting().character()
    );
    assert_eq!(
        formatting_for(&reparsed, "Maximum"),
        formatting_for(&document, "Maximum")
    );
    let mut second = Vec::new();
    RtfWriter::new(&mut second)
        .write_document(&reparsed)
        .unwrap();
    assert_eq!(first, second);
}

#[test]
fn rejects_missing_out_of_range_duplicate_and_misplaced_references() {
    for source in [
        r#"{\rtf1\cs X}"#,
        r#"{\rtf1\cs-1 X}"#,
        r#"{\rtf1\cs65536 X}"#,
        r#"{\rtf1{\stylesheet{\*\cs Missing;}}}"#,
        r#"{\rtf1{\stylesheet{\*\cs-1 Negative;}}}"#,
        r#"{\rtf1{\stylesheet{\*\cs65536 Overflow;}}}"#,
        r#"{\rtf1{\stylesheet{\b\cs1 Late;}}}"#,
        r#"{\rtf1{\*\defchp\cs1\cs2}X}"#,
        r#"{\rtf1{\*\defchp\cs}X}"#,
    ] {
        assert!(RtfDocument::parse(source).is_err(), "accepted {source}");
    }

    let inert =
        RtfDocument::parse(r#"{\rtf1{\field{\*\fldinst TEST \cs65536}{\fldrslt Result}}Body}"#)
            .unwrap();
    assert_eq!(inert.fields().len(), 1);
    assert!(
        inert
            .blocks()
            .iter()
            .all(|block| block.formatting.character_style.is_none())
    );
}

#[test]
fn canonical_round_trip_is_stable_for_body_style_and_default_owners() {
    let source = concat!(
        r#"{\rtf1{\*\defchp\cs3\b}{\stylesheet"#,
        r#"{\*\cs3\additive\i Character;}"#,
        r#"{\s4\cs3\b Paragraph;}"#,
        r#"}{\cs3\i Styled}}"#,
    );
    let document = RtfDocument::parse(source).unwrap();
    let mut first = Vec::new();
    RtfWriter::new(&mut first)
        .write_document(&document)
        .unwrap();
    let reparsed = RtfDocument::parse_bytes(&first).unwrap();
    assert_eq!(reparsed.stylesheet(), document.stylesheet());
    assert_eq!(
        reparsed.default_formatting().character(),
        document.default_formatting().character()
    );
    assert_eq!(
        formatting_for(&reparsed, "Styled"),
        formatting_for(&document, "Styled")
    );
    let mut second = Vec::new();
    RtfWriter::new(&mut second)
        .write_document(&reparsed)
        .unwrap();
    assert_eq!(first, second);
}

#[test]
fn parses_libreoffice_character_style_producer() {
    let source = include_bytes!(
        "../../../test-data/libreoffice-core/sw/qa/extras/rtfimport/data/tdf163003.rtf"
    );
    let document = RtfDocument::parse_bytes(source).unwrap();
    assert!(
        document
            .stylesheet()
            .get_typed(StyleType::Character, 37)
            .is_some()
    );
}
