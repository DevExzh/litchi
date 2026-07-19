use litchi_rtf::{CharacterGrid, CharacterType, Formatting, RtfDocument, RtfWriter};

fn formatting_for<'a>(document: &'a RtfDocument<'a>, text: &str) -> Formatting {
    document
        .blocks()
        .iter()
        .find(|block| block.text == text)
        .unwrap_or_else(|| panic!("missing block {text:?}"))
        .formatting
}

#[test]
fn character_type_complex_script_and_grid_scope_and_round_trip() {
    let source = concat!(
        r#"{\rtf1\fcs1\loch\cgrid Root"#,
        r#"{\hich\fcs0\cgrid-12 Nested}"#,
        r#"Tail{\dbch\cgrid0 Double}}"#,
    );
    let document = RtfDocument::parse(source).unwrap();

    let root = formatting_for(&document, "Root");
    assert_eq!(root.character_type, Some(CharacterType::LowAnsi));
    assert_eq!(root.complex_script, Some(true));
    assert_eq!(root.character_grid, Some(CharacterGrid::Parameterless));

    let nested = formatting_for(&document, "Nested");
    assert_eq!(nested.character_type, Some(CharacterType::HighAnsi));
    assert_eq!(nested.complex_script, Some(false));
    assert_eq!(nested.character_grid, Some(CharacterGrid::Value(-12)));
    assert_eq!(formatting_for(&document, "Tail"), root);

    let double = formatting_for(&document, "Double");
    assert_eq!(double.character_type, Some(CharacterType::DoubleByte));
    assert_eq!(double.character_grid, Some(CharacterGrid::Value(0)));

    let mut first = Vec::new();
    RtfWriter::new(&mut first)
        .write_document(&document)
        .unwrap();
    let serialized = String::from_utf8(first.clone()).unwrap();
    assert!(serialized.contains(r#"\fcs1\loch\cgrid\fs24"#));
    assert!(serialized.contains(r#"\fcs0\hich\cgrid-12\fs24"#));
    assert!(serialized.contains(r#"\fcs1\dbch\cgrid0\fs24"#));

    let reparsed = RtfDocument::parse_bytes(&first).unwrap();
    for text in ["Root", "Nested", "Tail", "Double"] {
        assert_eq!(
            formatting_for(&reparsed, text),
            formatting_for(&document, text)
        );
    }
    let mut second = Vec::new();
    RtfWriter::new(&mut second)
        .write_document(&reparsed)
        .unwrap();
    assert_eq!(first, second);
}

#[test]
fn mutation_api_preserves_explicit_false_zero_and_omission() {
    let mut formatting = Formatting::default();
    formatting.set_character_type(Some(CharacterType::DoubleByte));
    formatting.set_complex_script(Some(false));
    formatting.set_character_grid(Some(CharacterGrid::Value(0)));
    assert_eq!(formatting.character_type, Some(CharacterType::DoubleByte));
    assert_eq!(formatting.complex_script, Some(false));
    assert_eq!(formatting.character_grid, Some(CharacterGrid::Value(0)));

    formatting.set_character_type(None);
    formatting.set_complex_script(None);
    formatting.set_character_grid(None);
    assert_eq!(formatting, Formatting::default());
}

#[test]
fn rejects_invalid_parameters_and_duplicate_document_defaults() {
    for source in [
        r#"{\rtf1\fcs X}"#,
        r#"{\rtf1\fcs-1 X}"#,
        r#"{\rtf1\fcs2 X}"#,
        r#"{\rtf1\loch0 X}"#,
        r#"{\rtf1\hich1 X}"#,
        r#"{\rtf1\dbch-1 X}"#,
        r#"{\rtf1\cgrid32768 X}"#,
        r#"{\rtf1\cgrid-32769 X}"#,
        r#"{\rtf1{\*\defchp\fcs0\fcs1}X}"#,
        r#"{\rtf1{\*\defchp\cgrid\cgrid0}X}"#,
    ] {
        assert!(RtfDocument::parse(source).is_err(), "accepted {source}");
    }
}

#[test]
fn keeps_selector_like_controls_in_field_instructions_inert() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1{\field{\*\fldinst TEST \fcs2\loch0\cgrid32768}"#,
        r#"{\fldrslt Result}}}"#,
    ))
    .unwrap();
    assert_eq!(document.fields().len(), 1);
    assert!(document.blocks().iter().all(|block| {
        block.formatting.character_type.is_none()
            && block.formatting.complex_script.is_none()
            && block.formatting.character_grid.is_none()
    }));
}
