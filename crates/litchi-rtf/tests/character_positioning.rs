use litchi_rtf::{CharacterBaseline, CharacterExpansion, RtfDocument, RtfWriter};

fn block<'a>(document: &'a RtfDocument<'a>, needle: &str) -> &'a litchi_rtf::StyleBlock<'a> {
    document
        .blocks()
        .iter()
        .find(|block| block.text.contains(needle))
        .unwrap()
}

#[test]
fn parses_inherits_resets_and_keeps_destinations_inert() {
    let source = r#"{\rtf1\ansi\super\expnd4\charscalex80\kerning16 Outer{\dn3\expndtw20 Inner}{Tail}\nosupersub Normal{\up2 Raised}{\plain Plain}{\*\unknown\up999999\expndtw999999 ignored}Visible}"#;
    let document = RtfDocument::parse(source).unwrap();
    let outer = block(&document, "Outer");
    assert_eq!(
        outer.formatting.character_positioning.baseline,
        CharacterBaseline::Superscript
    );
    assert_eq!(
        outer.formatting.character_positioning.expansion,
        CharacterExpansion::QuarterPoints(4)
    );
    assert_eq!(
        outer
            .formatting
            .character_positioning
            .horizontal_scale_percent,
        80
    );
    assert_eq!(
        outer.formatting.character_positioning.kerning_half_points,
        16
    );
    assert_eq!(
        block(&document, "Inner")
            .formatting
            .character_positioning
            .baseline,
        CharacterBaseline::LoweredHalfPoints(3)
    );
    assert_eq!(
        block(&document, "Inner")
            .formatting
            .character_positioning
            .expansion,
        CharacterExpansion::Twips(20)
    );
    assert_eq!(
        block(&document, "Tail")
            .formatting
            .character_positioning
            .baseline,
        CharacterBaseline::Superscript
    );
    assert_eq!(
        block(&document, "Normal")
            .formatting
            .character_positioning
            .baseline,
        CharacterBaseline::Normal
    );
    assert_eq!(
        block(&document, "Raised")
            .formatting
            .character_positioning
            .baseline,
        CharacterBaseline::RaisedHalfPoints(2)
    );
    assert_eq!(
        block(&document, "Plain").formatting.character_positioning,
        Default::default()
    );
    assert_eq!(
        block(&document, "Visible")
            .formatting
            .character_positioning
            .baseline,
        CharacterBaseline::Normal
    );
}

#[test]
fn writer_is_deterministic_and_preserves_units() {
    let document =
        RtfDocument::parse(r#"{\rtf1 A{\up2\expndtw-15\charscalex75\kerning8 B}{\sub\expnd3 C}}"#)
            .unwrap();
    let mut first = Vec::new();
    RtfWriter::new(&mut first)
        .write_document(&document)
        .unwrap();
    let first_text = String::from_utf8(first).unwrap();
    assert!(first_text.contains("\\up2"));
    assert!(first_text.contains("\\expndtw-15"));
    let reparsed = RtfDocument::parse(&first_text).unwrap();
    assert_eq!(
        block(&reparsed, "B").formatting.character_positioning,
        block(&document, "B").formatting.character_positioning
    );
    let mut second = Vec::new();
    RtfWriter::new(&mut second)
        .write_document(&reparsed)
        .unwrap();
    assert_eq!(first_text, String::from_utf8(second).unwrap());
}

#[test]
fn parses_libreoffice_superscript_fixture() {
    let source = std::fs::read_to_string(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/libreoffice-core/sw/qa/extras/rtfexport/data/tdf87034.rtf"),
    )
    .unwrap();
    let document = RtfDocument::parse(&source).unwrap();
    assert!(
        document
            .runs()
            .iter()
            .any(|run| run.formatting.character_positioning.baseline
                == CharacterBaseline::Superscript)
    );
}

#[test]
fn rejects_out_of_range_parameters() {
    for source in [
        r#"{\rtf1\up-1 X}"#,
        r#"{\rtf1\up31681 X}"#,
        r#"{\rtf1\dn-1 X}"#,
        r#"{\rtf1\expnd31681 X}"#,
        r#"{\rtf1\expndtw-31681 X}"#,
        r#"{\rtf1\charscalex0 X}"#,
        r#"{\rtf1\charscalex601 X}"#,
        r#"{\rtf1\kerning-1 X}"#,
        r#"{\rtf1\kerning32768 X}"#,
    ] {
        assert!(RtfDocument::parse(source).is_err(), "accepted {source}");
    }
}
