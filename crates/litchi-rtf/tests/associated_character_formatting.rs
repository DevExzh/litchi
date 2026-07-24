use litchi_rtf::{
    AssociatedCharacterBaseline, AssociatedCharacterFormatting, AssociatedUnderlineStyle,
    LanguageId, RtfDocument, RtfWriter,
};

fn associated_for<'a>(
    document: &'a RtfDocument<'a>,
    text: &str,
) -> litchi_rtf::AssociatedCharacterFormatting {
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
        r#"{\ab\acaps0\acf4\adn3\aexpnd-20\af3\afs22\ai0\alang1031"#,
        r#"\aoutl\ascaps0\ashad\astrike0\auldb A{\ai\aup8\auld B}C}"#,
        r#"{\plain D}}"#,
    );
    let document = RtfDocument::parse(source).unwrap();

    let a = associated_for(&document, "A");
    assert_eq!(a.font_ref, Some(3));
    assert_eq!(a.font_size.unwrap().get(), 22);
    assert_eq!(a.language, Some(LanguageId::new(1031).unwrap()));
    assert_eq!(a.bold, Some(true));
    assert_eq!(a.all_caps, Some(false));
    assert_eq!(a.color_ref, Some(4));
    assert_eq!(
        a.baseline,
        Some(AssociatedCharacterBaseline::LoweredHalfPoints(3))
    );
    assert_eq!(a.expansion_quarter_points, Some(-20));
    assert_eq!(a.italic, Some(false));
    assert_eq!(a.outline, Some(true));
    assert_eq!(a.small_caps, Some(false));
    assert_eq!(a.shadow, Some(true));
    assert_eq!(a.strike, Some(false));
    assert_eq!(a.underline, Some(AssociatedUnderlineStyle::Double));

    let b = associated_for(&document, "B");
    assert_eq!(b.font_ref, Some(3));
    assert_eq!(b.bold, Some(true));
    assert_eq!(b.italic, Some(true));
    assert_eq!(
        b.baseline,
        Some(AssociatedCharacterBaseline::RaisedHalfPoints(8))
    );
    assert_eq!(b.underline, Some(AssociatedUnderlineStyle::Dotted));
    assert_eq!(associated_for(&document, "C"), a);

    let reset = associated_for(&document, "D");
    assert_eq!(reset.font_ref, None);
    assert_eq!(reset.font_size, None);
    assert_eq!(reset.bold, None);
    assert_eq!(reset.italic, None);
    assert_eq!(reset.all_caps, None);
    assert_eq!(reset.color_ref, None);
    assert_eq!(reset.baseline, None);
    assert_eq!(reset.expansion_quarter_points, None);
    assert_eq!(reset.outline, None);
    assert_eq!(reset.small_caps, None);
    assert_eq!(reset.shadow, None);
    assert_eq!(reset.strike, None);
    assert_eq!(reset.underline, None);
    assert_eq!(reset.language, Some(LanguageId::new(1025).unwrap()));

    let mut first = Vec::new();
    RtfWriter::new(&mut first)
        .write_document(&document)
        .unwrap();
    let serialized = String::from_utf8(first.clone()).unwrap();
    assert!(serialized.contains(concat!(
        r#"\ab1\acaps0\acf4\adn3\aexpnd-20\af3\afs22\ai0\alang1031"#,
        r#"\aoutl1\ascaps0\ashad1\astrike0\auldb"#,
    )));
    let reparsed = RtfDocument::parse_bytes(&first).unwrap();
    assert_eq!(associated_for(&reparsed, "A"), a);
    assert_eq!(associated_for(&reparsed, "B"), b);
    assert_eq!(associated_for(&reparsed, "D"), reset);
    let mut second = Vec::new();
    RtfWriter::new(&mut second)
        .write_document(&reparsed)
        .unwrap();
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
        r#"{\rtf1\ab2 X}"#,
        r#"{\rtf1\ai-1 X}"#,
        r#"{\rtf1\acaps2 X}"#,
        r#"{\rtf1\acf X}"#,
        r#"{\rtf1\acf-1 X}"#,
        r#"{\rtf1\acf65536 X}"#,
        r#"{\rtf1\adn X}"#,
        r#"{\rtf1\adn-1 X}"#,
        r#"{\rtf1\adn31681 X}"#,
        r#"{\rtf1\aexpnd X}"#,
        r#"{\rtf1\aexpnd-31681 X}"#,
        r#"{\rtf1\aexpnd31681 X}"#,
        r#"{\rtf1\aoutl2 X}"#,
        r#"{\rtf1\ascaps2 X}"#,
        r#"{\rtf1\ashad2 X}"#,
        r#"{\rtf1\astrike2 X}"#,
        r#"{\rtf1\aul2 X}"#,
        r#"{\rtf1\auld0 X}"#,
        r#"{\rtf1\auldb1 X}"#,
        r#"{\rtf1\aulnone0 X}"#,
        r#"{\rtf1\aulw1 X}"#,
        r#"{\rtf1\aup X}"#,
        r#"{\rtf1\aup-1 X}"#,
        r#"{\rtf1\aup31681 X}"#,
        r#"{\rtf1{\*\defchp\adn1\aup2}X}"#,
        r#"{\rtf1{\*\defchp\aul\aulnone}X}"#,
        r#"{\rtf1{\stylesheet{\s0\af-1 Bad;}}Body}"#,
    ] {
        assert!(RtfDocument::parse(source).is_err(), "accepted {source}");
    }

    let inert =
        RtfDocument::parse(r#"{\rtf1{\field{\*\fldinst TEST \af-1\afs0\alang-1}{\fldrslt X}}}"#)
            .unwrap();
    assert_eq!(inert.fields().len(), 1);
    assert!(
        inert
            .blocks()
            .iter()
            .all(|block| block.formatting.associated.font_ref.is_none())
    );
}

#[test]
fn associated_mutation_validates_and_clears_exact_state() {
    let mut associated = AssociatedCharacterFormatting::default();
    assert!(
        associated
            .set_baseline(Some(AssociatedCharacterBaseline::RaisedHalfPoints(31_681)))
            .is_err()
    );
    assert!(
        associated
            .set_expansion_quarter_points(Some(31_681))
            .is_err()
    );

    associated
        .set_baseline(Some(AssociatedCharacterBaseline::LoweredHalfPoints(31_680)))
        .unwrap();
    associated
        .set_expansion_quarter_points(Some(-31_680))
        .unwrap();
    associated.underline = Some(AssociatedUnderlineStyle::None);
    associated.validate().unwrap();
    associated.clear();
    assert_eq!(associated, AssociatedCharacterFormatting::default());
}

#[test]
fn parses_specification_association_examples() {
    let document = RtfDocument::parse(
        r#"{\rtf1{\ltrch\af2\ab\aul\rtlch First}{\rtlch\af5\ab\ai\ltrch Second}}"#,
    )
    .unwrap();
    let first = associated_for(&document, "First");
    assert_eq!(first.font_ref, Some(2));
    assert_eq!(first.bold, Some(true));
    assert_eq!(first.underline, Some(AssociatedUnderlineStyle::Single));
    let second = associated_for(&document, "Second");
    assert_eq!(second.font_ref, Some(5));
    assert_eq!(second.bold, Some(true));
    assert_eq!(second.italic, Some(true));
}

#[test]
fn parses_libreoffice_associated_character_fixture() {
    let source = include_bytes!(
        "../../../test-data/libreoffice-core/sw/qa/extras/rtfimport/data/sbkeven.rtf"
    );
    let document = RtfDocument::parse_bytes(source).unwrap();
    assert!(document.blocks().iter().any(|block| {
        let associated = block.formatting.associated;
        associated.font_ref == Some(31507)
            && associated.font_size.is_some_and(|size| size.get() == 22)
            && associated.language == LanguageId::new(1025).ok()
    }));
}
