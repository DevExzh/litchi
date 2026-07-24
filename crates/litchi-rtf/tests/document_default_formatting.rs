use litchi_rtf::{DefaultFormattingDestination, ParagraphWrapping, RtfDocument, RtfWriter};

const PRODUCER: &str = r#"{\rtf1\ansi\adeff7\deff2\stshfdbch31505\stshfloch31506\stshfhich31507\stshfbi31508
{\fonttbl{\f2 Arial;}}
{\*\defpap\qj\li120\ri240\sa200\sl276\slmult1\nowidctlpar\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0\itap0}
{\*\defchp\f2\fs22\lang1033\langfe1041\langnp1033\langfenp1041\kerning2\loch\af31506\hich\af31507\dbch\af31505}
Body}"#;

#[test]
fn parses_preserves_order_and_round_trips_defaults_inertly() {
    let document = RtfDocument::parse(PRODUCER).unwrap();
    let defaults = document.default_formatting();
    assert_eq!(defaults.fonts.primary, Some(2));
    assert_eq!(defaults.fonts.associated, Some(7));
    assert_eq!(defaults.fonts.stylesheet_double_byte, Some(31505));
    assert_eq!(
        defaults.destination_order(),
        &[
            DefaultFormattingDestination::Paragraph,
            DefaultFormattingDestination::Character
        ]
    );
    let character = defaults.character().unwrap();
    assert_eq!(character.formatting.font_ref, 2);
    assert_eq!(character.formatting.font_size.get(), 22);
    assert_eq!(character.low_ansi_font, Some(31506));
    assert_eq!(character.high_ansi_font, Some(31507));
    assert_eq!(character.double_byte_font, Some(31505));
    let paragraph = defaults.paragraph().unwrap();
    assert_eq!(paragraph.paragraph.indentation.left, 120);
    assert_eq!(paragraph.table_nesting_level, Some(0));
    assert_eq!(
        paragraph.paragraph.line_breaking.wrapping,
        ParagraphWrapping::Default
    );
    let body = document
        .blocks()
        .iter()
        .find(|block| block.text.contains("Body"))
        .unwrap();
    assert_eq!(body.formatting.font_size.get(), 24);
    assert_eq!(body.paragraph.indentation.left, 0);
    let mut first = Vec::new();
    RtfWriter::new(&mut first)
        .write_document(&document)
        .unwrap();
    let text = String::from_utf8_lossy(&first);
    assert!(text.find("defpap").unwrap() < text.find("defchp").unwrap());
    let reparsed = RtfDocument::parse_bytes(&first).unwrap();
    assert_eq!(defaults, reparsed.default_formatting());
    let mut second = Vec::new();
    RtfWriter::new(&mut second)
        .write_document(&reparsed)
        .unwrap();
    assert_eq!(first, second);
}

#[test]
fn parses_real_producer_default_destinations() {
    let first = RtfDocument::parse_bytes(include_bytes!(
        "../../../test-data/libreoffice-core/sw/qa/core/text/data/A011-charheight.rtf"
    ))
    .unwrap();
    assert!(first.default_formatting().character().is_some());
    assert!(first.default_formatting().paragraph().is_some());
    let cjk = RtfDocument::parse_bytes(include_bytes!(
        "../../../test-data/libreoffice-core/sw/qa/extras/rtfexport/data/cjklist24.rtf"
    ))
    .unwrap();
    assert_eq!(
        cjk.default_formatting()
            .character()
            .unwrap()
            .double_byte_font,
        Some(31505)
    );
}

#[test]
fn rejects_malformed_default_formatting() {
    let malformed = [
        r#"{\rtf1\deff X}"#,
        r#"{\rtf1\adeff-1 X}"#,
        r#"{\rtf1\deff0\deff1 X}"#,
        r#"{\rtf1{\defchp\fs22}X}"#,
        r#"{\rtf1{\*\defchp1\fs22}X}"#,
        r#"{\rtf1{\*\defchp\fs0}X}"#,
        r#"{\rtf1{\*\defchp\fs22\fs24}X}"#,
        r#"{\rtf1{\*\defchp\loch}X}"#,
        r#"{\rtf1{\*\defchp\loch\af}X}"#,
        r#"{\rtf1{\*\defchp\ql}X}"#,
        r#"{\rtf1{\*\defchp{\field X}}X}"#,
        r#"{\rtf1{\*\defpap\b}X}"#,
        r#"{\rtf1{\*\defpap\li100\li200}X}"#,
        r#"{\rtf1{\*\defpap\tqc}X}"#,
        r#"{\rtf1{\*\defpap\itap33}X}"#,
        r#"{\rtf1{\*\defpap}{\*\defpap}X}"#,
        r#"{\rtf1 X{\*\defchp\fs22}}"#,
    ];
    for source in malformed {
        assert!(
            RtfDocument::parse(source).is_err(),
            "accepted malformed RTF: {source}"
        );
    }
}
