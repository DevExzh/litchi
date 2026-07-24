use litchi_rtf::{LegacyNumberingAlignment, LegacyNumberingFormat, RtfDocument, RtfWriter};

const SYNTHETIC: &str = r#"{\rtf1\ansi\ansicpg1250
{\*\pnseclvl1\pnucrm\pnqc\pnstart3\pnindent720\pnsp144\pnhang\pnprev\pnf2
{\pntxtb \'8a(}{\pntxta )\u20320?}}
{\*\pnseclvl2\pndec\pnstart1\pnindent360{\pntxta .}}
Body}"#;

#[test]
fn parses_decodes_and_round_trips_legacy_section_numbering() {
    let doc = RtfDocument::parse(SYNTHETIC).unwrap();
    let numbering = doc.legacy_section_numbering();
    assert_eq!(numbering.levels().len(), 2);
    let first = numbering.get(1).unwrap();
    assert_eq!(first.format, LegacyNumberingFormat::UpperRoman);
    assert_eq!(first.alignment, Some(LegacyNumberingAlignment::Center));
    assert_eq!(first.start_at, Some(3));
    assert_eq!(first.indent, Some(720));
    assert_eq!(first.space, Some(144));
    assert!(first.hanging);
    assert!(first.previous);
    assert_eq!(first.font_ref, Some(2));
    assert_eq!(first.text_before, "Š(");
    assert_eq!(first.text_after, ")你");
    assert_eq!(doc.text().trim(), "Body");

    let mut first_bytes = Vec::new();
    RtfWriter::new(&mut first_bytes)
        .write_document(&doc)
        .unwrap();
    let reparsed = RtfDocument::parse_bytes(&first_bytes).unwrap();
    assert_eq!(numbering, reparsed.legacy_section_numbering());
    let mut second_bytes = Vec::new();
    RtfWriter::new(&mut second_bytes)
        .write_document(&reparsed)
        .unwrap();
    assert_eq!(first_bytes, second_bytes);
}

#[test]
fn rejects_malformed_legacy_section_numbering() {
    let malformed = [
        r#"{\rtf1{\pnseclvl1\pndec}}"#,
        r#"{\rtf1{\*\pnseclvl0\pndec}}"#,
        r#"{\rtf1{\*\pnseclvl1\pndec}{\*\pnseclvl1\pndec}}"#,
        r#"{\rtf1{\*\pnseclvl2\pndec}{\*\pnseclvl1\pndec}}"#,
        r#"{\rtf1{\*\pnseclvl1\pnstart1}}"#,
        r#"{\rtf1{\*\pnseclvl1\pndec{\pntxta X}{\pntxta Y}}}"#,
        r#"{\rtf1{\*\pnseclvl1\pndec{\pntxta {\field X}}}}"#,
        r#"{\rtf1 Body{\*\pnseclvl1\pndec}}"#,
    ];
    for source in malformed {
        assert!(
            RtfDocument::parse(source).is_err(),
            "accepted malformed RTF: {source}"
        );
    }
}

#[test]
fn parses_real_libreoffice_pnseclvl_fixture() {
    let doc = RtfDocument::parse_bytes(include_bytes!(
        "../../../test-data/libreoffice-core/sw/qa/uibase/uiview/data/tdf152839_formtext.rtf"
    ))
    .unwrap();
    let numbering = doc.legacy_section_numbering();
    assert_eq!(numbering.levels().len(), 9);
    assert_eq!(
        numbering.get(1).unwrap().format,
        LegacyNumberingFormat::UpperRoman
    );
    assert_eq!(
        numbering.get(3).unwrap().format,
        LegacyNumberingFormat::Decimal
    );
    assert_eq!(numbering.get(5).unwrap().text_before, "(");
    assert_eq!(numbering.get(5).unwrap().text_after, ")");
}
