use litchi_rtf::{
    EndnoteRestart, FootnoteRestart, NoteNumberingStyle, NoteOptions, NotePlacement,
    PresentNoteKinds, RtfDocument, RtfWriter,
};

#[test]
fn parses_last_wins_note_options_and_round_trips_stably() {
    let source = concat!(
        r#"{\rtf1\ansi\fet0\fet2\endnotes\ftnbj\ftnstart4\ftnstart7"#,
        r#"\ftnrstpg\ftnrestart\ftnnalc\ftnnchi"#,
        r#"\aendnotes\aenddoc\aftnstart3\aftnrestart\aftnnruc Body}"#,
    );
    let document = RtfDocument::parse(source).unwrap();
    assert_eq!(document.text(), "Body");
    assert_eq!(
        *document.note_options(),
        NoteOptions {
            present_kinds: Some(PresentNoteKinds::FootnotesAndEndnotes),
            footnote_placement: Some(NotePlacement::BottomOfPage),
            endnote_placement: Some(NotePlacement::EndOfDocument),
            footnote_start: Some(7),
            endnote_start: Some(3),
            footnote_restart: Some(FootnoteRestart::EachSection),
            endnote_restart: Some(EndnoteRestart::EachSection),
            footnote_numbering: Some(NoteNumberingStyle::Chicago),
            endnote_numbering: Some(NoteNumberingStyle::UppercaseRoman),
        }
    );

    let mut first = Vec::new();
    RtfWriter::new(&mut first)
        .write_document(&document)
        .unwrap();
    let reparsed = RtfDocument::parse_bytes(&first).unwrap();
    assert_eq!(reparsed.text(), document.text());
    assert_eq!(reparsed.note_options(), document.note_options());
    let mut second = Vec::new();
    RtfWriter::new(&mut second)
        .write_document(&reparsed)
        .unwrap();
    assert_eq!(first, second);
}

#[test]
fn supports_all_42_note_numbering_control_spellings() {
    let cases = [
        ("ftnnar", "aftnnar", NoteNumberingStyle::Arabic),
        ("ftnnalc", "aftnnalc", NoteNumberingStyle::LowercaseLetter),
        ("ftnnauc", "aftnnauc", NoteNumberingStyle::UppercaseLetter),
        ("ftnnrlc", "aftnnrlc", NoteNumberingStyle::LowercaseRoman),
        ("ftnnruc", "aftnnruc", NoteNumberingStyle::UppercaseRoman),
        ("ftnnchi", "aftnnchi", NoteNumberingStyle::Chicago),
        (
            "ftnnchosung",
            "aftnnchosung",
            NoteNumberingStyle::KoreanChosung,
        ),
        ("ftnncnum", "aftnncnum", NoteNumberingStyle::Circle),
        (
            "ftnndbnum",
            "aftnndbnum",
            NoteNumberingStyle::KanjiDigitless,
        ),
        (
            "ftnndbnumd",
            "aftnndbnumd",
            NoteNumberingStyle::KanjiWithDigit,
        ),
        ("ftnndbnumt", "aftnndbnumt", NoteNumberingStyle::KanjiThree),
        ("ftnndbnumk", "aftnndbnumk", NoteNumberingStyle::KanjiFour),
        ("ftnndbar", "aftnndbar", NoteNumberingStyle::DoubleByte),
        (
            "ftnnganada",
            "aftnnganada",
            NoteNumberingStyle::KoreanGanada,
        ),
        ("ftnngbnum", "aftnngbnum", NoteNumberingStyle::ChineseOne),
        ("ftnngbnumd", "aftnngbnumd", NoteNumberingStyle::ChineseTwo),
        (
            "ftnngbnuml",
            "aftnngbnuml",
            NoteNumberingStyle::ChineseThree,
        ),
        ("ftnngbnumk", "aftnngbnumk", NoteNumberingStyle::ChineseFour),
        ("ftnnzodiac", "aftnnzodiac", NoteNumberingStyle::ZodiacOne),
        ("ftnnzodiacd", "aftnnzodiacd", NoteNumberingStyle::ZodiacTwo),
        (
            "ftnnzodiacl",
            "aftnnzodiacl",
            NoteNumberingStyle::ZodiacThree,
        ),
    ];

    for (footnote, endnote, style) in cases {
        let source = format!(r#"{{\rtf1\{footnote}\{endnote} X}}"#);
        let document = RtfDocument::parse(&source).unwrap();
        assert_eq!(document.note_options().footnote_numbering, Some(style));
        assert_eq!(document.note_options().endnote_numbering, Some(style));

        let mut output = Vec::new();
        RtfWriter::new(&mut output)
            .write_document(&document)
            .unwrap();
        let output = String::from_utf8(output).unwrap();
        assert!(output.contains(&format!(r#"\{footnote}"#)));
        assert!(output.contains(&format!(r#"\{endnote}"#)));
    }
}

#[test]
fn rejects_invalid_values_and_non_root_or_late_note_options() {
    let malformed = [
        r#"{\rtf1\fet-1 Body}"#,
        r#"{\rtf1\fet3 Body}"#,
        r#"{\rtf1\fet4 Body}"#,
        r#"{\rtf1\ftnstart0 Body}"#,
        r#"{\rtf1\ftnstart-1 Body}"#,
        r#"{\rtf1\aftnstart0 Body}"#,
        r#"{\rtf1\aftnstart-1 Body}"#,
        r#"{\rtf1\ftnstart2147483648 Body}"#,
        r#"{\rtf1 Body\ftnbj}"#,
        r#"{\rtf1\'41\ftnbj}"#,
        r#"{\rtf1\u65?\ftnbj}"#,
        r#"{\rtf1{\ftnbj}Body}"#,
        r#"{\rtf1{\*\ftnbj}Body}"#,
        r#"{\rtf1{\header\ftnbj X}Body}"#,
        r#"{\rtf1{\footer\aftnbj X}Body}"#,
        r#"{\rtf1{\annotation\ftnstart2 X}Body}"#,
        r#"{\rtf1{\footnote\ftnrstpg X}Body}"#,
        r#"{\rtf1{\field\ftnnar X}Body}"#,
        r#"{\rtf1{\object\aftnnar X}Body}"#,
        r#"{\rtf1{\ftnbj\bin2 AB}Body}"#,
    ];
    for source in malformed {
        assert!(
            RtfDocument::parse(source).is_err(),
            "accepted malformed note options: {source}"
        );
    }

    let invalid = NoteOptions {
        footnote_start: Some(0),
        ..NoteOptions::default()
    };
    assert!(
        RtfWriter::new(Vec::new())
            .write_note_options(&invalid)
            .is_err()
    );
}

#[test]
fn omission_stays_empty_and_note_bodies_do_not_infer_options() {
    let document = RtfDocument::parse(r#"{\rtf1 A{\footnote Note}B}"#).unwrap();
    assert!(document.note_options().is_empty());
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(&document)
        .unwrap();
    assert!(!String::from_utf8(output).unwrap().contains(r#"\fet"#));
}

#[test]
fn parses_named_libreoffice_note_option_fixtures() {
    let section_end = include_bytes!(
        "../../../test-data/libreoffice-core/sw/qa/writerfilter/rtftok/data/endnote-at-section-end.rtf"
    );
    let document = RtfDocument::parse_bytes(section_end).unwrap();
    assert_eq!(
        document.note_options().present_kinds,
        Some(PresentNoteKinds::EndnotesOnly)
    );
    assert_eq!(
        document.note_options().endnote_placement,
        Some(NotePlacement::EndOfSection)
    );

    let restart = include_bytes!(
        "../../../test-data/libreoffice-core/sw/qa/extras/rtfexport/data/tdf158982.rtf"
    );
    let document = RtfDocument::parse_bytes(restart).unwrap();
    let options = document.note_options();
    assert_eq!(
        options.present_kinds,
        Some(PresentNoteKinds::FootnotesAndEndnotes)
    );
    assert_eq!(
        options.footnote_placement,
        Some(NotePlacement::BottomOfPage)
    );
    assert_eq!(options.footnote_start, Some(1));
    assert_eq!(options.footnote_restart, Some(FootnoteRestart::EachSection));
    assert_eq!(
        options.endnote_placement,
        Some(NotePlacement::EndOfDocument)
    );
    assert_eq!(options.endnote_start, Some(1));

    let arabic = include_bytes!(
        "../../../test-data/libreoffice-core/sw/qa/extras/rtfimport/data/tdf108947.rtf"
    );
    let document = RtfDocument::parse_bytes(arabic).unwrap();
    let options = document.note_options();
    assert_eq!(
        options.present_kinds,
        Some(PresentNoteKinds::FootnotesAndEndnotes)
    );
    assert_eq!(
        options.footnote_placement,
        Some(NotePlacement::BottomOfPage)
    );
    assert_eq!(options.footnote_numbering, Some(NoteNumberingStyle::Arabic));
    assert_eq!(options.endnote_numbering, Some(NoteNumberingStyle::Arabic));
}
