use litchi_rtf::{
    EndnoteRestart, FootnoteRestart, NoteNumberingStyle, RtfDocument, RtfWriter,
    SectionFootnotePlacement, SectionNoteOptions,
};

#[test]
fn parses_last_wins_section_note_options_and_round_trips_stably() {
    let source = concat!(
        r#"{\rtf1\sectd\sftntj\sftnbj\sftnstart2\sftnstart7"#,
        r#"\sftnrstpg\sftnrestart\sftnnalc\sftnnchi"#,
        r#"\saftnstart3\saftnrestart\saftnnar\saftnnruc Body}"#,
    );
    let document = RtfDocument::parse(source).unwrap();
    assert_eq!(document.text(), "Body");
    assert_eq!(
        document.sections()[0].properties.note_options,
        SectionNoteOptions {
            footnote_placement: Some(SectionFootnotePlacement::BottomOfPage),
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
    assert_eq!(
        reparsed.sections()[0].properties.note_options,
        document.sections()[0].properties.note_options
    );
    let mut second = Vec::new();
    RtfWriter::new(&mut second)
        .write_document(&reparsed)
        .unwrap();
    assert_eq!(first, second);
}

#[test]
fn supports_all_42_section_numbering_control_spellings() {
    let cases = [
        ("sftnnar", "saftnnar", NoteNumberingStyle::Arabic),
        ("sftnnalc", "saftnnalc", NoteNumberingStyle::LowercaseLetter),
        ("sftnnauc", "saftnnauc", NoteNumberingStyle::UppercaseLetter),
        ("sftnnrlc", "saftnnrlc", NoteNumberingStyle::LowercaseRoman),
        ("sftnnruc", "saftnnruc", NoteNumberingStyle::UppercaseRoman),
        ("sftnnchi", "saftnnchi", NoteNumberingStyle::Chicago),
        (
            "sftnnchosung",
            "saftnnchosung",
            NoteNumberingStyle::KoreanChosung,
        ),
        ("sftnncnum", "saftnncnum", NoteNumberingStyle::Circle),
        (
            "sftnndbnum",
            "saftnndbnum",
            NoteNumberingStyle::KanjiDigitless,
        ),
        (
            "sftnndbnumd",
            "saftnndbnumd",
            NoteNumberingStyle::KanjiWithDigit,
        ),
        (
            "sftnndbnumt",
            "saftnndbnumt",
            NoteNumberingStyle::KanjiThree,
        ),
        ("sftnndbnumk", "saftnndbnumk", NoteNumberingStyle::KanjiFour),
        ("sftnndbar", "saftnndbar", NoteNumberingStyle::DoubleByte),
        (
            "sftnnganada",
            "saftnnganada",
            NoteNumberingStyle::KoreanGanada,
        ),
        ("sftnngbnum", "saftnngbnum", NoteNumberingStyle::ChineseOne),
        (
            "sftnngbnumd",
            "saftnngbnumd",
            NoteNumberingStyle::ChineseTwo,
        ),
        (
            "sftnngbnuml",
            "saftnngbnuml",
            NoteNumberingStyle::ChineseThree,
        ),
        (
            "sftnngbnumk",
            "saftnngbnumk",
            NoteNumberingStyle::ChineseFour,
        ),
        ("sftnnzodiac", "saftnnzodiac", NoteNumberingStyle::ZodiacOne),
        (
            "sftnnzodiacd",
            "saftnnzodiacd",
            NoteNumberingStyle::ZodiacTwo,
        ),
        (
            "sftnnzodiacl",
            "saftnnzodiacl",
            NoteNumberingStyle::ZodiacThree,
        ),
    ];

    for (footnote, endnote, style) in cases {
        let source = format!(r#"{{\rtf1\sectd\{footnote}\{endnote} X}}"#);
        let document = RtfDocument::parse(&source).unwrap();
        let options = document.sections()[0].properties.note_options;
        assert_eq!(options.footnote_numbering, Some(style));
        assert_eq!(options.endnote_numbering, Some(style));

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
fn preserves_inheritance_and_sectd_reset() {
    let inherited =
        RtfDocument::parse(r#"{\rtf1\sectd\sftnbj\sftnnchi First\sect\sftnstart2 Second}"#)
            .unwrap();
    assert_eq!(inherited.sections().len(), 2);
    let first = inherited.sections()[0].properties.note_options;
    let second = inherited.sections()[1].properties.note_options;
    assert_eq!(second.footnote_placement, first.footnote_placement);
    assert_eq!(second.footnote_numbering, first.footnote_numbering);
    assert_eq!(second.footnote_start, Some(2));

    let reset =
        RtfDocument::parse(r#"{\rtf1\sectd\sftnbj\sftnnchi First\sect\sectd Second}"#).unwrap();
    assert_eq!(reset.sections().len(), 2);
    assert!(reset.sections()[1].properties.note_options.is_empty());
}

#[test]
fn rejects_invalid_values_and_non_root_or_late_section_note_options() {
    let malformed = [
        r#"{\rtf1\sectd\sftnstart0 Body}"#,
        r#"{\rtf1\sectd\sftnstart-1 Body}"#,
        r#"{\rtf1\sectd\saftnstart0 Body}"#,
        r#"{\rtf1\sectd\saftnstart-1 Body}"#,
        r#"{\rtf1\sectd\sftnstart2147483648 Body}"#,
        r#"{\rtf1\sectd Body\sftnbj}"#,
        r#"{\rtf1\sectd\'41\sftnbj}"#,
        r#"{\rtf1\sectd\u65?\sftnbj}"#,
        r#"{\rtf1\sectd{\sftnbj}Body}"#,
        r#"{\rtf1\sectd{\*\sftnbj}Body}"#,
        r#"{\rtf1\sectd{\header\sftnbj X}Body}"#,
        r#"{\rtf1\sectd{\footer\saftnnar X}Body}"#,
        r#"{\rtf1\sectd{\annotation\sftnstart2 X}Body}"#,
        r#"{\rtf1\sectd{\footnote\sftnrstpg X}Body}"#,
        r#"{\rtf1\sectd{\field\sftnnar X}Body}"#,
        r#"{\rtf1\sectd{\object\saftnnar X}Body}"#,
        r#"{\rtf1\sectd{\sftnbj\bin2 AB}Body}"#,
        r#"{\rtf1\sectd{\header H}\sftnbj Body}"#,
    ];
    for source in malformed {
        assert!(
            RtfDocument::parse(source).is_err(),
            "accepted malformed section note options: {source}"
        );
    }

    let invalid = SectionNoteOptions {
        endnote_start: Some(0),
        ..SectionNoteOptions::default()
    };
    assert!(
        RtfWriter::new(Vec::new())
            .write_section_note_options(&invalid)
            .is_err()
    );
}

#[test]
fn document_note_options_do_not_materialize_section_overrides() {
    let document = RtfDocument::parse(r#"{\rtf1\ftnnar\aftnnruc\sectd Body}"#).unwrap();
    assert!(document.sections()[0].properties.note_options.is_empty());
}

#[test]
fn permits_inert_section_note_controls_only_in_body_field_results() {
    let source = concat!(
        r#"{\rtf1{\field{\*\fldinst TOC}{\fldrslt Outer "#,
        r#"{\field{\*\fldinst PAGEREF mark}{\fldrslt 1\sectd\sftnbj}}"#,
        r#"\sectd\sftnnar}}}"#,
    );
    let document = RtfDocument::parse(source).unwrap();
    assert_eq!(document.fields().len(), 2);
    assert!(
        document
            .sections()
            .iter()
            .all(|section| section.properties.note_options.is_empty())
    );

    let malformed = [
        r#"{\rtf1{\field{\*\fldinst TOC \sftnbj}{\fldrslt X}}}"#,
        r#"{\rtf1{\header{\field{\*\fldinst X}{\fldrslt\sftnbj X}}}Body}"#,
        r#"{\rtf1{\footer{\field{\*\fldinst X}{\fldrslt\sftnbj X}}}Body}"#,
        r#"{\rtf1{\annotation{\field{\*\fldinst X}{\fldrslt\sftnbj X}}}Body}"#,
        r#"{\rtf1{\footnote{\field{\*\fldinst X}{\fldrslt\sftnbj X}}}Body}"#,
        r#"{\rtf1{\object{\result{\field{\*\fldinst X}{\fldrslt\sftnbj X}}}}Body}"#,
        r#"{\rtf1{\pict{\field{\*\fldinst X}{\fldrslt\sftnbj X}}}Body}"#,
        r#"{\rtf1{\shp{\field{\*\fldinst X}{\fldrslt\sftnbj X}}}Body}"#,
    ];
    for source in malformed {
        assert!(
            RtfDocument::parse(source).is_err(),
            "accepted section note control in a non-body field result: {source}"
        );
    }
}

#[test]
fn permits_section_note_controls_in_explicit_body_section_format_snapshots() {
    let document =
        RtfDocument::parse(r#"{\rtf1\sectd\sftnnar Before{\sectd\sftnbj{\b snapshot}}After}"#)
            .unwrap();
    assert_eq!(document.text(), "BeforesnapshotAfter");

    for source in [
        r#"{\rtf1\sectd{\sftnbj}Body}"#,
        r#"{\rtf1\sectd{\*\sectd\sftnbj}Body}"#,
        r#"{\rtf1\sectd{\header{\sectd\sftnbj}H}Body}"#,
        r#"{\rtf1\sectd{\object{\sectd\sftnbj}}Body}"#,
    ] {
        assert!(
            RtfDocument::parse(source).is_err(),
            "accepted non-body section-format snapshot: {source}"
        );
    }
}

#[test]
fn permits_section_note_controls_in_explicit_root_section_format_runs() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1\sectd Before\sectd\ltrsect\psz9\linex0\headery0"#,
        r#"\footery454\colsx708\endnhere\sectlinegrid360\sectdefaultcl"#,
        r#"\sftnbj After}"#,
    ))
    .unwrap();
    assert_eq!(document.text(), "BeforeAfter");
    assert_eq!(
        document.sections()[0]
            .properties
            .note_options
            .footnote_placement,
        Some(SectionFootnotePlacement::BottomOfPage)
    );

    for source in [
        r#"{\rtf1\sectd Before\sectd\ltrsect Boundary\sftnbj}"#,
        r#"{\rtf1\sectd Before\sectd{\b boundary}\sftnbj}"#,
        r#"{\rtf1\sectd Before\sectd\sbk\sftnbj}"#,
    ] {
        assert!(
            RtfDocument::parse(source).is_err(),
            "accepted section note control after a root section-format run closed: {source}"
        );
    }
}

#[test]
fn permits_inert_section_format_runs_at_direct_header_footer_level() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1\sectd\sftntj"#,
        r#"{\headerr Header{\field{\*\fldinst PAGE}{\fldrslt 2}}"#,
        r#"\sectd\ltrsect\linex0\endnhere\sectdefaultcl\sftnbj{\b -}}"#,
        r#"{\footerl Footer\sectd\ltrsect\sftnbj{\i -}}Body}"#,
    ))
    .unwrap();
    assert_eq!(document.text(), "Body");
    assert_eq!(
        document.sections()[0]
            .properties
            .note_options
            .footnote_placement,
        Some(SectionFootnotePlacement::BeneathText)
    );
}

#[test]
fn permits_inert_section_format_runs_at_direct_field_instruction_level() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1{\field{\*\fldinst IF {\b nested}"#,
        r#"\sectd\ltrsect\psz9\linex0\sectdefaultcl\sftnbj "x"}"#,
        r#"{\fldrslt value}}}"#,
    ))
    .unwrap();
    assert_eq!(document.fields().len(), 1);
    assert!(
        document
            .sections()
            .iter()
            .all(|section| section.properties.note_options.is_empty())
    );

    let nested = r#"{\rtf1{\field{\*\fldinst IF {\sectd\sftnbj} x}{\fldrslt value}}}"#;
    assert!(RtfDocument::parse(nested).is_err());
}

#[test]
fn ignores_section_note_controls_after_the_parsed_document_group() {
    let document = RtfDocument::parse(r#"{\rtf1 Body}\sectd\sftnbj"#).unwrap();
    assert_eq!(document.text(), "Body");
    assert!(
        document
            .sections()
            .iter()
            .all(|section| section.properties.note_options.is_empty())
    );

    assert!(RtfDocument::parse(r#"{\rtf1 Body\sftnbj}"#).is_err());
}

#[test]
fn parses_named_libreoffice_section_note_option_fixtures() {
    let uiwriter = include_bytes!(
        "../../../test-data/libreoffice-core/sw/qa/extras/uiwriter/data/tdf147006.rtf"
    );
    let document = RtfDocument::parse_bytes(uiwriter).unwrap();
    let options = document.sections()[0].properties.note_options;
    assert_eq!(options.footnote_numbering, Some(NoteNumberingStyle::Arabic));
    assert_eq!(
        options.endnote_numbering,
        Some(NoteNumberingStyle::LowercaseRoman)
    );

    let tracking = include_bytes!(
        "../../../test-data/libreoffice-core/sw/qa/extras/rtfexport/data/text-change-tracking.rtf"
    );
    let document = RtfDocument::parse_bytes(tracking).unwrap();
    let options = document.sections()[0].properties.note_options;
    assert_eq!(options.footnote_numbering, Some(NoteNumberingStyle::Arabic));
    assert_eq!(
        options.endnote_numbering,
        Some(NoteNumberingStyle::LowercaseRoman)
    );

    let header = include_bytes!(
        "../../../test-data/libreoffice-core/sw/qa/writerfilter/dmapper/data/NoFirstPageHeadFooter.rtf"
    );
    let document = RtfDocument::parse_bytes(header).unwrap();
    assert!(document.sections().iter().any(|section| {
        section.properties.note_options.footnote_placement
            == Some(SectionFootnotePlacement::BottomOfPage)
    }));
}
