use litchi_rtf::{
    MAX_SECTION_LINE_INCREMENT, MAX_SECTION_LINE_START, RtfDocument, RtfWriter, Section,
    SectionLineNumberRestart, SectionLineNumbering,
};

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn parses_grouped_line_numbering_and_round_trips_in_canonical_order() {
    let source = r#"{\rtf1\sectd\linemod5{\linex283}\linestarts2\linecont Body}"#;
    let document = RtfDocument::parse(source).unwrap();
    let numbering = document.sections()[0].properties.line_numbering;
    assert_eq!(
        numbering,
        SectionLineNumbering {
            increment: Some(5),
            distance: Some(283),
            start: Some(2),
            restart: Some(SectionLineNumberRestart::Continuous),
        }
    );
    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    assert!(serialized.contains(r#"\linemod5\linex283\linestarts2\linecont"#));
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.sections()[0].properties.line_numbering, numbering);
}

#[test]
fn parses_bundled_libreoffice_line_numbering_fixtures() {
    let continuous = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../3rdparty/libreoffice-core/sw/qa/extras/rtfexport/data/linenumbering.rtf"
    ));
    let continuous = RtfDocument::parse_bytes(continuous).unwrap();
    assert!(continuous.sections().iter().any(|section| {
        section.properties.line_numbering
            == SectionLineNumbering {
                increment: Some(5),
                distance: Some(283),
                start: None,
                restart: Some(SectionLineNumberRestart::Continuous),
            }
    }));

    let start = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../3rdparty/libreoffice-core/sw/qa/extras/rtfexport/data/tdf66543.rtf"
    ));
    let start = RtfDocument::parse_bytes(start).unwrap();
    assert!(start.sections().iter().any(|section| {
        section.properties.line_numbering.increment == Some(3)
            && section.properties.line_numbering.distance == Some(0)
            && section.properties.line_numbering.start == Some(2)
    }));
}

#[test]
fn preserves_disabled_offset_defaults_restart_modes_and_sectd_reset() {
    let offset_only = RtfDocument::parse(r#"{\rtf1\linex0 Body}"#).unwrap();
    assert_eq!(
        offset_only.sections()[0].properties.line_numbering,
        SectionLineNumbering {
            distance: Some(0),
            ..SectionLineNumbering::default()
        }
    );
    assert!(
        !offset_only.sections()[0]
            .properties
            .line_numbering
            .is_enabled()
    );

    let defaults =
        RtfDocument::parse(r#"{\rtf1\linemod\linex\linestarts\linerestart X\sectd Y}"#).unwrap();
    assert_eq!(
        defaults.sections()[0].properties.line_numbering,
        SectionLineNumbering::default()
    );

    for (control, expected) in [
        ("linerestart", SectionLineNumberRestart::Section),
        ("lineppage", SectionLineNumberRestart::Page),
        ("linecont", SectionLineNumberRestart::Continuous),
    ] {
        let source = format!(r#"{{\rtf1\linemod1\{control} X}}"#);
        assert_eq!(
            RtfDocument::parse(&source).unwrap().sections()[0]
                .properties
                .line_numbering
                .restart,
            Some(expected)
        );
    }
}

#[test]
fn inert_destinations_do_not_change_section_line_numbering() {
    for source in [
        r#"{\rtf1{\*\unknown\linemod9\linex900\linecont}Body}"#,
        r#"{\rtf1{\header\linemod8\linex800 Header}Body}"#,
        r#"{\rtf1{\footer\linemod8\linex800 Footer}Body}"#,
        r#"{\rtf1{\field{\*\fldinst IF \linemod7 1 1}{\fldrslt visible}}Body}"#,
        r#"{\rtf1 A{\footnote\linemod6\linex600 Note}Body}"#,
        r#"{\rtf1{\object\linemod5\linex500}Body}"#,
    ] {
        let document = RtfDocument::parse(source).unwrap();
        assert!(
            document
                .sections()
                .iter()
                .all(|section| section.properties.line_numbering.is_empty()),
            "destination leaked line numbering: {source}"
        );
    }
}

#[test]
fn rejects_out_of_range_values_and_invalid_public_writer_state() {
    for source in [
        r#"{\rtf1\linemod-1 X}"#.to_string(),
        format!(
            r#"{{\rtf1\linemod{} X}}"#,
            u32::from(MAX_SECTION_LINE_INCREMENT) + 1
        ),
        r#"{\rtf1\linex-1 X}"#.to_string(),
        r#"{\rtf1\linex31681 X}"#.to_string(),
        r#"{\rtf1\linestarts0 X}"#.to_string(),
        format!(r#"{{\rtf1\linestarts{} X}}"#, MAX_SECTION_LINE_START + 1),
    ] {
        assert!(RtfDocument::parse(&source).is_err(), "accepted {source}");
    }

    let mut section = Section::new();
    section.properties.line_numbering.increment = Some(0);
    let mut output = Vec::new();
    assert!(RtfWriter::new(&mut output).write_section(&section).is_err());
}
