use litchi_rtf::{
    MAX_SECTION_LINE_GRID_TWIPS, RtfDocument, RtfWriter, Section, SectionDocumentGrid,
    SectionDocumentGridType,
};

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn parses_line_grid_and_grid_type_and_round_trips_in_canonical_order() {
    let source = r#"{\rtf1\sectd\sectlinegrid360{\sectspecifyl}Body}"#;
    let document = RtfDocument::parse(source).unwrap();
    let grid = document.sections()[0].properties.document_grid;
    assert_eq!(
        grid,
        SectionDocumentGrid {
            line_grid: Some(360),
            grid_type: Some(SectionDocumentGridType::LinesAndCharacters),
        }
    );
    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    assert!(serialized.contains(r#"\sectlinegrid360\sectspecifyl"#));
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.sections()[0].properties.document_grid, grid);
}

#[test]
fn parses_bundled_libreoffice_document_grid_fixture() {
    let fixture = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/libreoffice-core/sw/qa/extras/rtfexport/data/cjklist24.rtf"
    ));
    let document = RtfDocument::parse_bytes(fixture).unwrap();
    assert!(document.sections().iter().any(|section| {
        section.properties.document_grid
            == SectionDocumentGrid {
                line_grid: Some(360),
                grid_type: Some(SectionDocumentGridType::LinesAndCharacters),
            }
    }));

    let fixture = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/libreoffice-core/sw/qa/extras/rtfimport/data/tdf165923.rtf"
    ));
    let document = RtfDocument::parse_bytes(fixture).unwrap();
    assert!(document.sections().iter().any(|section| {
        section.properties.document_grid.line_grid == Some(360)
            && section.properties.document_grid.grid_type.is_none()
    }));
}

#[test]
fn parses_all_grid_types() {
    for (control, expected) in [
        ("sectspecifyl", SectionDocumentGridType::LinesAndCharacters),
        ("sectspecifycl", SectionDocumentGridType::CharactersOnly),
        ("sectspecifygen", SectionDocumentGridType::Default),
    ] {
        let source = format!(r#"{{\rtf1\{control} X}}"#);
        let document = RtfDocument::parse(&source).unwrap();
        assert_eq!(
            document.sections()[0].properties.document_grid.grid_type,
            Some(expected),
            "wrong grid type for {control}"
        );
        let output = write(&document);
        let serialized = String::from_utf8(output).unwrap();
        assert!(
            serialized.contains(&format!(r#"\{control}"#)),
            "missing {control} in {serialized}"
        );
    }
}

#[test]
fn preserves_missing_parameter_default_and_sectd_reset() {
    let bare = RtfDocument::parse(r#"{\rtf1\sectlinegrid X}"#).unwrap();
    assert_eq!(
        bare.sections()[0].properties.document_grid.line_grid,
        Some(360)
    );

    let zero = RtfDocument::parse(r#"{\rtf1\sectlinegrid0 X}"#).unwrap();
    assert_eq!(
        zero.sections()[0].properties.document_grid.line_grid,
        Some(0)
    );

    let reset = RtfDocument::parse(r#"{\rtf1\sectlinegrid312\sectspecifycl X\sectd Y}"#).unwrap();
    assert!(reset.sections()[0].properties.document_grid.is_empty());
}

#[test]
fn inert_destinations_do_not_change_document_grid() {
    for source in [
        r#"{\rtf1{\*\unknown\sectlinegrid312\sectspecifyl}Body}"#,
        r#"{\rtf1{\header\sectlinegrid312\sectspecifycl Header}Body}"#,
        r#"{\rtf1{\footer\sectlinegrid312 Footer}Body}"#,
        r#"{\rtf1{\field{\*\fldinst IF \sectlinegrid312 1 1}{\fldrslt visible}}Body}"#,
        r#"{\rtf1 A{\footnote\sectlinegrid312\sectspecifygen Note}Body}"#,
        r#"{\rtf1{\object\sectlinegrid312}Body}"#,
    ] {
        let document = RtfDocument::parse(source).unwrap();
        assert!(
            document
                .sections()
                .iter()
                .all(|section| section.properties.document_grid.is_empty()),
            "destination leaked document grid: {source}"
        );
    }
}

#[test]
fn rejects_out_of_range_pitch_and_invalid_public_writer_state() {
    for source in [
        r#"{\rtf1\sectlinegrid-1 X}"#.to_string(),
        format!(
            r#"{{\rtf1\sectlinegrid{} X}}"#,
            MAX_SECTION_LINE_GRID_TWIPS + 1
        ),
    ] {
        assert!(RtfDocument::parse(&source).is_err(), "accepted {source}");
    }

    let mut section = Section::new();
    section.properties.document_grid.line_grid = Some(-1);
    let mut output = Vec::new();
    assert!(RtfWriter::new(&mut output).write_section(&section).is_err());
}
