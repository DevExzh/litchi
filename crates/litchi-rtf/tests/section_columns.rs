use litchi_rtf::{
    MAX_SECTION_COLUMNS, RtfDocument, RtfWriter, Section, SectionColumn, SectionColumns,
};

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn parses_variable_columns_with_group_inheritance_and_round_trips_canonically() {
    let source = concat!(
        r#"{\rtf1\sectd\cols2\linebetcol\colno1"#,
        r#"{\colw3000\colsr240}\colno2\colw4000 Body}"#,
    );
    let document = RtfDocument::parse(source).unwrap();
    let columns = &document.sections()[0].properties.columns;
    assert_eq!(columns.count, 2);
    assert!(columns.separator);
    assert_eq!(
        columns.explicit,
        vec![
            SectionColumn {
                width: 3000,
                space_after: Some(240)
            },
            SectionColumn {
                width: 4000,
                space_after: None
            },
        ]
    );

    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    assert!(
        serialized
            .contains(r#"\cols2\linebetcol\colsx720\colno1\colw3000\colsr240\colno2\colw4000"#)
    );
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.sections()[0].properties.columns, *columns);
}

#[test]
fn parses_bundled_libreoffice_equal_and_variable_column_fixtures() {
    let equal = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/rtf/column-break.rtf"
    ));
    let equal = RtfDocument::parse_bytes(equal).unwrap();
    assert_eq!(equal.sections()[0].properties.columns.count, 2);
    assert!(!equal.sections()[0].properties.columns.is_variable());

    let variable = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/libreoffice-core/sw/qa/extras/rtfexport/data/tdf100507.rtf"
    ));
    let variable = RtfDocument::parse_bytes(variable).unwrap();
    assert!(variable.sections().iter().any(|section| {
        section.properties.columns.explicit
            == [SectionColumn {
                width: 9032,
                space_after: None,
            }]
    }));
}

#[test]
fn sectd_resets_columns_and_ignorable_destinations_are_inert() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1{\*\unknown\cols3\linebetcol\colno1\colw100}"#,
        r#"\cols2\colsx200\linebetcol\sectd Body}"#,
    ))
    .unwrap();
    assert_eq!(
        document.sections()[0].properties.columns,
        SectionColumns::default()
    );
}

#[test]
fn public_builders_and_writer_enforce_column_bounds() {
    let columns = SectionColumns::variable(
        vec![
            SectionColumn::new(2000, Some(180)).unwrap(),
            SectionColumn::new(2400, None).unwrap(),
        ],
        720,
        true,
    )
    .unwrap();
    assert_eq!(columns.count, 2);
    assert!(SectionColumns::equal(MAX_SECTION_COLUMNS, 0, false).is_ok());
    assert!(SectionColumns::equal(MAX_SECTION_COLUMNS + 1, 0, false).is_err());

    let mut section = Section::new();
    section.properties.columns.count = 2;
    section.properties.columns.explicit = vec![SectionColumn {
        width: 1000,
        space_after: None,
    }];
    let mut output = Vec::new();
    assert!(RtfWriter::new(&mut output).write_section(&section).is_err());
}

#[test]
fn rejects_malformed_or_incomplete_explicit_column_sequences() {
    for source in [
        r#"{\rtf1\cols0 X}"#,
        r#"{\rtf1\cols65 X}"#,
        r#"{\rtf1\cols2\colno X}"#,
        r#"{\rtf1\cols2\colw100 X}"#,
        r#"{\rtf1\cols2\colno2\colw100 X}"#,
        r#"{\rtf1\cols2\colno1\colw X}"#,
        r#"{\rtf1\cols2\colno1\colw0 X}"#,
        r#"{\rtf1\cols2\colno1\colw100\colw200 X}"#,
        r#"{\rtf1\cols2\colno1\colsr20\colw100 X}"#,
        r#"{\rtf1\cols2\colno1\colw100\colsr-1 X}"#,
        r#"{\rtf1\cols2\colno1\colw100 X}"#,
        r#"{\rtf1\cols2\colno1{\colno2}\colw100\colno2\colw200 X}"#,
    ] {
        assert!(RtfDocument::parse(source).is_err(), "accepted {source}");
    }
}
