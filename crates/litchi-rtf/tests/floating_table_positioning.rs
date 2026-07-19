use litchi_rtf::{
    RtfDocument, RtfWriter, TableHorizontalPosition, TableHorizontalReference,
    TableVerticalPosition, TableVerticalReference,
};

#[test]
fn parses_full_family_and_round_trips_deterministically() {
    let source = r#"{\rtf1\trowd\tphmrg\tposxc\tpvpg\tposnegy-16\tdfrmtxtLeft187\tdfrmtxtRight188\tdfrmtxtTop189\tdfrmtxtBottom190\tabsnoovrlp1\cellx1000\intbl A\cell\row}"#;
    let document = RtfDocument::parse(source).unwrap();
    let position = document.tables()[0].rows()[0].positioning();
    assert_eq!(
        position.horizontal_reference,
        Some(TableHorizontalReference::Margin)
    );
    assert_eq!(
        position.horizontal_position,
        Some(TableHorizontalPosition::Center)
    );
    assert_eq!(
        position.vertical_reference,
        Some(TableVerticalReference::Page)
    );
    assert_eq!(
        position.vertical_position,
        Some(TableVerticalPosition::NegativeOffset(-16))
    );
    assert_eq!(
        (
            position.wrap_distances.left,
            position.wrap_distances.right,
            position.wrap_distances.top,
            position.wrap_distances.bottom
        ),
        (Some(187), Some(188), Some(189), Some(190))
    );
    assert!(position.no_overlap);
    let mut first = Vec::new();
    RtfWriter::new(&mut first)
        .write_document(&document)
        .unwrap();
    let first = String::from_utf8(first).unwrap();
    assert!(first.contains("\\tphmrg\\tposxc\\tpvpg\\tposnegy-16\\tdfrmtxtLeft187\\tdfrmtxtRight188\\tdfrmtxtTop189\\tdfrmtxtBottom190\\tabsnoovrlp"));
    let reparsed = RtfDocument::parse(&first).unwrap();
    assert_eq!(reparsed.tables()[0].rows()[0].positioning(), position);
    let mut second = Vec::new();
    RtfWriter::new(&mut second)
        .write_document(&reparsed)
        .unwrap();
    assert_eq!(first, String::from_utf8(second).unwrap());
}

#[test]
fn parses_all_symbolic_positions_and_references() {
    for (word, expected) in [
        ("tposxc", TableHorizontalPosition::Center),
        ("tposxi", TableHorizontalPosition::Inside),
        ("tposxl", TableHorizontalPosition::Left),
        ("tposxo", TableHorizontalPosition::Outside),
        ("tposxr", TableHorizontalPosition::Right),
    ] {
        let source = format!("{{\\rtf1\\trowd\\{word}\\cellx1\\intbl X\\cell\\row}}");
        assert_eq!(
            RtfDocument::parse(&source).unwrap().tables()[0].rows()[0]
                .positioning()
                .horizontal_position,
            Some(expected)
        );
    }
    for (word, expected) in [
        ("tposyb", TableVerticalPosition::Bottom),
        ("tposyc", TableVerticalPosition::Center),
        ("tposyil", TableVerticalPosition::Inline),
        ("tposyin", TableVerticalPosition::Inside),
        ("tposyout", TableVerticalPosition::Outside),
        ("tposyt", TableVerticalPosition::Top),
    ] {
        let source = format!("{{\\rtf1\\trowd\\{word}\\cellx1\\intbl X\\cell\\row}}");
        assert_eq!(
            RtfDocument::parse(&source).unwrap().tables()[0].rows()[0]
                .positioning()
                .vertical_position,
            Some(expected)
        );
    }
    for (word, expected) in [
        ("tphcol", TableHorizontalReference::Column),
        ("tphmrg", TableHorizontalReference::Margin),
        ("tphpg", TableHorizontalReference::Page),
    ] {
        let source = format!("{{\\rtf1\\trowd\\{word}\\cellx1\\intbl X\\cell\\row}}");
        assert_eq!(
            RtfDocument::parse(&source).unwrap().tables()[0].rows()[0]
                .positioning()
                .horizontal_reference,
            Some(expected)
        );
    }
    for (word, expected) in [
        ("tpvmrg", TableVerticalReference::Margin),
        ("tpvpara", TableVerticalReference::Paragraph),
        ("tpvpg", TableVerticalReference::Page),
    ] {
        let source = format!("{{\\rtf1\\trowd\\{word}\\cellx1\\intbl X\\cell\\row}}");
        assert_eq!(
            RtfDocument::parse(&source).unwrap().tables()[0].rows()[0]
                .positioning()
                .vertical_reference,
            Some(expected)
        );
    }
}

#[test]
fn resets_groups_and_inert_destinations() {
    let source = r#"{\rtf1\trowd\tphpg\tposxr{\tphmrg\tposxl}\trowd{\*\unknown\tphpg\tposx999\tdfrmtxtLeft999\tabsnoovrlp1 ignored}\cellx1000\intbl A\cell\row}"#;
    let document = RtfDocument::parse(source).unwrap();
    assert!(document.tables()[0].rows()[0].positioning().is_empty());
}

#[test]
fn parses_real_libreoffice_fixtures() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../3rdparty/libreoffice-core/sw/qa/writerfilter/rtftok/data");
    let floating =
        RtfDocument::parse(&std::fs::read_to_string(root.join("floating-table.rtf")).unwrap())
            .unwrap();
    assert!(
        floating
            .tables()
            .iter()
            .flat_map(|table| table.rows())
            .any(|row| {
                let p = row.positioning();
                p.horizontal_reference == Some(TableHorizontalReference::Column)
                    && p.horizontal_position == Some(TableHorizontalPosition::Offset(20))
                    && p.vertical_reference == Some(TableVerticalReference::Paragraph)
                    && p.vertical_position == Some(TableVerticalPosition::Offset(10))
                    && p.wrap_distances.left == Some(30)
                    && p.wrap_distances.right == Some(40)
            })
    );
    let overlap = RtfDocument::parse(
        &std::fs::read_to_string(root.join("floattable-tbl-overlap.rtf")).unwrap(),
    )
    .unwrap();
    assert!(
        overlap
            .tables()
            .iter()
            .flat_map(|table| table.rows())
            .any(|row| row.positioning().no_overlap)
    );
}

#[test]
fn rejects_malformed_parameters_and_caps() {
    for source in [
        r#"{\rtf1\trowd\tposx X}"#,
        r#"{\rtf1\trowd\tposx-1 X}"#,
        r#"{\rtf1\trowd\tposx31681 X}"#,
        r#"{\rtf1\trowd\tposnegx1 X}"#,
        r#"{\rtf1\trowd\tposnegx-31681 X}"#,
        r#"{\rtf1\trowd\tposyc1 X}"#,
        r#"{\rtf1\trowd\tphmrg1 X}"#,
        r#"{\rtf1\trowd\tdfrmtxtLeft X}"#,
        r#"{\rtf1\trowd\tdfrmtxtLeft-1 X}"#,
        r#"{\rtf1\trowd\tdfrmtxtLeft31681 X}"#,
        r#"{\rtf1\trowd\tabsnoovrlp2 X}"#,
    ] {
        assert!(RtfDocument::parse(source).is_err(), "accepted {source}");
    }
    assert!(
        RtfDocument::parse(r#"{\rtf1\trowd\tabsnoovrlp0\cellx1\intbl X\cell\row}"#)
            .unwrap()
            .tables()[0]
            .rows()[0]
            .positioning()
            .is_empty()
    );
}
