use litchi_rtf::{
    RtfDocument, RtfWriter, TableAutoformatFlag, TableAutoformatFlags, TableRowBandIndex,
    TableRowBanding,
};

fn write(document: &RtfDocument<'_>) -> String {
    let mut bytes = Vec::new();
    RtfWriter::new(&mut bytes).write_document(document).unwrap();
    String::from_utf8(bytes).unwrap()
}

#[test]
fn parses_all_flags_and_canonically_round_trips() {
    let source = r#"{\rtf1\trowd\irow7\irowband-1\tbllknocolband\tbllklastcol\tbllkborder\tbllkshading\tbllkfont\tbllkcolor\tbllkbestfit\tbllkhdrrows\tbllklastrow\tbllkhdrcols\tbllknorowband\lastrow\cellx1000\intbl X\cell\row}"#;
    let document = RtfDocument::parse(source).unwrap();
    let row = &document.tables()[0].rows()[0];
    assert_eq!(
        row.banding(),
        TableRowBanding {
            row_index: Some(7),
            band_index: Some(TableRowBandIndex::Header),
            last_row: true
        }
    );
    for flag in [
        TableAutoformatFlag::Border,
        TableAutoformatFlag::Shading,
        TableAutoformatFlag::Font,
        TableAutoformatFlag::Color,
        TableAutoformatFlag::BestFit,
        TableAutoformatFlag::HeaderRows,
        TableAutoformatFlag::LastRow,
        TableAutoformatFlag::HeaderColumns,
        TableAutoformatFlag::LastColumn,
        TableAutoformatFlag::NoRowBanding,
        TableAutoformatFlag::NoColumnBanding,
    ] {
        assert!(row.autoformat_flags().contains(flag));
    }
    let first = write(&document);
    assert!(first.contains(r#"\irow7\irowband-1\tbllkborder\tbllkshading\tbllkfont\tbllkcolor\tbllkbestfit\tbllkhdrrows\tbllklastrow\tbllkhdrcols\tbllklastcol\tbllknorowband\tbllknocolband\lastrow"#));
    let reparsed = RtfDocument::parse(&first).unwrap();
    assert_eq!(reparsed.tables()[0].rows()[0].banding(), row.banding());
    assert_eq!(
        reparsed.tables()[0].rows()[0].autoformat_flags(),
        row.autoformat_flags()
    );
    assert_eq!(write(&reparsed), first);
}

#[test]
fn restores_groups_resets_trowd_and_preserves_owned_rows() {
    let source = r#"{\rtf1\trowd{\irow9\irowband2\tbllkborder\lastrow}\cellx1000\intbl A\cell\row\trowd\irow1\irowband0\tbllkhdrrows\cellx1000\intbl B\cell\row\trowd\cellx1000\intbl C\cell\row}"#;
    let document = RtfDocument::parse(source).unwrap();
    let expected_banding;
    let expected_flags;
    {
        let rows = document.tables()[0].rows();
        assert_eq!(rows[0].banding(), TableRowBanding::default());
        assert_eq!(rows[0].autoformat_flags(), TableAutoformatFlags::default());
        assert_eq!(rows[1].banding().row_index, Some(1));
        assert!(
            rows[1]
                .autoformat_flags()
                .contains(TableAutoformatFlag::HeaderRows)
        );
        assert_eq!(rows[2].banding(), TableRowBanding::default());
        expected_banding = rows[1].banding();
        expected_flags = rows[1].autoformat_flags();
    }
    let owned = RtfDocument::parse_bytes(source.as_bytes()).unwrap();
    assert_eq!(owned.tables()[0].rows()[1].banding(), expected_banding);
    assert_eq!(
        owned.tables()[0].rows()[1].autoformat_flags(),
        expected_flags
    );
}

#[test]
fn snapshots_end_defined_nested_rows() {
    let source = r#"{\rtf1\trowd\cellx3000\intbl\itap2 X\nestcell{\*\nesttableprops\itap2\trowd\irow4\irowband3\tbllknorowband\lastrow\cellx1000\nestrow}\intbl\itap1\cell\row}"#;
    let document = RtfDocument::parse(source).unwrap();
    let row = &document.tables()[0].rows()[0].cells()[0].nested_tables()[0]
        .table
        .rows()[0];
    assert_eq!(
        row.banding(),
        TableRowBanding {
            row_index: Some(4),
            band_index: Some(TableRowBandIndex::Row(3)),
            last_row: true
        }
    );
    assert!(
        row.autoformat_flags()
            .contains(TableAutoformatFlag::NoRowBanding)
    );
    let expected_banding = row.banding();
    let expected_flags = row.autoformat_flags();
    let reparsed = RtfDocument::parse(&write(&document)).unwrap();
    let round_trip = &reparsed.tables()[0].rows()[0].cells()[0].nested_tables()[0]
        .table
        .rows()[0];
    assert_eq!(round_trip.banding(), expected_banding);
    assert_eq!(round_trip.autoformat_flags(), expected_flags);
}

#[test]
fn rejects_missing_parameters_parameters_on_flags_duplicates_and_out_of_range_values() {
    for controls in [
        "irow",
        "irow-1",
        "irow65536",
        "irowband",
        "irowband-2",
        "irowband65536",
        "lastrow1",
        "tbllkborder1",
        "irow0\\irow1",
        "irowband0\\irowband1",
        "lastrow\\lastrow",
        "tbllkfont\\tbllkfont",
    ] {
        let source = format!(r#"{{\rtf1\trowd\{controls}\cellx1000\intbl X\cell\row}}"#);
        assert!(RtfDocument::parse(&source).is_err(), "accepted {controls}");
    }
    assert!(
        RtfDocument::parse(r#"{\rtf1\trowd\irow65535\irowband65535\cellx1000\intbl X\cell\row}"#)
            .is_ok()
    );
}

#[test]
fn parses_real_libreoffice_row_banding_fixture() {
    let bytes = include_bytes!(
        "../../../3rdparty/libreoffice-core/sw/qa/extras/rtfimport/data/tdf148544.rtf"
    );
    let document = RtfDocument::parse_bytes(bytes).unwrap();
    let rows = document
        .tables()
        .iter()
        .flat_map(|table| table.rows())
        .collect::<Vec<_>>();
    assert!(rows.iter().any(|row| row.banding().row_index.is_some()));
    assert!(
        rows.iter()
            .any(|row| matches!(row.banding().band_index, Some(TableRowBandIndex::Row(_))))
    );
    assert!(rows.iter().any(|row| row.banding().last_row));
}
