use litchi_rtf::{Cell, RtfDocument, RtfWriter, Table};

fn nested_count(table: &Table<'_>) -> usize {
    table
        .rows()
        .iter()
        .flat_map(|row| row.cells())
        .map(|cell| {
            cell.nested_tables()
                .iter()
                .map(|entry| 1 + nested_count(&entry.table))
                .sum::<usize>()
        })
        .sum()
}
fn assert_table_eq(left: &Table<'_>, right: &Table<'_>) {
    assert_eq!(left.direction(), right.direction());
    assert_eq!(left.rows().len(), right.rows().len());
    for (a, b) in left.rows().iter().zip(right.rows()) {
        assert_eq!(a.direction(), b.direction());
        assert_eq!(a.padding(), b.padding());
        assert_eq!(a.spacing(), b.spacing());
        assert_eq!(a.positioning(), b.positioning());
        assert_eq!(a.cells().len(), b.cells().len());
        for (a, b) in a.cells().iter().zip(b.cells()) {
            assert_eq!(a.text(), b.text());
            assert_eq!(a.padding(), b.padding());
            assert_eq!(a.spacing(), b.spacing());
            assert_eq!(a.nested_tables().len(), b.nested_tables().len());
            for (a, b) in a.nested_tables().iter().zip(b.nested_tables()) {
                assert_eq!(a.text_offset, b.text_offset);
                assert_table_eq(&a.table, &b.table);
            }
        }
    }
}

fn sample() -> String {
    r#"{\rtf1\trowd\cellx5000\intbl\itap1 Before \intbl\itap2 Inner\nestcell\intbl\itap2\nestcell{\*\nesttableprops\trowd\cellx1000\cellx2000\nestrow}{\nonesttables ignored fallback}\intbl\itap1 After\cell\row}"#.into()
}

#[test]
fn parses_ordered_nested_content_empty_cells_and_round_trips() {
    let document = RtfDocument::parse(&sample()).unwrap();
    let outer = &document.tables()[0];
    let cell = &outer.rows()[0].cells()[0];
    assert_eq!(cell.text(), "Before After");
    assert_eq!(cell.nested_tables().len(), 1);
    assert_eq!(cell.nested_tables()[0].text_offset, 7);
    let nested = &cell.nested_tables()[0].table;
    assert_eq!(nested.rows()[0].cells()[0].text(), "Inner");
    assert_eq!(nested.rows()[0].cells()[1].text(), "");
    let mut first = Vec::new();
    RtfWriter::new(&mut first)
        .write_document(&document)
        .unwrap();
    let first = String::from_utf8(first).unwrap();
    assert!(first.contains("\\itap2 Inner\\nestcell"));
    assert!(first.contains("{\\*\\nesttableprops\\itap2\\trowd"));
    assert!(first.contains("\\nestrow}{\\nonesttables\\par}"));
    let reparsed = RtfDocument::parse(&first).unwrap();
    assert_table_eq(outer, &reparsed.tables()[0]);
    let mut second = Vec::new();
    RtfWriter::new(&mut second)
        .write_document(&reparsed)
        .unwrap();
    assert_eq!(first, String::from_utf8(second).unwrap());
}

#[test]
fn parses_real_libreoffice_end_defined_nested_table() {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../3rdparty/libreoffice-core/sw/qa/extras/rtfexport/data/tdf117268.rtf");
    let document = RtfDocument::parse(&std::fs::read_to_string(path).unwrap()).unwrap();
    assert!(
        document
            .tables()
            .iter()
            .any(|table| nested_count(table) > 0)
    );
    let nested = document
        .tables()
        .iter()
        .flat_map(|table| table.rows())
        .flat_map(|row| row.cells())
        .flat_map(Cell::nested_tables)
        .next()
        .unwrap();
    assert_eq!(nested.table.rows()[0].cells()[0].text().trim(), "Text 3");
}

#[test]
fn restores_groups_and_ignores_fallback_and_unknown_destinations() {
    let source = r#"{\rtf1\trowd\cellx2000\intbl\itap1 Outer{\itap2}{\*\unknown\itap32\nestcell bad}{\nonesttables\itap32\nestcell fallback}\cell\row}"#;
    let document = RtfDocument::parse(source).unwrap();
    assert_eq!(document.tables()[0].rows()[0].cells()[0].text(), "Outer");
    assert!(
        document.tables()[0].rows()[0].cells()[0]
            .nested_tables()
            .is_empty()
    );
}

#[test]
fn rejects_malformed_levels_destinations_and_boundaries() {
    for source in [
        r#"{\rtf1\itap X}"#,
        r#"{\rtf1\itap-1 X}"#,
        r#"{\rtf1\itap33 X}"#,
        r#"{\rtf1\trowd\cellx1\intbl\itap1 X\nestcell}"#,
        r#"{\rtf1\trowd\cellx1\intbl\itap2 X\nestrow}"#,
        r#"{\rtf1\trowd\cellx1\intbl\itap2 X{\nesttableprops\trowd\cellx1\nestrow}}"#,
        r#"{\rtf1\trowd\cellx1\intbl\itap3 X}"#,
        r#"{\rtf1\trowd\cellx1\intbl\itap2 X\nestcell{\nesttableprops\trowd\cellx1\cellx2\nestrow}}"#,
    ] {
        assert!(RtfDocument::parse(source).is_err(), "accepted {source}");
    }
}

#[test]
fn enforces_nested_floating_invariant_and_trowd_reset() {
    let source = r#"{\rtf1\trowd\cellx5000\intbl\itap2 A\nestcell{\nesttableprops\trowd\tposx1\cellx1000\nestrow}\intbl\itap2 B\nestcell{\nesttableprops\trowd\cellx1000\nestrow}\itap1\cell\row}"#;
    assert!(RtfDocument::parse(source).is_err());
}

#[test]
fn enforces_nested_cell_cap() {
    let mut source = String::from("{\\rtf1\\trowd\\cellx5000\\intbl\\itap2 ");
    for _ in 0..=4096 {
        source.push_str("\\nestcell");
    }
    source.push_str("{\\nesttableprops\\trowd\\nestrow}}");
    assert!(RtfDocument::parse(&source).is_err());
}

#[test]
fn retains_nested_table_in_implicit_outer_cell() {
    let source = r#"{\rtf1\trowd\cellx5000\intbl\itap2 Inner\nestcell{\*\nesttableprops\itap2\trowd\cellx1000\nestrow}\intbl\itap1\row}"#;
    let document = RtfDocument::parse(source).unwrap();
    let cells = document.tables()[0].rows()[0].cells();
    assert_eq!(cells.len(), 1);
    assert_eq!(cells[0].text(), "");
    assert_eq!(cells[0].nested_tables().len(), 1);
    assert_eq!(
        cells[0].nested_tables()[0].table.rows()[0].cells()[0].text(),
        "Inner"
    );
}
