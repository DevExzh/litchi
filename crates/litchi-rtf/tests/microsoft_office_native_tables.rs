//! Stored nested-table geometry and bounded source edits on native Word RTF.
use litchi_rtf::{
    CellStoryEvent, FloatingTablePosition, Package, Table, TableCellCoordinate, TableCellMergeRole,
    TableCellPath, TableIndentUnit, TablePreferredWidthUnit,
};

const FILE: &str = "word-tables-layout.rtf";
fn bytes() -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/word")
            .join(FILE),
    )
    .unwrap()
}

fn check(table: &Table<'_>, fruit: &str) {
    assert_eq!(table.row_count(), 5);
    assert!(table.direction().is_none());
    let expected: [&[&str]; 5] = [
        &["Item", "Quantity", "Price", "Notes"],
        &[fruit, "3", "1.25", "Fruit"],
        &["Pears", "5", "2.50", "Unicode 中文"],
        &["Tea", "2", "4.75", "Nested table follows\n"],
        &["Merged summary\nTABLE_END"],
    ];
    for (index, (row, values)) in table.rows().iter().zip(expected).enumerate() {
        assert_eq!(
            row.cells()
                .iter()
                .map(|cell| cell.text())
                .collect::<Vec<_>>(),
            values
        );
        assert_eq!(row.layout().header, index == 0);
        let geometry = row.geometry();
        assert_eq!(geometry.half_gap_twips(), Some(108));
        assert_eq!(geometry.left_edge_twips(), Some(0));
        assert_eq!(
            geometry.preferred_width().unwrap().unit(),
            TablePreferredWidthUnit::Percent
        );
        assert_eq!(geometry.preferred_width().unwrap().value(), Some(5000));
        assert!(geometry.auto_fit());
        assert_eq!(geometry.indent().unwrap().unit(), TableIndentUnit::Percent);
        assert_eq!(geometry.indent().unwrap().value(), 0);
        // These are in-flow native tables; no floating-table layout is inferred.
        assert_eq!(row.positioning(), &FloatingTablePosition::default());
        for (column, cell) in row.cells().iter().enumerate() {
            assert_eq!(
                cell.right_boundary(),
                Some(if index == 4 {
                    9360
                } else {
                    [2340, 4680, 7020, 9360][column]
                })
            );
            assert_eq!(
                cell.preferred_width().unwrap().unit(),
                TablePreferredWidthUnit::Percent
            );
            assert_eq!(
                cell.preferred_width().unwrap().value(),
                Some(if index == 4 { 5000 } else { 1250 })
            );
            // Word represents the summary as one wide cell, without clmgf/clmrg.
            assert_eq!(cell.merge().horizontal, None);
            assert_eq!(cell.merge().vertical, None);
            assert_eq!(
                cell.nested_tables().len(),
                usize::from(index == 3 && column == 3)
            );
        }
    }
    let owner = &table.rows()[3].cells()[3];
    assert_eq!(owner.story_events(), [CellStoryEvent::NestedTable(0)]);
    let nested = &owner.nested_tables()[0];
    assert_eq!(nested.text_offset, "Nested table follows\n".len());
    assert_eq!(nested.table.row_count(), 2);
    for (index, row) in nested.table.rows().iter().enumerate() {
        assert_eq!(row.cells().len(), 2);
        assert_eq!(row.geometry().left_edge_twips(), Some(5));
        assert_eq!(
            row.geometry().preferred_width().unwrap().unit(),
            TablePreferredWidthUnit::Auto
        );
        assert_eq!(row.geometry().preferred_width().unwrap().value(), None);
        assert_eq!(
            row.cells()[0].merge().vertical,
            Some(if index == 0 {
                TableCellMergeRole::First
            } else {
                TableCellMergeRole::Continuation
            })
        );
        assert_eq!(row.cells()[1].merge().vertical, None);
        let expected = if index == 0 {
            ["N1\n10", "N2"]
        } else {
            ["", "20"]
        };
        for (column, (cell, text)) in row.cells().iter().zip(expected).enumerate() {
            assert_eq!(cell.text(), text);
            assert_eq!(cell.right_boundary(), Some([1062, 2119][column]));
            assert_eq!(
                cell.preferred_width().unwrap().unit(),
                TablePreferredWidthUnit::Twips
            );
            assert_eq!(cell.preferred_width().unwrap().value(), Some(1057));
            assert!(cell.nested_tables().is_empty());
            assert_eq!(cell.merge().horizontal, None);
        }
    }
}

#[test]
fn native_nested_table_keeps_vertical_merge_roles_width_units_and_parent_insertion() {
    let original = bytes();
    let source = Package::from_bytes(&original).unwrap();
    assert_eq!(source.tables().len(), 1);
    check(&source.tables()[0], "Apples");
    let detached = source.tables()[0].clone().into_owned();
    let file = tempfile::NamedTempFile::with_suffix(".rtf").unwrap();
    source.save(file.path()).unwrap();
    assert_eq!(std::fs::read(file.path()).unwrap(), original);
    let reopened = Package::open(file.path()).unwrap();
    check(&reopened.tables()[0], "Apples");
    drop(source);
    check(&detached, "Apples");
    assert_eq!(bytes(), original);
}

#[test]
fn native_outer_cell_edit_preserves_nested_merge_geometry_and_rejected_staging() {
    use litchi_rtf::edit::Error;
    let original = bytes();
    let source = Package::from_bytes(&original).unwrap();
    let mut edit = source.edit();
    edit.set_table_cell_text(TableCellPath::outer(0, 1, 0), "Peaches")
        .unwrap();
    // Shortening the parent text past the nested-table anchor is invalid.
    assert!(
        matches!(edit.set_table_cell_text(TableCellPath::outer(0, 3, 3), "short"),
        Err(Error::Rtf(litchi_rtf::Error::MalformedDocument(reason)))
        if reason == "RTF table-cell story event is outside a UTF-8 boundary")
    );
    let commit = edit.commit().unwrap();
    let file = tempfile::NamedTempFile::with_suffix(".rtf").unwrap();
    commit.package().save(file.path()).unwrap();
    let output = std::fs::read(file.path()).unwrap();
    let text = std::str::from_utf8(&original).unwrap();
    assert_eq!(text.matches("Apples").count(), 1);
    assert_eq!(output, text.replace("Apples", "Peaches").as_bytes());
    let reopened = Package::open(file.path()).unwrap();
    check(&reopened.tables()[0], "Peaches");
    assert_eq!(reopened.text(), source.text());
    assert_eq!(reopened.notes().len(), 1);
    assert_eq!(
        reopened.notes()[0].content,
        " ENDNOTE_TEXT_002: native endnote value 84."
    );
    assert!(!reopened.notes()[0].is_footnote);
    assert_eq!(reopened.notes()[0].position, source.notes()[0].position);
    assert_eq!(
        reopened
            .apply(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        original
    );
    assert_eq!(source.to_bytes().unwrap(), original);
    assert_eq!(bytes(), original);
}

#[test]
fn native_nested_cell_edit_refuses_unproven_source_publication_without_mutating_input() {
    use litchi_core::edit::RefusalKind;
    use litchi_rtf::edit::Error;
    let original = bytes();
    let source = Package::from_bytes(&original).unwrap();
    let mut edit = source.edit();
    edit.set_table_cell_text(
        TableCellPath::outer(0, 3, 3).with_nested(TableCellCoordinate {
            table_index: 0,
            row_index: 0,
            cell_index: 1,
        }),
        "Nested cell edited",
    )
    .unwrap();
    let Err(error) = edit.commit() else {
        panic!("nested edit unexpectedly published")
    };
    assert!(
        matches!(&error, Error::Refused(refusal)
        if refusal.kind() == RefusalKind::UnsupportedContent
        && refusal.reason() == "canonical destination publication would relocate body tables"),
        "{error:?}"
    );
    check(&source.tables()[0], "Apples");
    assert_eq!(source.to_bytes().unwrap(), original);
    assert_eq!(bytes(), original);
}
