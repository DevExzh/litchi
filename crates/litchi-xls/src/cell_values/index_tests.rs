#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "index fixtures intentionally fail fast"
)]

use super::*;
use std::io::Cursor;

fn entry(row: u16, column: u8, value: f64) -> Entry {
    Entry {
        cell: Cell {
            reference: Reference { row, column },
            storage: Storage::Number,
            style: StyleIndex(0),
            value: Value::Number(value),
        },
        value_offset: Some(0),
        kind_offset: 0,
        sst_index: None,
    }
}

fn fixed(sheet: usize, row: u16, column: u8, value: f64) -> Change {
    Change {
        sheet,
        entry: usize::from(row),
        reference: Reference { row, column },
        storage: Storage::Number,
        value: Value::Number(value),
    }
}

fn package() -> Vec<u8> {
    let mut writer = crate::Builder::new();
    let sheet = writer.add_worksheet("Sheet1").unwrap();
    writer.write_number(sheet, 0, 0, 1.0).unwrap();
    writer.write_number(sheet, 0, 1, 2.0).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.build().unwrap().write_to(&mut output).unwrap();
    output.into_inner()
}

fn multi_sheet_package() -> Vec<u8> {
    let mut writer = crate::Builder::new();
    let first = writer.add_worksheet("First").unwrap();
    let second = writer.add_worksheet("Second").unwrap();
    for (row, column, value) in [(0, 0, 1.0), (1, 1, 2.0), (2, 2, 3.0)] {
        writer.write_number(first, row, column, value).unwrap();
    }
    for (row, column, value) in [(0, 0, 10.0), (1, 1, 20.0), (2, 2, 30.0)] {
        writer.write_number(second, row, column, value).unwrap();
    }
    let mut output = Cursor::new(Vec::new());
    writer.build().unwrap().write_to(&mut output).unwrap();
    output.into_inner()
}

fn reference(row: u16, column: u8) -> Reference {
    Reference { row, column }
}

#[test]
fn coordinate_index_preserves_source_order_and_refuses_duplicates() {
    let entries = vec![entry(9, 1, 9.0), entry(2, 4, 2.0), entry(7, 0, 7.0)];
    let index = CellIndex::from_entries(&entries).unwrap();
    assert_eq!(entries[0].cell.reference, Reference { row: 9, column: 1 });
    assert_eq!(
        index
            .unique_entry_index(&entries, Reference { row: 2, column: 4 })
            .unwrap(),
        Some(1)
    );
    assert!(index.any_rows(7, 7));
    assert!(index.any_columns(4, 4));

    let duplicate_entries = vec![entry(2, 4, 2.0), entry(2, 4, 3.0)];
    let duplicate_index = CellIndex::from_entries(&duplicate_entries).unwrap();
    assert!(matches!(
        duplicate_index.unique_entry_index(&duplicate_entries, Reference { row: 2, column: 4 }),
        Err(Error::Refused(_) | Error::PatchConflict { .. } | Error::Internal(_))
    ));
}

#[test]
fn fixed_index_handles_repeated_coordinates_and_bounded_ranges() {
    let changes = [fixed(0, 10, 1, 1.0), fixed(0, 2, 4, 2.0)];
    let mut index = FixedChangeIndex::default();
    index.reserve_for_join(changes.len()).unwrap();
    for (change, item) in changes.iter().enumerate() {
        index.insert(StagedCell {
            sheet: item.sheet,
            reference: item.reference,
            change,
        });
    }
    assert_eq!(index.first(0, Reference { row: 2, column: 4 }), Some(1));
    assert!(index.has_sheet(0));
    let range = CellRange::new(
        Reference { row: 2, column: 0 },
        Reference { row: 2, column: 10 },
    )
    .unwrap();
    assert!(index.any_in_range(0, range));
}

#[test]
fn structural_index_distinguishes_disjoint_cells_and_sheet_wide_axes() {
    let cell = StructuralChange::Cell {
        sheet: 0,
        reference: Reference { row: 3, column: 2 },
        before: None,
        after: None,
    };
    let axis = StructuralChange::Rows {
        sheet: 1,
        start: 4,
        count: 1,
        insert: true,
    };
    let rename = StructuralChange::RenameSheet {
        sheet: 0,
        before: "Sheet1".into(),
        after: "Renamed".into(),
    };
    let mut index = StructuralChangeIndex::default();
    index.reserve_one(&cell).unwrap();
    index.insert(&cell, 0);
    index.reserve_one(&axis).unwrap();
    index.insert(&axis, 1);
    index.reserve_one(&rename).unwrap();
    index.insert(&rename, 2);

    let disjoint = StructuralChange::Cell {
        sheet: 0,
        reference: Reference { row: 3, column: 3 },
        before: None,
        after: None,
    };
    assert!(!index.overlaps(&disjoint));
    assert!(index.overlaps(&cell));
    assert!(index.overlaps(&StructuralChange::Cell {
        sheet: 1,
        reference: Reference {
            row: 100,
            column: 100
        },
        before: None,
        after: None,
    }));
    assert!(index.overlaps(&StructuralChange::RenameSheet {
        sheet: 0,
        before: "Sheet1".into(),
        after: "Other".into(),
    }));
    assert!(index.equal_with(&[cell.clone(), axis, rename], &cell));
}

#[test]
fn joining_a_noop_then_staging_reuses_the_coordinate_slot() {
    let source = Snapshot::from_bytes(package(), Limits::default()).unwrap();
    let mut left = source.edit();
    left.set_number(
        Selector::Position(Position::new(0)),
        Reference { row: 0, column: 0 },
        1.0,
    )
    .unwrap();
    let mut incoming = source.edit();
    incoming
        .set_number(
            Selector::Position(Position::new(0)),
            Reference { row: 0, column: 0 },
            2.0,
        )
        .unwrap();
    left.join(incoming).unwrap();
    left.set_number(
        Selector::Position(Position::new(0)),
        Reference { row: 0, column: 0 },
        3.0,
    )
    .unwrap();
    let commit = left.commit().unwrap();
    let snapshot = commit.package();
    let cell = snapshot
        .worksheet(Selector::Position(Position::new(0)))
        .unwrap()
        .unwrap()
        .number(Reference { row: 0, column: 0 })
        .unwrap()
        .unwrap();
    assert_eq!(cell.value(), 3.0);
}

#[test]
fn joined_fixed_and_structural_cell_operations_conflict_in_both_orders() {
    let source = Snapshot::from_bytes(package(), Limits::default()).unwrap();
    let reference = Reference { row: 0, column: 0 };
    let style = StyleIndex::new(&source, 0).unwrap();

    let mut fixed = source.edit();
    fixed
        .set_number(Selector::Position(Position::new(0)), reference, 4.0)
        .unwrap();
    let mut structural = source.edit();
    structural
        .set_style(Selector::Position(Position::new(0)), reference, style)
        .unwrap();
    assert!(matches!(
        fixed.join(structural.clone()),
        Err(JoinError::StructuralConflicts(_))
    ));

    let mut structural_first = source.edit();
    structural_first
        .set_style(Selector::Position(Position::new(0)), reference, style)
        .unwrap();
    let mut fixed_second = source.edit();
    fixed_second
        .set_number(Selector::Position(Position::new(0)), reference, 4.0)
        .unwrap();
    assert!(matches!(
        structural_first.join(fixed_second),
        Err(JoinError::StructuralConflicts(_))
    ));
}

#[test]
fn joined_cloned_edits_rebuild_indexes_across_sheets_and_invert_exactly() {
    let source = Snapshot::from_bytes(multi_sheet_package(), Limits::default()).unwrap();
    let expected_order = [reference(0, 0), reference(1, 1), reference(2, 2)];
    for worksheet in source.worksheets() {
        let actual: Vec<_> = worksheet.cells().map(Cell::reference).collect();
        assert_eq!(actual, expected_order);
    }

    // Stage fixed edits in reverse sheet and coordinate order, join a second
    // reverse-order transaction, then clone the joined transaction and stage
    // one more target.  The clone must carry the rebuilt coordinate sidecar.
    let mut fixed = source.edit();
    fixed
        .set_numeric(Selector::Position(Position::new(1)), reference(2, 2), 33.0)
        .unwrap();
    fixed
        .set_numeric(Selector::Position(Position::new(0)), reference(1, 1), 22.0)
        .unwrap();
    let mut incoming_fixed = source.edit();
    incoming_fixed
        .set_numeric(Selector::Position(Position::new(1)), reference(0, 0), 11.0)
        .unwrap();
    incoming_fixed
        .set_numeric(Selector::Position(Position::new(0)), reference(0, 0), 12.0)
        .unwrap();
    fixed.join(incoming_fixed).unwrap();
    let mut fixed_clone = fixed.clone();
    fixed_clone
        .set_numeric(Selector::Position(Position::new(1)), reference(1, 1), 23.0)
        .unwrap();
    let fixed_commit = fixed_clone.commit().unwrap();
    let fixed_snapshot = fixed_commit.package();
    for (sheet, row, column, value) in [
        (0, 0, 0, 12.0),
        (0, 1, 1, 22.0),
        (0, 2, 2, 3.0),
        (1, 0, 0, 11.0),
        (1, 1, 1, 23.0),
        (1, 2, 2, 33.0),
    ] {
        let actual = fixed_snapshot
            .worksheet(Selector::Position(Position::new(sheet)))
            .unwrap()
            .unwrap()
            .number(reference(row, column))
            .unwrap()
            .unwrap()
            .value();
        assert_eq!(actual, value);
    }
    assert_eq!(
        fixed_commit
            .patch()
            .inverse()
            .apply(fixed_snapshot)
            .unwrap()
            .bytes(),
        source.bytes()
    );

    // Repeat the same reverse/join/clone order through structural cell
    // removals.  These operations use the structural sidecar and exercise
    // its post-join insertion path independently of fixed edits.
    let mut structural = source.edit();
    structural
        .remove_cell(Selector::Position(Position::new(1)), reference(2, 2))
        .unwrap();
    structural
        .remove_cell(Selector::Position(Position::new(0)), reference(1, 1))
        .unwrap();
    let mut incoming_structural = source.edit();
    incoming_structural
        .remove_cell(Selector::Position(Position::new(1)), reference(0, 0))
        .unwrap();
    incoming_structural
        .remove_cell(Selector::Position(Position::new(0)), reference(0, 0))
        .unwrap();
    structural.join(incoming_structural).unwrap();
    let mut structural_clone = structural.clone();
    structural_clone
        .remove_cell(Selector::Position(Position::new(1)), reference(1, 1))
        .unwrap();
    let structural_commit = structural_clone.commit().unwrap();
    for (position, worksheet) in structural_commit.package().worksheets().enumerate() {
        if position == 0 {
            assert_eq!(worksheet.cells().count(), 1);
            assert_eq!(
                worksheet.cells().next().unwrap().reference(),
                reference(2, 2)
            );
        } else {
            assert_eq!(worksheet.cells().count(), 0);
        }
    }
    assert_eq!(
        structural_commit
            .patch()
            .inverse()
            .apply(structural_commit.package())
            .unwrap()
            .bytes(),
        source.bytes()
    );
}
