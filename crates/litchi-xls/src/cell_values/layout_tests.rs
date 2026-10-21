#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "layout fixtures intentionally fail fast"
)]

use super::*;
use crate::writer::{Column, Row};

fn fixture() -> Vec<u8> {
    let mut builder = crate::Builder::new();
    let data = builder.add_worksheet("Data").unwrap();
    builder.write_number(data, 0, 0, 1.0).unwrap();
    builder.write_number(data, 2, 3, 2.0).unwrap();
    // A three-column ColInfo-backed width range and one custom row.
    for column in 1..=3 {
        builder
            .set_column_width(data, Column::new(column).unwrap(), 20.0)
            .unwrap();
    }
    builder
        .set_row_height(data, Row::new(2).unwrap(), 30.0)
        .unwrap();
    builder.add_worksheet("Other").unwrap();
    builder.build().unwrap().to_bytes().unwrap()
}

fn reader(snapshot: &Snapshot) -> Reader<Cursor<Vec<u8>>> {
    Reader::new(Cursor::new(snapshot.bytes().to_vec())).unwrap()
}

fn row(snapshot: &Snapshot, index: u16) -> Option<(u16, bool, bool)> {
    let reader = reader(snapshot);
    let worksheet = reader.xls_worksheet(0).unwrap();
    worksheet
        .row_layout(index)
        .map(|row| (row.height_twips(), row.is_hidden(), row.has_custom_height()))
}

fn column(snapshot: &Snapshot, index: u8) -> Option<(u16, bool)> {
    let reader = reader(snapshot);
    let worksheet = reader.xls_worksheet(0).unwrap();
    worksheet
        .column_layout(index)
        .map(|column| (column.width_256ths(), column.is_hidden()))
}

fn cells(snapshot: &Snapshot) -> Vec<String> {
    let reader = reader(snapshot);
    let worksheet = reader.xls_worksheet(0).unwrap();
    worksheet
        .stored_cells()
        .map(|cell| format!("{}:{}={:?}", cell.row(), cell.column(), cell.value()))
        .collect()
}

#[test]
fn row_heights_and_hidden_flags_round_trip() {
    let bytes = fixture();
    let source = Snapshot::from_bytes(bytes.clone(), Limits::default()).unwrap();
    let mut edit = source.edit();
    // A row with cells, the custom row, and an empty row without a record.
    edit.set_row_height(Selector::Position(Position::new(0)), 0, 500)
        .unwrap();
    edit.set_row_hidden(Selector::Position(Position::new(0)), 2, true)
        .unwrap();
    edit.set_row_height(Selector::Name("data"), 9, 720).unwrap();
    edit.set_row_hidden(Selector::Position(Position::new(0)), 9, true)
        .unwrap();
    // Unhiding an absent row is an exact no-op.
    edit.set_row_hidden(Selector::Position(Position::new(0)), 30, false)
        .unwrap();
    let (target, patch) = edit.commit().unwrap().into_parts();
    assert_eq!(row(&target, 0), Some((500, false, true)));
    assert_eq!(row(&target, 2), Some((600, true, true)));
    assert_eq!(row(&target, 9), Some((720, true, true)));
    assert_eq!(row(&target, 30), None);
    assert_eq!(cells(&target), cells(&source));
    assert_eq!(patch.inverse().apply(&target).unwrap().bytes(), &bytes[..]);
    let replayed = patch.semantic().apply(&source).unwrap().into_parts().0;
    assert_eq!(replayed.bytes(), target.bytes());
    let restored = patch
        .semantic()
        .inverse()
        .apply(&target)
        .unwrap()
        .into_parts()
        .0;
    assert_eq!(row(&restored, 0).map(|row| row.2), Some(false));
    assert_eq!(row(&restored, 2), Some((600, false, true)));
    assert_eq!(row(&restored, 9), None);
}

#[test]
fn column_widths_split_ranges_and_create_ranges() {
    let bytes = fixture();
    let source = Snapshot::from_bytes(bytes.clone(), Limits::default()).unwrap();
    let original = column(&source, 2).unwrap();
    let mut edit = source.edit();
    // Inside an existing range, then an uncovered column, then hide one.
    edit.set_column_width(Selector::Position(Position::new(0)), 2, 3_000)
        .unwrap();
    edit.set_column_width(Selector::Position(Position::new(0)), 7, 4_096)
        .unwrap();
    edit.set_column_hidden(Selector::Position(Position::new(0)), 3, true)
        .unwrap();
    edit.set_column_hidden(Selector::Position(Position::new(0)), 9, true)
        .unwrap();
    let (target, patch) = edit.commit().unwrap().into_parts();
    assert_eq!(column(&target, 1), Some(original));
    assert_eq!(column(&target, 2), Some((3_000, false)));
    assert_eq!(column(&target, 3), Some((original.0, true)));
    assert_eq!(column(&target, 7), Some((4_096, false)));
    assert_eq!(column(&target, 9).map(|column| column.1), Some(true));
    assert_eq!(column(&target, 5), None);
    assert_eq!(cells(&target), cells(&source));
    assert_eq!(patch.inverse().apply(&target).unwrap().bytes(), &bytes[..]);
    let replayed = patch.semantic().apply(&source).unwrap().into_parts().0;
    assert_eq!(replayed.bytes(), target.bytes());
    let restored = patch
        .semantic()
        .inverse()
        .apply(&target)
        .unwrap()
        .into_parts()
        .0;
    for index in [1, 2, 3] {
        assert_eq!(column(&restored, index), Some(original), "{index}");
    }
    assert_eq!(column(&restored, 7), None);
}

#[test]
fn layout_composes_with_cells_but_not_axis_shifts() {
    let source = Snapshot::from_bytes(fixture(), Limits::default()).unwrap();
    let mut edit = source.edit();
    edit.insert_cell(
        Selector::Position(Position::new(0)),
        Reference::new(9, 0).unwrap(),
        Value::Number(5.0),
    )
    .unwrap();
    edit.set_row_height(Selector::Position(Position::new(0)), 9, 400)
        .unwrap();
    edit.set_column_width(Selector::Position(Position::new(0)), 0, 2_560)
        .unwrap();
    assert!(
        edit.insert_rows(Selector::Position(Position::new(0)), 4, 1)
            .is_err()
    );
    let target = edit.commit().unwrap().into_parts().0;
    assert_eq!(row(&target, 9), Some((400, false, true)));
    assert_eq!(column(&target, 0), Some((2_560, false)));
    assert!(cells(&target).contains(&"9:0=Float(5.0)".to_string()));

    let mut edit = source.edit();
    edit.insert_rows(Selector::Name("Other"), 4, 1).unwrap();
    assert!(
        edit.set_row_height(Selector::Name("Other"), 0, 400)
            .is_err()
    );
    assert!(
        edit.set_column_hidden(Selector::Name("Other"), 1, true)
            .is_err()
    );
    let mut edit = source.edit();
    assert!(
        edit.set_row_height(Selector::Position(Position::new(0)), 0, 0)
            .is_err()
    );
    assert!(
        edit.set_row_height(Selector::Position(Position::new(0)), 0, 9_000)
            .is_err()
    );
    assert!(
        edit.set_column_width(Selector::Position(Position::new(0)), 0, u16::MAX)
            .is_err()
    );
    assert!(edit.layout_changes.is_empty());
}
