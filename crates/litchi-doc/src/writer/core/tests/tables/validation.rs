use super::super::support::*;

#[test]
fn test_add_table_invalid_dimensions() {
    let mut writer = Writer::new();
    assert!(writer.add_table(0, 3).is_err());
    assert!(writer.add_table(2, 0).is_err());
    assert!(writer.add_table(0, 0).is_err());
    assert!(writer.add_table(1, 64).is_err());
}

#[test]
fn test_set_table_cell_invalid_indices() {
    let mut writer = Writer::new();
    let idx = writer.add_table(2, 2).unwrap();
    assert!(writer.set_table_cell_text(idx, 2, 0, "Invalid").is_err());
    assert!(writer.set_table_cell_text(idx, 0, 2, "Invalid").is_err());
    assert!(writer.set_table_cell_text(999, 0, 0, "Invalid").is_err());
}

#[test]
fn rejects_invalid_table_row_formatting() {
    let mut writer = Writer::new();
    let table = writer.add_table(2, 2).unwrap();
    let one_cell = crate::writer::TableRow {
        cells: vec![crate::writer::TableCell {
            width: 1000,
            merged: false,
            ..crate::writer::TableCell::default()
        }],
        ..crate::writer::TableRow::default()
    };
    assert!(writer.set_table_row_formatting(table, 0, one_cell).is_err());

    let invalid_merge = crate::writer::TableRow {
        cells: vec![
            crate::writer::TableCell {
                width: 1000,
                merged: true,
                ..crate::writer::TableCell::default()
            },
            crate::writer::TableCell {
                width: 1000,
                merged: false,
                ..crate::writer::TableCell::default()
            },
        ],
        ..crate::writer::TableRow::default()
    };
    assert!(
        writer
            .set_table_row_formatting(table, 0, invalid_merge)
            .is_err()
    );

    let late_header = crate::writer::TableRow {
        cells: vec![
            crate::writer::TableCell {
                width: 1000,
                merged: false,
                ..crate::writer::TableCell::default()
            },
            crate::writer::TableCell {
                width: 1000,
                merged: false,
                ..crate::writer::TableCell::default()
            },
        ],
        is_header: true,
        ..crate::writer::TableRow::default()
    };
    writer
        .set_table_row_formatting(table, 1, late_header)
        .unwrap();
    assert!(writer.write_to(&mut Cursor::new(Vec::new())).is_err());

    let mut writer = Writer::new();
    let table = writer.add_table(1, 1).unwrap();
    writer
        .set_table_row_formatting(
            table,
            0,
            crate::writer::TableRow {
                cells: vec![crate::writer::TableCell {
                    width: 1000,
                    vertical_merge: crate::parts::tap::VerticalMergeStatus::Merged,
                    ..crate::writer::TableCell::default()
                }],
                ..crate::writer::TableRow::default()
            },
        )
        .unwrap();
    assert!(writer.write_to(&mut Cursor::new(Vec::new())).is_err());
}
