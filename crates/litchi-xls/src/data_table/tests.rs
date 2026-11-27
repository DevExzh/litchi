use super::*;

fn range() -> DataTableRange {
    DataTableRange::new(2, 8, 3, 5).unwrap()
}

#[test]
fn one_variable_round_trips() {
    let table = DataTable::one_variable(
        range(),
        true,
        DataTableInputCell::Present { row: 0, col: 6 },
    );
    let parsed = DataTable::parse(&table.to_payload()).unwrap();
    assert_eq!(parsed, table);
    assert!(!parsed.is_two_variable());
    assert!(parsed.row_orientation());
    let DataTableKind::OneVariable { input, .. } = parsed.kind() else {
        panic!()
    };
    assert_eq!(*input, DataTableInputCell::Present { row: 0, col: 6 });
}

#[test]
fn two_variable_with_deleted_input_round_trips() {
    let mut table = DataTable::two_variable(
        range(),
        DataTableInputCell::Present { row: 1, col: 2 },
        DataTableInputCell::Deleted,
    );
    table.set_always_calc(true);
    let parsed = DataTable::parse(&table.to_payload()).unwrap();
    assert_eq!(parsed, table);
    assert!(parsed.is_two_variable());
    assert!(parsed.always_calc());
}

#[test]
fn one_variable_preserves_undefined_tail() {
    let mut payload = DataTable::one_variable(
        range(),
        false,
        DataTableInputCell::Present { row: 1, col: 2 },
    )
    .to_payload();
    // Scribble the undefined rwInpCol/colInpCol pair and fDeleted2.
    payload[6] |= 0x20;
    payload[12..14].copy_from_slice(&7u16.to_le_bytes());
    payload[14..16].copy_from_slice(&9u16.to_le_bytes());
    let parsed = DataTable::parse(&payload).unwrap();
    assert_eq!(parsed.to_payload(), payload);
}

#[test]
fn ptg_tbl_tokens_name_the_range_origin() {
    let table = DataTable::one_variable(range(), false, DataTableInputCell::Deleted);
    assert_eq!(table.ptg_tbl_tokens(), [codec::PTG_TBL, 2, 0, 3, 0]);
}

#[test]
fn rejects_malformed_records() {
    assert!(DataTable::parse(&[0; 15]).is_err());
    assert!(DataTable::parse(&[0; 17]).is_err());
    // Zero-based origin.
    let mut payload = DataTable::one_variable(
        range(),
        false,
        DataTableInputCell::Present { row: 1, col: 2 },
    )
    .to_payload();
    payload[0..2].copy_from_slice(&0u16.to_le_bytes());
    assert!(DataTable::parse(&payload).is_err());
    // Deleted input without the -1 coordinates.
    let mut payload = DataTable::one_variable(
        range(),
        false,
        DataTableInputCell::Present { row: 1, col: 2 },
    )
    .to_payload();
    payload[6] |= 0x10;
    assert!(DataTable::parse(&payload).is_err());
    // A present input column outside the BIFF8 cell grid.
    let mut payload = DataTable::one_variable(
        range(),
        false,
        DataTableInputCell::Present { row: 1, col: 2 },
    )
    .to_payload();
    payload[10..12].copy_from_slice(&256u16.to_le_bytes());
    assert!(DataTable::parse(&payload).is_err());
    // Reversed or zero-based range.
    assert!(DataTableRange::new(5, 2, 3, 3).is_err());
    assert!(DataTableRange::new(0, 2, 3, 3).is_err());
    assert!(DataTableRange::new(1, 2, 0, 3).is_err());
}
