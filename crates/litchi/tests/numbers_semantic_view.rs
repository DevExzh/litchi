#![cfg(feature = "numbers")]

use litchi::numbers::semantic::Table;
use litchi::numbers::semantic::cell::Value;
use litchi::numbers::semantic::table::{CellPosition, Dimensions, View};

fn assert_send_sync_static<T: Send + Sync + 'static>() {}

#[test]
fn semantic_table_view_preserves_missing_and_explicit_empty_cells() {
    assert_send_sync_static::<View<'static>>();

    let mut builder = Table::builder("Values", Dimensions::new(1, 2));
    assert!(builder.set(CellPosition::new(0, 0), Value::Empty).is_ok());
    let table = builder
        .finish()
        .unwrap_or_else(|error| panic!("semantic table should be valid: {error}"));

    assert!(matches!(
        table.view(CellPosition::new(0, 0)),
        View::Stored(value) if value.is_empty()
    ));
    assert!(matches!(table.view(CellPosition::new(0, 1)), View::Missing));
}
