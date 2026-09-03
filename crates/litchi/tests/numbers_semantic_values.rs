#![cfg(feature = "numbers")]

use litchi::numbers::semantic::cell::{FiniteF64, Type, Value};
use litchi::numbers::semantic::formula::{
    AxisReference, BinaryOperator, CachedValue, CellReference, FormulaError,
};
use litchi::numbers::semantic::table::cells::{State, Storage};
use litchi::numbers::semantic::table::{CellPosition, CellRange, Dimensions};
use litchi::numbers::{Package, SheetSelector, TableSelector};

const FIXTURE: &str = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/iwork/numbers/basic.numbers"
);

fn assert_send_sync_static<T: Send + Sync + 'static>() {}

#[test]
fn semantic_numbers_namespace_exposes_typed_values_without_package_types() {
    assert_send_sync_static::<FiniteF64>();
    assert_send_sync_static::<Type>();
    assert_send_sync_static::<Value>();
    assert_send_sync_static::<AxisReference>();
    assert_send_sync_static::<BinaryOperator>();
    assert_send_sync_static::<CachedValue>();
    assert_send_sync_static::<CellReference>();
    assert_send_sync_static::<FormulaError>();
    assert_send_sync_static::<CellPosition>();
    assert_send_sync_static::<CellRange>();
    assert_send_sync_static::<Dimensions>();
    assert_send_sync_static::<State>();
    assert_send_sync_static::<Storage>();

    let number = Value::number(42.0).expect("finite semantic values are accepted");
    assert_eq!(number.cell_type(), Type::Number);
    assert_eq!(number.as_number(), Some(42.0));
    assert_eq!(CellPosition::from_a1("B2").unwrap().to_string(), "B2");
    assert_eq!(CellRange::from_a1("B2:C3").unwrap().area(), Some(4));
    assert_eq!(Dimensions::new(4, 5).area(), Some(20));
    assert_eq!(CachedValue::boolean(true), CachedValue::boolean(true));
}

#[test]
fn semantic_numbers_cell_state_preserves_typed_presence() -> Result<(), Box<dyn std::error::Error>>
{
    let package = Package::open(FIXTURE)?;
    let position = CellPosition::from_a1("B2")?;
    let state = package.table_cell(
        SheetSelector::name("Sheet 1"),
        TableSelector::name("Table 1"),
        position,
    )?;

    assert_eq!(state.position(), position);
    let value = state
        .storage()
        .value()
        .ok_or_else(|| std::io::Error::other("fixture cell should be materialized"))?;
    assert_eq!(value.cell_type(), Type::Text);
    assert!(matches!(value, Value::Text(_)));
    Ok(())
}
