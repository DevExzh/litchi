//! Native scalar and reference formulas remain readable under aggregate budgets.

use std::{error::Error, path::PathBuf};

use litchi_numbers::{Package, cell::Value, table::CellPosition, table::cells::Storage};

#[test]
fn native_scalar_and_reference_formulas_survive_bounded_readback() -> Result<(), Box<dyn Error>> {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/numbers/formula-budget-native.numbers");
    let package = Package::open(path)?;
    for (address, expected) in [("C2", "=SUM(1,2)"), ("C3", "=SUM(B2,C2)")] {
        let state =
            package.table_cell("Sheet 1", "shared-model", CellPosition::from_a1(address)?)?;
        match state.storage() {
            Storage::Stored(Value::Formula(formula)) => assert_eq!(formula, expected),
            other => panic!("expected native formula at {address}, got {other:?}"),
        }
    }
    Ok(())
}
