//! Create a source-built Numbers workbook with cells across a sparse tile boundary.
//!
//! The 257 × 2 table deliberately places values in row 256, the first row of
//! the second native tile.  The two commits retain the host example's one-cell
//! batch followed by a two-cell batch while exercising exact inverse and
//! reopened semantic reads through the focused package API.

use std::error::Error;
use std::path::PathBuf;

use litchi_numbers::cell::Value;
use litchi_numbers::table::cells::{Change, Input, Storage};
use litchi_numbers::{At, Builder, Identity, Package, SheetSelector, TableSelector};

const SHEET: SheetSelector<'static> = SheetSelector::index(0);
const TABLE: TableSelector<'static> = TableSelector::index(0);
const NUMBER_CELL: At = At::new(256, 1);
const FIRST_TILE_CELL: At = At::new(0, 0);
const BOUNDARY_CELL: At = At::new(256, 0);

fn main() -> Result<(), Box<dyn Error>> {
    let output = PathBuf::from(
        std::env::args()
            .nth(1)
            .ok_or("usage: create_numbers_sparse_table <output.numbers>")?,
    );

    // Identity::BLANK is explicit and deterministic; the example never
    // consults ambient randomness while preserving the source-built shape.
    let source = Builder::new()
        .identity(Identity::BLANK)
        .table_name("Sparse inventory")
        .table_dimensions(257, 2)
        .build()?;
    assert_dimensions(&source)?;
    let source_bytes = exact_bytes(&source)?;

    // Batch 1: one value in the second native tile.
    let first = source
        .edit_table_cells(SHEET, TABLE)?
        .extend([Change::set(NUMBER_CELL, Input::number(257.0)?)])?
        .commit()?;
    assert_number(first.package(), NUMBER_CELL, 257.0)?;
    assert_eq!(
        exact_bytes(
            &first
                .package()
                .apply_table_cells(&first.patch().inverse())?
                .into_package(),
        )?,
        source_bytes,
        "first inverse must restore the exact source"
    );
    let first_bytes = exact_bytes(first.package())?;

    // Batch 2: retain the first-tile value and the boundary value together.
    let second = first
        .package()
        .edit_table_cells(SHEET, TABLE)?
        .extend([
            Change::set(FIRST_TILE_CELL, Input::text("First tile")?),
            Change::set(BOUNDARY_CELL, Input::text("Boundary")?),
        ])?
        .commit()?;
    assert_text(second.package(), FIRST_TILE_CELL, "First tile")?;
    assert_text(second.package(), BOUNDARY_CELL, "Boundary")?;
    assert_number(second.package(), NUMBER_CELL, 257.0)?;
    assert_eq!(
        exact_bytes(
            &second
                .package()
                .apply_table_cells(&second.patch().inverse())?
                .into_package(),
        )?,
        first_bytes,
        "second inverse must restore the first committed source"
    );

    second.package().save(&output)?;
    let reopened = Package::open(&output)?;
    assert_dimensions(&reopened)?;
    assert_text(&reopened, FIRST_TILE_CELL, "First tile")?;
    assert_text(&reopened, BOUNDARY_CELL, "Boundary")?;
    assert_number(&reopened, NUMBER_CELL, 257.0)?;
    Ok(())
}

fn assert_dimensions(package: &Package) -> Result<(), Box<dyn Error>> {
    let table = package
        .table(SHEET, TABLE)?
        .ok_or("the source-built table is missing")?;
    assert_eq!(
        (table.dimensions().rows(), table.dimensions().columns()),
        (257, 2),
        "the sparse tile-boundary dimensions changed"
    );
    Ok(())
}

fn assert_number(package: &Package, position: At, expected: f64) -> Result<(), Box<dyn Error>> {
    let state = package.table_cell(SHEET, TABLE, position)?;
    match state.storage() {
        Storage::Stored(Value::Number(value)) if value.get() == expected => Ok(()),
        other => Err(format!("expected number {expected} at {position:?}, got {other:?}").into()),
    }
}

fn assert_text(package: &Package, position: At, expected: &str) -> Result<(), Box<dyn Error>> {
    let state = package.table_cell(SHEET, TABLE, position)?;
    match state.storage() {
        Storage::Stored(Value::Text(value)) if value == expected => Ok(()),
        other => Err(format!("expected text {expected:?} at {position:?}, got {other:?}").into()),
    }
}

fn exact_bytes(package: &Package) -> Result<Vec<u8>, Box<dyn Error>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}
