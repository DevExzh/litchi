//! Cross-owner coverage for the shared sparse table model and a native
//! Numbers workbook created from its CSV projection.

use std::{error::Error, fs, path::PathBuf};

use litchi_iwa_common::table::cell::value::Value as SharedValue;
use litchi_iwa_common::table::coordinate::CellPosition;
use litchi_iwa_common::table::model::{Builder, Dimensions};
use litchi_numbers::cell::Value;
use litchi_numbers::table::cells::Storage;
use litchi_numbers::{Package, Table as NumbersTable};

const SOURCE_CSV: &str = "../../test-data/iwork/numbers/shared-model-source.csv";
const NATIVE_FIXTURE: &str = "../../test-data/iwork/numbers/shared-model-native.numbers";
const SHEET_NAME: &str = "Sheet 1";
const TABLE_NAME: &str = "shared-model";
const EXPECTED_CSV: &str = "Kind,Value,Note\nNumber,42.5,\"Text Café, \"\"北京\"\"\"\nBoolean,true,\"Text line one\nline two\"\n";

fn repository_path(relative: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join(relative)
}

fn shared_table() -> litchi_iwa_common::table::model::Table {
    let mut builder = Builder::new("Shared model", Dimensions::new(3, 3));
    let rows = [
        [
            SharedValue::Text("Kind".to_owned()),
            SharedValue::Text("Value".to_owned()),
            SharedValue::Text("Note".to_owned()),
        ],
        [
            SharedValue::Text("Number".to_owned()),
            SharedValue::number(42.5).expect("finite number"),
            SharedValue::Text("Text Café, \"北京\"".to_owned()),
        ],
        [
            SharedValue::Text("Boolean".to_owned()),
            SharedValue::Boolean(true),
            SharedValue::Text("Text line one\nline two".to_owned()),
        ],
    ];

    for (row, values) in rows.into_iter().enumerate() {
        for (column, value) in values.into_iter().enumerate() {
            builder
                .set(CellPosition::new(row as u32, column as u32), value)
                .expect("shared model coordinate is in bounds");
        }
    }
    builder
        .finish()
        .expect("shared table has unique coordinates")
}

#[test]
fn shared_model_csv_matches_numbers_facade_and_source_file() -> Result<(), Box<dyn Error>> {
    let shared = shared_table();
    let numbers = NumbersTable::from_shared(shared.clone());

    assert_eq!(shared.to_csv(), EXPECTED_CSV);
    assert_eq!(numbers.to_csv(), EXPECTED_CSV);
    assert_eq!(numbers.to_csv(), shared.to_csv());
    assert_eq!(numbers.as_shared(), &shared);
    assert_eq!(
        fs::read_to_string(repository_path(SOURCE_CSV))?,
        EXPECTED_CSV
    );
    Ok(())
}

fn assert_text(package: &Package, address: &str, expected: &str) -> Result<(), Box<dyn Error>> {
    let state = package.table_cell(SHEET_NAME, TABLE_NAME, CellPosition::from_a1(address)?)?;
    match state.storage() {
        Storage::Stored(Value::Text(value)) => assert_eq!(value, expected, "cell {address}"),
        storage => panic!("cell {address} expected text {expected:?}, got {storage:?}"),
    }
    Ok(())
}

fn assert_number(package: &Package, address: &str, expected: f64) -> Result<(), Box<dyn Error>> {
    let state = package.table_cell(SHEET_NAME, TABLE_NAME, CellPosition::from_a1(address)?)?;
    match state.storage() {
        Storage::Stored(Value::Number(value)) => {
            assert_eq!(value.get().to_bits(), expected.to_bits(), "cell {address}");
        },
        storage => panic!("cell {address} expected number {expected}, got {storage:?}"),
    }
    Ok(())
}

fn assert_boolean(package: &Package, address: &str, expected: bool) -> Result<(), Box<dyn Error>> {
    let state = package.table_cell(SHEET_NAME, TABLE_NAME, CellPosition::from_a1(address)?)?;
    match state.storage() {
        Storage::Stored(Value::Boolean(value)) => assert_eq!(*value, expected, "cell {address}"),
        storage => panic!("cell {address} expected Boolean {expected}, got {storage:?}"),
    }
    Ok(())
}

#[test]
fn native_numbers_fixture_reads_all_shared_model_values() -> Result<(), Box<dyn Error>> {
    let package = Package::open(repository_path(NATIVE_FIXTURE))?;
    let table = package
        .document()
        .sheet(0)?
        .ok_or_else(|| std::io::Error::other("native Numbers fixture has no first sheet"))?
        .at(0)?
        .ok_or_else(|| std::io::Error::other("native Numbers fixture has no first table"))?;

    assert_eq!(table.name(), TABLE_NAME);
    assert_eq!(table.dimensions(), Dimensions::new(3, 3));
    assert_eq!(table.cell_count(), 9);
    assert_eq!(table.to_csv(), EXPECTED_CSV);

    assert_text(&package, "A1", "Kind")?;
    assert_text(&package, "B1", "Value")?;
    assert_text(&package, "C1", "Note")?;
    assert_text(&package, "A2", "Number")?;
    assert_number(&package, "B2", 42.5)?;
    assert_text(&package, "C2", "Text Café, \"北京\"")?;
    assert_text(&package, "A3", "Boolean")?;
    assert_boolean(&package, "B3", true)?;
    assert_text(&package, "C3", "Text line one\nline two")?;
    Ok(())
}
