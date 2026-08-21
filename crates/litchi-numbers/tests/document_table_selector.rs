//! Fixture coverage for the archive-free document table selector.

use std::path::PathBuf;

use litchi_numbers::{Document, SheetSelector, TableSelector, cell::Value};

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data/iwork/numbers/basic.numbers")
}

#[test]
fn document_table_selects_fixture_by_names_and_positions() -> TestResult {
    let document = Document::open(fixture_path())?;

    let by_name = document
        .table(
            SheetSelector::name("Sheet 1"),
            TableSelector::name("Table 1"),
        )?
        .expect("fixture table should exist");
    let by_position = document.table(0, 0)?.expect("fixture table should exist");
    let by_borrowed_selectors = document.table("Sheet 1", "Table 1")?;

    assert_eq!(by_name, by_position);
    assert_eq!(Some(by_name), by_borrowed_selectors);
    assert_eq!(by_name.name(), "Table 1");
    assert!(
        matches!(by_name.get_a1("B2")?, Some(Value::Text(value)) if value == "Litchi native Numbers fixture")
    );
    assert!(document.table("Missing", "Table 1")?.is_none());
    assert!(document.table("Sheet 1", "Missing")?.is_none());
    assert!(document.table(1, 0)?.is_none());
    assert!(document.table(0, 1)?.is_none());
    Ok(())
}
