//! Fixture coverage for the archive-free package sheet selector.

use std::path::PathBuf;

use litchi_numbers::{Package, SheetSelector, TableSelector};

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data/iwork/numbers/basic.numbers")
}

#[test]
fn package_sheet_selects_semantics_by_name_and_position() -> TestResult {
    let package = Package::open(fixture_path())?;

    let by_name = package
        .sheet(SheetSelector::name("Sheet 1"))?
        .expect("fixture sheet should exist");
    let by_position = package
        .sheet(SheetSelector::index(0))?
        .expect("fixture sheet should exist at position zero");

    assert_eq!(by_name, by_position);
    assert_eq!(by_name.name(), "Sheet 1");
    let table = by_name
        .select(TableSelector::name("Table 1"))?
        .expect("fixture table should exist");
    assert_eq!(table.name(), "Table 1");

    assert!(package.sheet(SheetSelector::name("Missing"))?.is_none());
    assert!(package.sheet(SheetSelector::index(1))?.is_none());
    Ok(())
}
