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

#[test]
fn package_table_selects_semantics_without_leaking_package_objects() -> TestResult {
    let package = Package::open(fixture_path())?;

    let by_name = package
        .table("Sheet 1", "Table 1")?
        .expect("fixture table should exist");
    let by_position = package.table(0, 0)?.expect("fixture table should exist");

    fn assert_semantic_table(_: &litchi_numbers::Table) {}
    assert_semantic_table(by_name);
    assert_eq!(by_name, by_position);
    assert_eq!(by_name.name(), "Table 1");
    assert!(matches!(
        by_name.get_a1("B3")?,
        Some(litchi_numbers::cell::Value::Number(value)) if value.get() == 42.0
    ));
    assert!(package.table("Missing", "Table 1")?.is_none());
    assert!(package.table("Sheet 1", "Missing")?.is_none());
    assert!(package.table(1, 0)?.is_none());
    assert!(package.table(0, 1)?.is_none());

    // Package-derived semantic snapshots intentionally do not claim source
    // acquisition diagnostics; adding the convenience lookup must not alter
    // the existing stats contract.
    assert_eq!(package.document().stats(), None);
    Ok(())
}
