#![cfg(feature = "numbers")]

use std::io;
use std::path::PathBuf;

use litchi::numbers::names::{Commit, Diagnostics, Edit, Error, LimitKind, Patch};
use litchi::numbers::{Package, SheetSelector, TableSelector};

fn assert_send_sync<T: Send + Sync>() {}

fn fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data/iwork/numbers/basic.numbers")
}

fn package_bytes(package: &Package) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn assert_names(
    package: &Package,
    sheet_name: &str,
    table_name: &str,
) -> Result<(), Box<dyn std::error::Error>> {
    let sheet = package
        .document()
        .sheet(SheetSelector::name(sheet_name))?
        .ok_or_else(|| io::Error::other("Numbers facade has no expected sheet"))?;
    let table = sheet
        .select(TableSelector::name(table_name))?
        .ok_or_else(|| io::Error::other("Numbers facade has no expected table"))?;
    assert_eq!(sheet.name(), sheet_name);
    assert_eq!(table.name(), table_name);
    Ok(())
}

#[test]
fn atomic_names_transaction_reaches_numbers_facade() -> Result<(), Box<dyn std::error::Error>> {
    assert_send_sync::<Edit<'static>>();
    assert_send_sync::<Patch>();
    assert_send_sync::<Commit>();
    assert_send_sync::<Diagnostics>();
    assert_send_sync::<Error>();
    assert_send_sync::<LimitKind>();

    let package = Package::open(fixture_path())?;
    let source = package_bytes(&package)?;
    assert_names(&package, "Sheet 1", "Table 1")?;

    let sheet_name = "Líneas 你好 🧪";
    let table_name = "表 Café №42";
    let changed = package
        .edit_names()
        .rename_sheet(SheetSelector::index(0), sheet_name)?
        .rename_table(SheetSelector::index(0), TableSelector::index(0), table_name)?
        .commit()?;

    assert_eq!(changed.patch().operation_count(), 2);
    assert!(!changed.patch().is_noop());
    assert!(changed.diagnostics().changed());
    assert_eq!(changed.diagnostics().operations(), 2);
    assert!(changed.diagnostics().touched_components() >= 1);
    assert_eq!(changed.diagnostics().deleted_previews(), 3);
    assert!(changed.diagnostics().full_reparse_performed());
    assert_names(changed.package(), sheet_name, table_name)?;
    assert_names(&package, "Sheet 1", "Table 1")?;
    assert_ne!(package_bytes(changed.package())?, source);
    assert_eq!(package_bytes(&package)?, source);

    let applied = package.apply_names(changed.patch())?;
    assert!(applied.diagnostics().changed());
    assert_eq!(applied.diagnostics().deleted_previews(), 3);
    assert_names(applied.package(), sheet_name, table_name)?;
    assert_eq!(
        package_bytes(applied.package())?,
        package_bytes(changed.package())?
    );

    let restored = changed.package().apply_names(&changed.patch().inverse())?;
    assert!(restored.diagnostics().changed());
    assert_eq!(restored.diagnostics().deleted_previews(), 0);
    assert_eq!(package_bytes(restored.package())?, source);
    assert_names(restored.package(), "Sheet 1", "Table 1")?;
    Ok(())
}
