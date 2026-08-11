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

#[test]
fn table_header_transaction_reaches_numbers_facade() -> Result<(), Box<dyn std::error::Error>> {
    use litchi::numbers::table::headers::Settings;
    use litchi::numbers::table::headers::transaction::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Path,
    };

    assert_send_sync::<Settings>();
    assert_send_sync::<Edit<'static>>();
    assert_send_sync::<Patch>();
    assert_send_sync::<Commit>();
    assert_send_sync::<Diagnostics>();
    assert_send_sync::<Error>();
    assert_send_sync::<LimitKind>();
    assert_send_sync::<Path>();

    let package = Package::open(fixture_path())?;
    let source = package_bytes(&package)?;
    let sheet = SheetSelector::index(0);
    let table = TableSelector::index(0);
    let before = package.table_header_settings(sheet, table)?;

    let edit = package.edit_table_headers(sheet, table)?;
    assert_eq!(edit.path(), Path::Table { sheet: 0, table: 0 });
    assert_eq!(edit.settings(), before);
    let noop = edit.set(before).commit()?;
    assert_eq!(noop.patch().path(), Path::Table { sheet: 0, table: 0 });
    assert_eq!(noop.patch().before(), before);
    assert_eq!(noop.patch().after(), before);
    assert!(noop.patch().is_noop());
    assert!(!noop.diagnostics().changed());
    assert_eq!(noop.diagnostics().touched_components(), 0);
    assert_eq!(noop.diagnostics().deleted_previews(), 0);
    assert!(!noop.diagnostics().full_reparse_performed());
    assert_eq!(package_bytes(noop.package())?, source);

    let noop_applied = package.apply_table_headers(noop.patch())?;
    assert!(!noop_applied.diagnostics().changed());
    assert_eq!(package_bytes(noop_applied.package())?, source);

    let mut after = before;
    after.header_rows_frozen = Some(!before.header_rows_are_frozen());
    let changed = package
        .edit_table_headers(sheet, table)?
        .set(after)
        .commit()?;
    assert_eq!(changed.patch().before(), before);
    assert_eq!(changed.patch().after(), after);
    assert!(!changed.patch().is_noop());
    assert!(changed.diagnostics().changed());
    assert_eq!(changed.diagnostics().touched_components(), 1);
    assert!(changed.diagnostics().full_reparse_performed());
    assert_eq!(
        changed.package().table_header_settings(sheet, table)?,
        after
    );
    assert_eq!(package_bytes(&package)?, source);
    assert_ne!(package_bytes(changed.package())?, source);

    let applied = package.apply_table_headers(changed.patch())?;
    assert!(applied.diagnostics().changed());
    assert_eq!(
        applied.package().table_header_settings(sheet, table)?,
        after
    );
    assert_eq!(
        package_bytes(applied.package())?,
        package_bytes(changed.package())?
    );

    let inverse = changed.patch().inverse();
    let restored = changed.package().apply_table_headers(&inverse)?;
    assert!(restored.diagnostics().changed());
    assert_eq!(
        restored.package().table_header_settings(sheet, table)?,
        before
    );
    assert_eq!(package_bytes(restored.package())?, source);
    Ok(())
}

#[test]
fn sheet_order_transaction_reaches_numbers_facade() -> Result<(), Box<dyn std::error::Error>> {
    use litchi::numbers::sheet::order::{Commit, Diagnostics, Edit, Error, LimitKind, Patch};

    assert_send_sync::<Edit<'static>>();
    assert_send_sync::<Patch>();
    assert_send_sync::<Commit>();
    assert_send_sync::<Diagnostics>();
    assert_send_sync::<Error>();
    assert_send_sync::<LimitKind>();

    let package = Package::open(fixture_path())?;
    let source = package_bytes(&package)?;
    assert_eq!(package.sheets().len(), 1);

    let edit = package.edit_sheet_order();
    let edit_debug = format!("{edit:?}");
    assert!(edit_debug.contains("operation: None"));
    assert!(!edit_debug.contains("Sheet 1"));
    assert!(!edit_debug.contains(".iwa"));

    let edit = edit.move_sheet(SheetSelector::index(0), 0)?;
    let staged_debug = format!("{edit:?}");
    assert!(staged_debug.contains("source_position: 0"));
    assert!(staged_debug.contains("destination_position: 0"));
    assert!(!staged_debug.contains("Sheet 1"));
    assert!(!staged_debug.contains(".iwa"));

    let noop = edit.commit()?;
    assert_eq!(noop.patch().source_position(), 0);
    assert_eq!(noop.patch().destination_position(), 0);
    assert_eq!(
        noop.patch().source_fingerprint(),
        noop.patch().target_fingerprint()
    );
    assert!(noop.patch().is_noop());
    assert!(!noop.diagnostics().changed());
    assert_eq!(noop.diagnostics().touched_components(), 0);
    assert_eq!(noop.diagnostics().deleted_previews(), 0);
    assert!(!noop.diagnostics().full_reparse_performed());
    assert_eq!(package_bytes(noop.package())?, source);

    let patch_debug = format!("{:?}", noop.patch());
    assert!(patch_debug.contains("source_position: 0"));
    assert!(patch_debug.contains("destination_position: 0"));
    assert!(!patch_debug.contains("fingerprint"));
    assert!(!patch_debug.contains("Sheet 1"));
    assert!(!patch_debug.contains(".iwa"));

    let applied = package.apply_sheet_order(noop.patch())?;
    assert!(!applied.diagnostics().changed());
    assert_eq!(package_bytes(applied.package())?, source);

    let inverse = noop.patch().inverse();
    assert!(inverse.is_noop());
    let restored = noop.package().apply_sheet_order(&inverse)?;
    assert!(!restored.diagnostics().changed());
    assert_eq!(package_bytes(restored.package())?, source);

    let renamed = package
        .edit_names()
        .rename_sheet(SheetSelector::index(0), "Conflict sheet")?
        .commit()?;
    let conflict = renamed
        .package()
        .apply_sheet_order(noop.patch())
        .expect_err("a sheet-order patch must authorize its exact source");
    assert!(matches!(&conflict, Error::PatchConflict));
    assert_eq!(format!("{conflict:?}"), "PatchConflict");
    Ok(())
}
