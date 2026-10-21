//! Native exact-source smoke coverage for sheet-order transactions.

use litchi_numbers::{
    Package,
    sheet::order::{Error, LimitKind},
};

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn fixture() -> std::path::PathBuf {
    std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/numbers/basic.numbers")
}

fn bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

#[test]
fn native_single_sheet_noop_and_typed_staging_errors_are_exact() -> TestResult<()> {
    let package = Package::open(fixture())?;
    let source = bytes(&package)?;
    assert!(matches!(
        package.edit_sheet_order().commit(),
        Err(Error::Refused(
            litchi_numbers::sheet::order::Refusal::NoStagedOperation
        ))
    ));
    let commit = package.edit_sheet_order().move_sheet(0usize, 0)?.commit()?;
    assert!(commit.patch().is_noop());
    assert!(!commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 0);
    assert_eq!(commit.diagnostics().deleted_previews(), 0);
    assert_eq!(bytes(commit.package())?, source);
    assert!(matches!(
        package.edit_sheet_order().move_sheet(0usize, 1),
        Err(Error::Refused(
            litchi_numbers::sheet::order::Refusal::DestinationOutOfRange {
                position: 1,
                sheet_count: 1
            }
        ))
    ));
    assert!(matches!(
        package.edit_sheet_order().move_sheet("missing", 0),
        Err(Error::Refused(
            litchi_numbers::sheet::order::Refusal::SheetNotFound
        ))
    ));
    fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<LimitKind>();
    Ok(())
}

/// Record 0297: a package without previews (here, stripped by an earlier
/// changed move) still reorders, and the inverse restores it exactly.
#[test]
fn a_package_without_previews_still_reorders() -> Result<(), Box<dyn std::error::Error>> {
    let source = std::fs::read(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/iwork/numbers/table-relocation-native.numbers"
    ))?;
    let package = Package::from_bytes(&source)?;
    let first = package
        .edit_sheet_order()
        .move_sheet(litchi_numbers::SheetSelector::index(1), 0)?
        .commit()?;
    let stripped = first.package();
    assert_eq!(stripped.sheets()[0].name(), "Sheet 2");
    let second = stripped
        .edit_sheet_order()
        .move_sheet(litchi_numbers::SheetSelector::index(1), 0)?
        .commit()?;
    assert_eq!(second.package().sheets()[0].name(), "Sheet 1");
    assert_eq!(second.diagnostics().deleted_previews(), 0);
    let restored = second
        .package()
        .apply_sheet_order(&second.patch().inverse())?;
    let mut expected = Vec::new();
    stripped.write_to(&mut expected)?;
    let mut actual = Vec::new();
    restored.package().write_to(&mut actual)?;
    assert_eq!(actual, expected);
    Ok(())
}
