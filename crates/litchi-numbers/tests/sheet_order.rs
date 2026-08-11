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
        Err(Error::NoStagedOperation)
    ));
    let commit = package.edit_sheet_order().move_sheet(0usize, 0)?.commit()?;
    assert!(commit.patch().is_noop());
    assert!(!commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 0);
    assert_eq!(commit.diagnostics().deleted_previews(), 0);
    assert_eq!(bytes(commit.package())?, source);
    assert!(matches!(
        package.edit_sheet_order().move_sheet(0usize, 1),
        Err(Error::DestinationOutOfRange {
            position: 1,
            sheet_count: 1
        })
    ));
    assert!(matches!(
        package.edit_sheet_order().move_sheet("missing", 0),
        Err(Error::SheetNotFound)
    ));
    fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<LimitKind>();
    Ok(())
}
