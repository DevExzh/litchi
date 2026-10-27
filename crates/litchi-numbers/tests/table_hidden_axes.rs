//! Table hidden axes through the shared engine (record 0379).

use std::path::PathBuf;

use litchi_numbers::Package;
use litchi_numbers::table::hidden_axes::transaction::{Error, Refusal};
use litchi_numbers::table::hidden_axes::{AxisIndex, HiddenAxes};

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn package() -> TestResult<Package> {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/numbers/custom-number-native.numbers");
    Ok(Package::open(path)?)
}

fn bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut out = Vec::new();
    package.write_to(&mut out)?;
    Ok(out)
}

#[test]
fn set_read_clear_and_inverse_are_exact() -> TestResult {
    let source = package()?;
    let original = bytes(&source)?;
    assert_eq!(
        source.table_hidden_axes(0usize, 0usize)?,
        HiddenAxes::empty()
    );
    let axes = HiddenAxes::new([AxisIndex::row(1), AxisIndex::column(0)])?;
    let commit = source
        .edit_table_hidden_axes(0usize, 0usize)?
        .set(axes.clone())
        .commit()?;
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.package().table_hidden_axes(0usize, 0usize)?, axes);
    let reopened = Package::from_bytes(&bytes(commit.package())?)?;
    assert_eq!(reopened.table_hidden_axes(0usize, 0usize)?, axes);

    let restored = commit
        .package()
        .apply_table_hidden_axes(&commit.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, original);

    let cleared = commit
        .package()
        .edit_table_hidden_axes(0usize, 0usize)?
        .clear()
        .commit()?;
    assert_eq!(
        cleared.package().table_hidden_axes(0usize, 0usize)?,
        HiddenAxes::empty()
    );
    Ok(())
}

#[test]
fn replay_noop_and_refusals_are_typed() -> TestResult {
    let source = package()?;
    let axes = HiddenAxes::new([AxisIndex::row(0)])?;
    let commit = source
        .edit_table_hidden_axes(0usize, 0usize)?
        .set(axes)
        .commit()?;
    let replayed = source.apply_table_hidden_axes(commit.patch())?;
    assert_eq!(bytes(replayed.package())?, bytes(commit.package())?);
    assert_eq!(
        commit
            .package()
            .apply_table_hidden_axes(commit.patch())
            .err(),
        Some(Error::PatchConflict)
    );
    let noop = source
        .edit_table_hidden_axes(0usize, 0usize)?
        .clear()
        .commit()?;
    assert!(noop.patch().is_noop());
    assert!(!noop.diagnostics().changed());
    assert_eq!(
        source.table_hidden_axes(0usize, 9usize).err(),
        Some(Error::Refused(Refusal::TableNotFound))
    );
    Ok(())
}
