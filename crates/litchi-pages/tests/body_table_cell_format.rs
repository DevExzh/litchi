//! Body-table cell display formats through the shared engine (record 0379).

use std::path::PathBuf;

use litchi_iwa_common::table::cell::data_format::Checkbox;
use litchi_iwa_common::table::cell::data_format::number::{
    DecimalPlaces, NegativeStyle, Number, Percentage, ThousandsSeparator,
};
use litchi_pages::Package;
use litchi_pages::table::format::{At, DataFormat, Error, Refusal};

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const B2: At = At::new(1, 1);

fn package() -> TestResult<Package> {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/pages/body-table-cells-native.pages");
    Ok(Package::open(path)?)
}

fn bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut out = Vec::new();
    package.write_to(&mut out)?;
    Ok(out)
}

fn number() -> TestResult<Number> {
    Ok(Number::new(
        DecimalPlaces::fixed(2)?,
        NegativeStyle::Parentheses,
        ThousandsSeparator::Shown,
    ))
}

#[test]
fn set_read_reset_and_inverse_are_exact() -> TestResult {
    let source = package()?;
    let original = bytes(&source)?;
    assert_eq!(source.body_table_cell_format(0, B2)?, DataFormat::Automatic);

    let commit = source
        .edit_body_table_cell_format(0, B2)?
        .set(number()?)
        .commit()?;
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().rewrites(), 1);
    assert_eq!(
        commit.package().body_table_cell_format(0, B2)?,
        DataFormat::Number(number()?)
    );

    let restored = commit
        .package()
        .apply_body_table_cell_format(&commit.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, original);

    let reset = commit
        .package()
        .edit_body_table_cell_format(0, B2)?
        .reset()
        .commit()?;
    assert_eq!(
        reset.package().body_table_cell_format(0, B2)?,
        DataFormat::Automatic
    );
    Ok(())
}

#[test]
fn replay_and_noop_follow_the_exact_source() -> TestResult {
    let source = package()?;
    let commit = source
        .edit_body_table_cell_format(0, B2)?
        .set(number()?)
        .commit()?;
    let replayed = source.apply_body_table_cell_format(commit.patch())?;
    assert_eq!(bytes(replayed.package())?, bytes(commit.package())?);
    // The patch belongs to its source artifact only.
    assert_eq!(
        commit
            .package()
            .apply_body_table_cell_format(commit.patch())
            .err(),
        Some(Error::PatchConflict)
    );
    let noop = source
        .edit_body_table_cell_format(0, B2)?
        .reset()
        .commit()?;
    assert!(!noop.diagnostics().changed());
    assert!(noop.patch().is_noop());
    assert_eq!(bytes(noop.package())?, bytes(&source)?);
    Ok(())
}

#[test]
fn refusals_are_typed() -> TestResult {
    let source = package()?;
    let missing = source.body_table_cell_format(0, At::new(4_000, 0)).err();
    assert_eq!(missing, Some(Error::Refused(Refusal::CellNotFound)));
    let table = source.body_table_cell_format(99, B2).err();
    assert_eq!(table, Some(Error::Refused(Refusal::TableNotFound)));
    // Scalar controls and Automatic to Percentage on the native cell's
    // implicit Number entry are written since record 0412.
    let control = source
        .edit_body_table_cell_format(0, B2)?
        .set(Checkbox)
        .commit()?;
    assert_eq!(
        control.package().body_table_cell_format(0, B2)?,
        DataFormat::Checkbox(Checkbox)
    );
    let percentage = Percentage::new(
        DecimalPlaces::fixed(1)?,
        NegativeStyle::MinusSign,
        ThousandsSeparator::Hidden,
    );
    let written = source
        .edit_body_table_cell_format(0, B2)?
        .set(percentage)
        .commit()?;
    assert_eq!(
        written.package().body_table_cell_format(0, B2)?,
        DataFormat::Percentage(percentage)
    );
    Ok(())
}
