//! Synthetic and native coverage for the focused Numbers chart-data writer.

use std::path::PathBuf;

use litchi_iwa_archive::{
    Limits,
    package::{Catalog, EntryInsertion},
};
use litchi_numbers::{ChartData, ChartDataError, ChartSelector, Package, SheetSelector};

#[path = "support/chart_arrangement_fixture.rs"]
mod fixture;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn native_fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/numbers/chart-data-native.numbers")
}

fn retirement_fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/numbers/chart-arrangement-retirement-resaved.numbers")
}

fn edited_native_fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/numbers/chart-data-edited-native.numbers")
}

fn synthetic_source() -> TestResult<Vec<u8>> {
    fixture::with_single_rooted_sheet(&fixture::fixture()?)
}

fn synthetic_source_with_previews() -> TestResult<Vec<u8>> {
    let source = synthetic_source()?;
    let catalog = Catalog::from_bytes(&source)?;
    Ok(catalog.reassemble_with_insertions_to_bytes(
        &[
            EntryInsertion::new("preview.jpg", b"full preview"),
            EntryInsertion::new("preview-micro.jpg", b"micro preview"),
            EntryInsertion::new("preview-web.jpg", b"web preview"),
        ],
        Limits::default(),
    )?)
}

fn replacement(before: &ChartData) -> Result<ChartData, litchi_iwa_common::chart::data::DataError> {
    let mut values = before.values().to_owned();
    values[0][0] = Some(-0.0);
    values[1][1] = None;
    ChartData::new(
        before.row_names().to_owned(),
        before.column_names().to_owned(),
        values,
    )
}

#[test]
fn standalone_chart_data_write_reopens_and_inverts_exactly() -> TestResult {
    let source = synthetic_source()?;
    let package = Package::from_bytes(&source)?;
    let before = package.sheet_chart_data(SheetSelector::index(0), ChartSelector::index(0))?;
    let target = replacement(&before)?;

    let changed = package
        .edit_sheet_chart_data(SheetSelector::index(0), ChartSelector::index(0))?
        .set(target.clone())
        .commit()?;
    assert!(changed.diagnostics().changed());
    assert_eq!(changed.patch().before(), &before);
    assert!(changed.patch().after().bitwise_eq(&target));

    let reopened = Package::from_bytes(&exact_bytes(changed.package())?)?;
    let observed = reopened.sheet_chart_data(SheetSelector::index(0), ChartSelector::index(0))?;
    assert!(observed.bitwise_eq(&target));

    let restored = changed
        .package()
        .apply_sheet_chart_data(&changed.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    let restored_data = restored
        .package()
        .sheet_chart_data(SheetSelector::index(0), ChartSelector::index(0))?;
    assert!(restored_data.bitwise_eq(&before));
    Ok(())
}

#[test]
fn chart_data_patch_rejects_stale_target_before_publication() -> TestResult {
    let source = synthetic_source()?;
    let package = Package::from_bytes(&source)?;
    let before = package.sheet_chart_data(SheetSelector::index(0), ChartSelector::index(0))?;
    let changed = package
        .edit_sheet_chart_data(SheetSelector::index(0), ChartSelector::index(0))?
        .set(replacement(&before)?)
        .commit()?;
    let target_bytes = exact_bytes(changed.package())?;
    let error = changed
        .package()
        .apply_sheet_chart_data(changed.patch())
        .expect_err("a forward patch must reject its own target as a source");
    assert!(matches!(error, ChartDataError::PatchConflict));
    assert_eq!(exact_bytes(changed.package())?, target_bytes);
    Ok(())
}

#[test]
fn changed_chart_data_invalidates_previews_but_inverse_restores_exact_source() -> TestResult {
    let source = synthetic_source_with_previews()?;
    let package = Package::from_bytes(&source)?;
    let before = package.sheet_chart_data(SheetSelector::index(0), ChartSelector::index(0))?;
    let target = replacement(&before)?;

    let changed = package
        .edit_sheet_chart_data(SheetSelector::index(0), ChartSelector::index(0))?
        .set(target)
        .commit()?;
    assert_eq!(changed.diagnostics().deleted_previews(), 3);
    let target_catalog = Catalog::from_bytes(&exact_bytes(changed.package())?)?;
    for name in ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"] {
        assert!(target_catalog.iter().all(|entry| entry.name() != name));
    }

    let restored = changed
        .package()
        .apply_sheet_chart_data(&changed.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    Ok(())
}

#[test]
fn chart_data_noop_retains_previews_and_exact_source() -> TestResult {
    let source = synthetic_source_with_previews()?;
    let package = Package::from_bytes(&source)?;
    let before = package.sheet_chart_data(SheetSelector::index(0), ChartSelector::index(0))?;
    let no_op = package
        .edit_sheet_chart_data(SheetSelector::index(0), ChartSelector::index(0))?
        .set(before)
        .commit()?;
    assert!(no_op.patch().is_noop());
    assert_eq!(no_op.diagnostics().deleted_previews(), 0);
    assert_eq!(exact_bytes(no_op.package())?, source);
    Ok(())
}

#[test]
fn linked_native_chart_data_write_is_refused_before_publication() -> TestResult {
    let source = std::fs::read(native_fixture_path())?;
    let package = Package::from_bytes(&source)?;
    let before = package.sheet_chart_data(SheetSelector::index(0), ChartSelector::index(0))?;
    let target = replacement(&before)?;
    let error = package
        .edit_sheet_chart_data(SheetSelector::index(0), ChartSelector::index(0))
        .and_then(|edit| edit.set(target).commit())
        .expect_err("native table-backed chart data must be refused");
    assert!(matches!(error, ChartDataError::UnsupportedDependency));
    assert_eq!(exact_bytes(&package)?, source);
    Ok(())
}

#[test]
fn edited_native_standalone_chart_data_reads_exact_golden() -> TestResult {
    let source = std::fs::read(edited_native_fixture_path())?;
    let package = Package::from_bytes(&source)?;
    let data = package.sheet_chart_data(SheetSelector::index(0), ChartSelector::index(0))?;

    assert_eq!(data.row_names(), ["North", "South"]);
    assert_eq!(data.column_names(), ["April", "May", "June", "July"]);
    assert_eq!(
        data.values(),
        &[
            vec![Some(27.5), Some(12.75), Some(53.0), Some(96.0)],
            vec![Some(55.0), Some(43.0), Some(70.0), Some(58.0)],
        ]
    );
    assert_eq!(exact_bytes(&package)?, source);
    Ok(())
}

#[test]
fn edited_native_standalone_chart_data_accepts_focused_write() -> TestResult {
    let source = std::fs::read(edited_native_fixture_path())?;
    let package = Package::from_bytes(&source)?;
    let before = package.sheet_chart_data(SheetSelector::index(0), ChartSelector::index(0))?;
    let mut values = before.values().to_owned();
    values[0][0] = Some(28.5);
    let target = ChartData::new(
        before.row_names().to_owned(),
        before.column_names().to_owned(),
        values,
    )?;

    let changed = package
        .edit_sheet_chart_data(SheetSelector::index(0), ChartSelector::index(0))?
        .set(target.clone())
        .commit()?;
    assert!(changed.patch().after().bitwise_eq(&target));
    let reopened = Package::from_bytes(&exact_bytes(changed.package())?)?;
    assert!(
        reopened
            .sheet_chart_data(SheetSelector::index(0), ChartSelector::index(0))?
            .bitwise_eq(&target)
    );
    let restored = changed
        .package()
        .apply_sheet_chart_data(&changed.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    Ok(())
}

#[test]
fn native_retirement_chart_data_is_table_backed_and_refused() -> TestResult {
    let source = std::fs::read(retirement_fixture_path())?;
    let package = Package::from_bytes(&source)?;
    let error = package
        .edit_sheet_chart_data(SheetSelector::index(0), ChartSelector::index(0))
        .expect_err("the resaved retirement chart is table-backed");
    assert!(matches!(error, ChartDataError::UnsupportedDependency));
    assert_eq!(exact_bytes(&package)?, source);
    Ok(())
}

#[test]
fn chart_data_write_preserves_source_axes() -> TestResult {
    let source = synthetic_source()?;
    let package = Package::from_bytes(&source)?;
    let before = package.sheet_chart_data(SheetSelector::index(0), ChartSelector::index(0))?;
    let target = ChartData::new(
        vec![String::from("different")],
        before.column_names().to_owned(),
        vec![vec![Some(1.0), Some(2.0)]],
    )?;
    let error = package
        .edit_sheet_chart_data(SheetSelector::index(0), ChartSelector::index(0))?
        .set(target)
        .commit()
        .expect_err("axis replacement must be refused");
    assert!(matches!(error, ChartDataError::ShapeMismatch));
    assert_eq!(exact_bytes(&package)?, source);
    Ok(())
}

#[test]
fn signed_zero_is_not_a_chart_data_noop() -> TestResult {
    let source = synthetic_source()?;
    let package = Package::from_bytes(&source)?;
    let before = package.sheet_chart_data(SheetSelector::index(0), ChartSelector::index(0))?;
    let mut values = before.values().to_owned();
    values[0][0] = Some(-0.0);
    let target = ChartData::new(
        before.row_names().to_owned(),
        before.column_names().to_owned(),
        values,
    )?;
    if before.values()[0][0].is_some_and(|value| value.to_bits() == (-0.0_f64).to_bits()) {
        return Ok(());
    }
    let changed = package
        .edit_sheet_chart_data(SheetSelector::index(0), ChartSelector::index(0))?
        .set(target.clone())
        .commit()?;
    assert!(!changed.patch().is_noop());
    assert!(
        changed
            .package()
            .sheet_chart_data(SheetSelector::index(0), ChartSelector::index(0))?
            .bitwise_eq(&target)
    );
    Ok(())
}

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}
