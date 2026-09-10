//! Native Numbers coverage for the archive-free chart data read API.

use std::path::PathBuf;

use litchi_iwa_archive::Limits;
use litchi_numbers::{
    ChartDataError, ChartSelector, MAX_OBJECTS, Package, PackageReadOptions, PackageSemanticLimits,
    SheetSelector,
};

#[path = "support/chart_arrangement_fixture.rs"]
mod fixture;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/numbers/chart-data-native.numbers")
}

#[test]
fn native_numbers_chart_data_is_owned_and_rectangular() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let package = Package::from_bytes(&source)?;
    let data = package.sheet_chart_data(SheetSelector::index(0), ChartSelector::index(0))?;

    assert_eq!(data.row_names(), ["North", "South"]);
    assert_eq!(data.column_names(), ["April", "May", "June", "July"]);
    assert_eq!(
        data.values(),
        &[
            vec![Some(17.25), None, Some(53.0), Some(96.0)],
            vec![Some(55.0), Some(43.0), Some(70.0), Some(58.0)],
        ]
    );
    assert!(data.values().iter().all(|row| row.len() == 4));
    Ok(())
}

#[test]
fn native_numbers_chart_data_rejects_unmatched_selectors_without_mutation() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let package = Package::from_bytes(&source)?;

    let error = package
        .sheet_chart_data(SheetSelector::name("missing"), ChartSelector::index(0))
        .expect_err("unknown sheet must be rejected");
    assert!(matches!(error, ChartDataError::SheetNameNotFound));

    let error = package
        .sheet_chart_data(SheetSelector::index(0), ChartSelector::index(99))
        .expect_err("unknown chart position must be rejected");
    assert!(matches!(
        error,
        ChartDataError::ChartPositionNotFound { .. }
    ));

    let mut round_trip = Vec::new();
    package.write_to(&mut round_trip)?;
    assert_eq!(round_trip, source);
    Ok(())
}

#[test]
fn native_numbers_chart_data_round_trips_exact_source_bytes() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let package = Package::from_bytes(&source)?;
    let _ = package.sheet_chart_data(SheetSelector::index(0), ChartSelector::index(0))?;

    let mut output = Vec::new();
    package.write_to(&mut output)?;
    assert_eq!(output, source);
    Ok(())
}

#[test]
fn native_numbers_chart_data_errors_do_not_expose_native_ids() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let package = Package::from_bytes(&source)?;
    let error = package
        .sheet_chart_data(SheetSelector::name("missing"), ChartSelector::index(0))
        .expect_err("unknown sheet must be rejected");

    let rendered = format!("{error:?} {error}");
    for identifier in [1_u64, 31, 100, 106] {
        assert!(!rendered.contains(&identifier.to_string()));
    }
    assert!(!rendered.is_empty());
    Ok(())
}

#[test]
fn native_numbers_chart_data_refuses_a_tight_reference_budget() -> TestResult {
    let source = fixture::with_single_rooted_sheet(&fixture::fixture()?)?;
    let semantic = PackageSemanticLimits::new(
        MAX_OBJECTS,
        PackageSemanticLimits::MAX_SHEETS,
        PackageSemanticLimits::MAX_TABLES,
        2,
    )?;
    let package = Package::from_bytes_with_options(
        &source,
        PackageReadOptions::new(Limits::default(), semantic),
    )?;
    let error = package
        .sheet_chart_data(SheetSelector::index(0), ChartSelector::index(0))
        .expect_err("the selected graph must honor the finite reference budget");
    assert!(matches!(error, ChartDataError::LimitExceeded { .. }));

    let mut round_trip = Vec::new();
    package.write_to(&mut round_trip)?;
    assert_eq!(round_trip, source);
    Ok(())
}
