//! Native Numbers coverage for the archive-free chart metadata read API.

use std::path::PathBuf;

use litchi_iwa_archive::Limits;
use litchi_iwa_common::chart::kind::Kind;
use litchi_numbers::{
    ChartMetadataError, ChartSelector, MAX_OBJECTS, Package, PackageReadOptions,
    PackageSemanticLimits, SheetSelector,
};

#[path = "support/chart_arrangement_fixture.rs"]
mod fixture;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/numbers/chart-arrangement-native.numbers")
}

#[test]
fn native_numbers_chart_metadata_is_semantic_and_owned() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let package = Package::from_bytes(&source)?;
    let metadata =
        package.sheet_chart_metadata(SheetSelector::index(0), ChartSelector::index(0))?;

    assert_eq!(metadata.kind(), Kind::Column2d);
    assert_eq!(metadata.title(), Some("Numbers Arrange native chart"));
    assert_eq!(metadata.row_names(), ["North", "South"]);
    assert_eq!(metadata.column_names(), ["April", "May", "June", "July"]);
    assert_eq!(metadata.series_count(), 2);
    assert!(!metadata.contains_default_data());
    assert_eq!(
        metadata.all_text().collect::<Vec<_>>(),
        [
            "Numbers Arrange native chart",
            "North",
            "South",
            "April",
            "May",
            "June",
            "July",
        ]
    );
    assert!(metadata.has_content());
    Ok(())
}

#[test]
fn native_numbers_chart_metadata_rejects_unmatched_selectors_without_mutation() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let package = Package::from_bytes(&source)?;
    let error = package
        .sheet_chart_metadata(SheetSelector::name("missing"), ChartSelector::index(0))
        .expect_err("unknown sheet must be rejected");
    assert!(matches!(error, ChartMetadataError::SheetNameNotFound));
    let mut round_trip = Vec::new();
    package.write_to(&mut round_trip)?;
    assert_eq!(round_trip, source);

    let error = package
        .sheet_chart_metadata(SheetSelector::index(0), ChartSelector::index(99))
        .expect_err("unknown chart position must be rejected");
    assert!(matches!(
        error,
        ChartMetadataError::ChartPositionNotFound { .. }
    ));
    Ok(())
}

#[test]
fn native_numbers_chart_metadata_error_does_not_expose_native_ids() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let package = Package::from_bytes(&source)?;
    let error = package
        .sheet_chart_metadata(SheetSelector::name("missing"), ChartSelector::index(0))
        .expect_err("unknown sheet must be rejected");
    let rendered = format!("{error:?} {error}");
    for identifier in [1_u64, 31, 100, 106] {
        assert!(!rendered.contains(&identifier.to_string()));
    }
    assert!(!rendered.is_empty());
    Ok(())
}

#[test]
fn chart_metadata_rejects_ambiguous_and_foreign_rooted_graphs() -> TestResult {
    let source = fixture::fixture()?;
    let ambiguous = fixture::with_duplicate_sheet_name(&source)?;
    assert!(
        Package::from_bytes(&ambiguous).is_err(),
        "duplicate visible sheet names must be rejected before metadata selection"
    );

    let foreign = fixture::with_foreign_chart_non_style(&source)?;
    let package = Package::from_bytes(&foreign)?;
    let error = package
        .sheet_chart_metadata(SheetSelector::index(0), ChartSelector::index(0))
        .expect_err("a chart must not read another chart's title owner");
    assert!(matches!(error, ChartMetadataError::ForeignReference));
    Ok(())
}

#[test]
fn chart_metadata_honors_the_shared_reference_budget() -> TestResult {
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
        .sheet_chart_metadata(SheetSelector::index(0), ChartSelector::index(0))
        .expect_err("metadata selection must stay within the finite reference budget");
    assert!(matches!(error, ChartMetadataError::LimitExceeded { .. }));
    Ok(())
}
