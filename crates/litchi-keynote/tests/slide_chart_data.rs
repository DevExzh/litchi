//! Native Keynote chart-data reads.

use std::path::PathBuf;

use litchi_keynote::{
    ChartData, ChartSelector, Package, Position, SlideChartDataError, SlideSelector,
};

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const ROWS: [&str; 2] = ["Region 1", "Region 2"];
const COLUMNS: [&str; 4] = ["April", "May", "June", "July"];
const VALUES: [[f64; 4]; 2] = [[17.0, 26.0, 53.0, 96.0], [55.0, 43.0, 70.0, 58.0]];

fn fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/keynote/chart-caption-native.key")
}

fn data_fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/keynote/chart-data-native.key")
}

fn expected(values: [[Option<f64>; 4]; 2]) -> ChartData {
    ChartData::new(
        ROWS.iter().map(|value| (*value).to_owned()).collect(),
        COLUMNS.iter().map(|value| (*value).to_owned()).collect(),
        values
            .into_iter()
            .map(|row| row.into_iter().collect())
            .collect(),
    )
    .expect("native chart data is rectangular")
}

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

#[test]
fn native_chart_data_reads_the_selected_grid_and_preserves_exact_bytes() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let package = Package::from_bytes(&source)?;

    let data = package.slide_chart_data(
        SlideSelector::position(Position::new(0)),
        ChartSelector::index(0),
    )?;

    assert_eq!(data, expected(VALUES.map(|row| row.map(Some))));
    assert_eq!(exact_bytes(&package)?, source);
    Ok(())
}

#[test]
fn native_chart_data_preserves_missing_numeric_cells() -> TestResult {
    let source = std::fs::read(data_fixture_path())?;
    let package = Package::from_bytes(&source)?;

    let data = package.slide_chart_data(SlideSelector::index(0), ChartSelector::index(0))?;

    assert_eq!(
        data,
        expected([
            [Some(17.25), None, Some(53.0), Some(96.0)],
            [Some(55.0), Some(43.0), Some(70.0), Some(58.0)],
        ])
    );
    assert_eq!(exact_bytes(&package)?, source);
    Ok(())
}

#[test]
fn chart_data_rejects_unknown_selectors_before_reading_a_grid() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let package = Package::from_bytes(&source)?;

    assert!(matches!(
        package.slide_chart_data(SlideSelector::index(0), ChartSelector::index(1)),
        Err(SlideChartDataError::ChartPositionNotFound { position })
            if position == Position::new(1)
    ));
    assert!(matches!(
        package.slide_chart_data(SlideSelector::index(1), ChartSelector::index(0)),
        Err(SlideChartDataError::SlidePositionNotFound { position })
            if position == Position::new(1)
    ));
    assert!(matches!(
        package.slide_chart_data(SlideSelector::index(0), ChartSelector::name("")),
        Err(SlideChartDataError::EmptyChartName)
    ));
    assert_eq!(exact_bytes(&package)?, source);
    Ok(())
}
