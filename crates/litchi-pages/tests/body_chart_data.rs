//! Native and source-preservation coverage for the focused Pages chart-data
//! reader.

use litchi_core::Position;
use litchi_pages::{BodyChartDataError, BodyChartSelector, Package};

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const NATIVE: &[u8] = include_bytes!("../../../test-data/iwork/pages/chart-data-native.pages");
const ORIGINAL_NATIVE: &[u8] =
    include_bytes!("../../../test-data/iwork/pages/chart-arrangement-native.pages");

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

#[test]
fn native_chart_data_matches_the_saved_grid_and_preserves_source() -> TestResult {
    let package = Package::from_bytes(NATIVE)?;
    let data = package.body_chart_data(BodyChartSelector::index(0))?;

    assert_eq!(
        data.row_names(),
        ["Region 1".to_owned(), "Region 2".to_owned()]
    );
    assert_eq!(
        data.column_names(),
        [
            "April".to_owned(),
            "May".to_owned(),
            "June".to_owned(),
            "July".to_owned(),
        ]
    );
    assert_eq!(
        data.values(),
        &[
            vec![Some(17.25), None, Some(53.0), Some(96.0)],
            vec![Some(55.0), Some(43.0), Some(70.0), Some(58.0)],
        ]
    );
    assert_eq!(exact_bytes(&package)?, NATIVE);
    Ok(())
}

#[test]
fn original_native_chart_data_remains_compatible() -> TestResult {
    let package = Package::from_bytes(ORIGINAL_NATIVE)?;
    let data = package.body_chart_data(BodyChartSelector::index(0))?;

    assert_eq!(
        data.values(),
        &[
            vec![Some(17.0), Some(26.0), Some(53.0), Some(96.0)],
            vec![Some(55.0), Some(43.0), Some(70.0), Some(58.0)],
        ]
    );
    assert_eq!(exact_bytes(&package)?, ORIGINAL_NATIVE);
    Ok(())
}

#[test]
fn selector_refuses_an_unmatched_body_chart() -> TestResult {
    let package = Package::from_bytes(NATIVE)?;
    assert!(matches!(
        package.body_chart_data(BodyChartSelector::index(1)),
        Err(BodyChartDataError::ChartNotFound { position })
            if position == Position::new(1)
    ));
    Ok(())
}
