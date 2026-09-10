//! Native and source-preservation coverage for the focused Pages chart
//! metadata reader.

use litchi_iwa_common::chart::kind::Kind;
use litchi_pages::{BodyChartSelector, Package};

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const NATIVE: &[u8] =
    include_bytes!("../../../test-data/iwork/pages/chart-arrangement-native.pages");

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

#[test]
fn native_chart_metadata_matches_the_saved_chart_and_preserves_source() -> TestResult {
    let package = Package::from_bytes(NATIVE)?;
    let metadata = package.body_chart_metadata(BodyChartSelector::index(0))?;

    assert_eq!(metadata.kind(), Kind::Column2d);
    assert_eq!(metadata.title(), Some("Pages Arrange native chart"));
    assert_eq!(
        metadata.row_names(),
        ["Region 1".to_owned(), "Region 2".to_owned()]
    );
    assert_eq!(
        metadata.column_names(),
        [
            "April".to_owned(),
            "May".to_owned(),
            "June".to_owned(),
            "July".to_owned(),
        ]
    );
    assert_eq!(metadata.series_count(), 2);
    assert!(!metadata.contains_default_data());
    assert_eq!(exact_bytes(&package)?, NATIVE);
    Ok(())
}
