//! Focused selector and exact-no-op coverage for body-chart text facades.

use litchi_iwa_common::chart::axis::Axis;
use litchi_pages::{BodyChartSelector, Package};

const NATIVE: &[u8] =
    include_bytes!("../../../test-data/iwork/pages/chart-arrangement-native.pages");

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

#[test]
fn native_body_chart_title_reads_and_unstaged_edit_is_an_exact_noop() -> TestResult {
    let package = Package::from_bytes(NATIVE)?;
    assert_eq!(
        package.body_chart_title(BodyChartSelector::index(0))?,
        Some("Pages Arrange native chart".to_owned())
    );

    let edit = package.edit_body_chart_title(BodyChartSelector::index(0))?;
    assert_eq!(edit.before(), Some("Pages Arrange native chart"));
    let commit = edit.commit()?;
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert_eq!(exact_bytes(commit.package())?, NATIVE);
    Ok(())
}

#[test]
fn body_chart_title_and_axis_title_mutations_are_reversible() -> TestResult {
    let package = Package::from_bytes(NATIVE)?;
    let title = package
        .edit_body_chart_title(BodyChartSelector::index(0))?
        .set("Pages W1-E title")?
        .commit()?;
    assert_eq!(
        title
            .package()
            .body_chart_title(BodyChartSelector::index(0))?,
        Some("Pages W1-E title".to_owned())
    );
    let restored = title
        .package()
        .apply_body_chart_title(&title.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, NATIVE);

    for axis in [Axis::Category, Axis::Value] {
        let commit = package
            .edit_body_chart_axis_title(BodyChartSelector::index(0), axis)?
            .set("Pages W1-E axis")?
            .commit()?;
        assert_eq!(
            commit
                .package()
                .body_chart_axis_title(BodyChartSelector::index(0), axis)?,
            Some("Pages W1-E axis".to_owned())
        );
        let restored = commit
            .package()
            .apply_body_chart_axis_title(&commit.patch().inverse())?;
        assert_eq!(exact_bytes(restored.package())?, NATIVE);
    }
    Ok(())
}
