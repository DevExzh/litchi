//! Focused selector and exact-no-op coverage for body-chart text facades.

use litchi_pages::chart::{caption, title};
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
fn body_chart_caption_keeps_unmatched_selector_typed() -> TestResult {
    let package = Package::from_bytes(NATIVE)?;
    let error = package
        .body_chart_caption(BodyChartSelector::index(99))
        .expect_err("unknown chart must be refused by the focused route");
    assert!(matches!(
        error,
        caption::Error::Refused(caption::Refusal::ChartNotFound)
    ));
    Ok(())
}
