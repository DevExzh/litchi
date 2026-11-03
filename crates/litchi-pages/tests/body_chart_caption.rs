//! Selector, reversibility and refusal coverage for body-chart captions
//! (record 0427).

use litchi_pages::chart::caption::{Error, Refusal};
use litchi_pages::{BodyChartSelector, Package};

const NATIVE: &[u8] = include_bytes!("../../../test-data/iwork/pages/chart-data-native.pages");

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

#[test]
fn body_chart_caption_set_replace_clear_and_inverse() -> TestResult {
    let chart = BodyChartSelector::index(0);
    let package = Package::from_bytes(NATIVE)?;
    assert_eq!(package.body_chart_caption(chart)?, None);

    let mut edit = package.edit_body_chart_caption(chart)?;
    assert_eq!(edit.before(), None);
    edit.set("Revenue — 2026")?;
    let created = edit.commit()?;
    assert!(created.diagnostics().changed());
    assert_eq!(created.diagnostics().created_objects(), 4);
    assert_eq!(
        created.package().body_chart_caption(chart)?.as_deref(),
        Some("Revenue — 2026")
    );
    let restored = created
        .package()
        .apply_body_chart_caption(&created.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, NATIVE);

    let mut edit = created.package().edit_body_chart_caption(chart)?;
    edit.set("Updated")?;
    let replaced = edit.commit()?;
    assert_eq!(replaced.diagnostics().created_objects(), 0);
    assert_eq!(
        replaced.package().body_chart_caption(chart)?.as_deref(),
        Some("Updated")
    );

    let mut edit = replaced.package().edit_body_chart_caption(chart)?;
    edit.clear()?;
    let cleared = edit.commit()?;
    assert_eq!(cleared.diagnostics().created_objects(), 1);
    assert_eq!(cleared.package().body_chart_caption(chart)?, None);
    Ok(())
}

#[test]
fn body_chart_caption_refusals_are_typed() -> TestResult {
    let package = Package::from_bytes(NATIVE)?;
    match package.body_chart_caption(BodyChartSelector::index(9)) {
        Err(Error::Refused(Refusal::NotFound { kind: "chart", .. })) => {},
        other => return Err(format!("expected NotFound, got {other:?}").into()),
    }
    let mut edit = package.edit_body_chart_caption(0)?;
    match edit.set("a\u{fffc}b") {
        Err(Error::Refused(Refusal::ReservedMarker)) => {},
        other => return Err(format!("expected ReservedMarker, got {other:?}").into()),
    }
    edit.set("ok")?;
    match edit.clear() {
        Err(Error::Refused(Refusal::OperationAlreadyStaged)) => {},
        other => return Err(format!("expected OperationAlreadyStaged, got {other:?}").into()),
    }
    let unstaged = package.edit_body_chart_caption(0)?.commit()?;
    assert!(unstaged.patch().is_noop());
    assert_eq!(exact_bytes(unstaged.package())?, NATIVE);
    Ok(())
}
