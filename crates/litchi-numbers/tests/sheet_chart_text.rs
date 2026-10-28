//! Focused selector and exact-no-op coverage for sheet-chart text facades.

use litchi_iwa_common::chart::axis::Axis;
use litchi_numbers::sheet::chart::{axis_title, title};
use litchi_numbers::{ChartSelector, OpenOptions, Package, SheetSelector, SourceLimits};

const NATIVE: &[u8] =
    include_bytes!("../../../test-data/iwork/numbers/chart-arrangement-native.numbers");

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

#[test]
fn native_chart_title_reads_and_unstaged_edit_is_an_exact_noop() -> TestResult {
    let package = Package::from_bytes(NATIVE)?;
    assert_eq!(
        package.sheet_chart_title(SheetSelector::index(0), ChartSelector::index(0))?,
        Some("Numbers Arrange native chart".to_owned())
    );

    let edit = package.edit_sheet_chart_title(SheetSelector::index(0), ChartSelector::index(0))?;
    assert_eq!(edit.before(), Some("Numbers Arrange native chart"));
    let commit = edit.commit()?;
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert_eq!(exact_bytes(commit.package())?, NATIVE);
    Ok(())
}

#[test]
fn chart_title_and_axis_title_mutations_are_reversible() -> TestResult {
    let package = Package::from_bytes(NATIVE)?;
    let title = package
        .edit_sheet_chart_title(SheetSelector::index(0), ChartSelector::index(0))?
        .set("Numbers W1-E title")?
        .commit()?;
    assert_eq!(
        title
            .package()
            .sheet_chart_title(SheetSelector::index(0), ChartSelector::index(0))?,
        Some("Numbers W1-E title".to_owned())
    );
    let restored = title
        .package()
        .apply_sheet_chart_title(&title.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, NATIVE);

    for axis in [Axis::Category, Axis::Value] {
        let commit = package
            .edit_sheet_chart_axis_title(SheetSelector::index(0), ChartSelector::index(0), axis)?
            .set("Numbers W1-E axis")?
            .commit()?;
        assert_eq!(
            commit.package().sheet_chart_axis_title(
                SheetSelector::index(0),
                ChartSelector::index(0),
                axis
            )?,
            Some("Numbers W1-E axis".to_owned())
        );
        let restored = commit
            .package()
            .apply_sheet_chart_axis_title(&commit.patch().inverse())?;
        assert_eq!(exact_bytes(restored.package())?, NATIVE);
    }
    Ok(())
}

#[test]
fn chart_text_facades_keep_selector_refusals_typed() -> TestResult {
    let package = Package::from_bytes(NATIVE)?;
    let title_error = package
        .sheet_chart_title(SheetSelector::name("missing"), ChartSelector::index(0))
        .expect_err("unknown sheet must not fall back to the host route");
    assert!(matches!(
        title_error,
        title::Error::Refused(title::Refusal::SheetNotFound)
    ));

    let axis_error = package
        .sheet_chart_axis_title(
            SheetSelector::index(0),
            ChartSelector::index(99),
            Axis::Value,
        )
        .expect_err("unknown chart must be refused by the focused route");
    assert!(matches!(
        axis_error,
        axis_title::Error::Refused(axis_title::Refusal::ChartNotFound)
    ));

    Ok(())
}

fn incompressible_text(length: usize) -> String {
    let mut state = 0x9e37_79b9_u32;
    (0..length)
        .map(|_| {
            state ^= state << 13;
            state ^= state >> 17;
            state ^= state << 5;
            char::from(b'!' + (state % 94) as u8)
        })
        .collect()
}

#[test]
fn chart_title_candidate_keeps_the_original_package_budget() -> TestResult {
    let source_limits = SourceLimits::builder()
        .max_input_bytes(NATIVE.len() as u64)?
        .build()?;
    let package =
        Package::from_bytes_with_options(NATIVE, OpenOptions::new().source_limits(source_limits))?;
    let result = package
        .edit_sheet_chart_title(SheetSelector::index(0), ChartSelector::index(0))?
        .set(incompressible_text(256))?
        .commit();
    assert!(matches!(
        result,
        Err(title::Error::Limit {
            kind: title::LimitKind::PackageBytes,
            ..
        })
    ));
    Ok(())
}
