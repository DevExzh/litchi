//! Public edit-entry selector-retention limits.

use std::path::PathBuf;

use litchi_keynote::{
    ChartSelector, OpenOptions, Package, SemanticLimits, SlideSelector, TableSelector,
    chart::{
        arrangement::LimitKind as ArrangementLimitKind, data::LimitKind as ChartDataLimitKind,
        title::LimitKind as ChartTitleLimitKind,
    },
    slide::{
        placeholder::Kind,
        table::{
            appearance::transaction::LimitKind as AppearanceLimitKind,
            lock::transaction::LimitKind as LockLimitKind,
        },
        text::{Error as TextError, LimitKind as TextLimitKind, Refusal as TextRefusal},
    },
};

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn package(max_text_bytes: usize) -> TestResult<Package> {
    let source = std::fs::read(
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/iwork/keynote/chart-caption-native.key"),
    )?;
    let semantic = SemanticLimits::builder()
        .max_text_bytes(max_text_bytes)?
        .build()?;
    Ok(Package::from_bytes_with_options(
        &source,
        OpenOptions::new().semantic_limits(semantic),
    )?)
}

#[test]
fn public_edit_entries_admit_exact_cumulative_names_and_zero_positions() -> TestResult {
    let package = package(10)?;

    package.edit_slide_text("1234567890", Kind::Title)?;
    package.edit_slide_table_appearance("1234567890", TableSelector::index(0))?;
    package.edit_slide_table_lock_state("1234567890", TableSelector::index(0))?;
    package.edit_slide_chart_arrangement("12345", "67890")?;
    package.edit_slide_chart_data("12345", "67890")?;
    package.edit_slide_chart_title("12345", "67890")?;

    package.edit_slide_text(SlideSelector::index(0), Kind::Title)?;
    package.edit_slide_table_appearance(SlideSelector::index(0), TableSelector::index(0))?;
    package.edit_slide_table_lock_state(SlideSelector::index(0), TableSelector::index(0))?;
    package.edit_slide_chart_arrangement(SlideSelector::index(0), ChartSelector::index(0))?;
    package.edit_slide_chart_data(SlideSelector::index(0), ChartSelector::index(0))?;
    package.edit_slide_chart_title(SlideSelector::index(0), ChartSelector::index(0))?;
    Ok(())
}

#[test]
fn public_edit_entries_report_one_over_selector_budget_with_typed_errors() -> TestResult {
    let package = package(10)?;

    assert!(matches!(
        package.edit_slide_text("12345678901", Kind::Title),
        Err(litchi_keynote::slide::text::Error::Limit {
            kind: TextLimitKind::SelectorBytes,
            observed: 11,
            maximum: 10,
        })
    ));
    assert!(matches!(
        package.edit_slide_table_appearance("12345678901", TableSelector::index(0)),
        Err(
            litchi_keynote::slide::table::appearance::transaction::Error::Limit {
                kind: AppearanceLimitKind::SelectorBytes,
                observed: 11,
                maximum: 10,
            }
        )
    ));
    assert!(matches!(
        package.edit_slide_table_lock_state("12345678901", TableSelector::index(0)),
        Err(
            litchi_keynote::slide::table::lock::transaction::Error::Limit {
                kind: LockLimitKind::SelectorBytes,
                observed: 11,
                maximum: 10,
            }
        )
    ));

    assert!(matches!(
        package.edit_slide_chart_arrangement("123456", "12345"),
        Err(litchi_keynote::chart::arrangement::Error::Limit {
            kind: ArrangementLimitKind::SelectorBytes,
            observed: 11,
            maximum: 10,
        })
    ));
    assert!(matches!(
        package.edit_slide_chart_data("123456", "12345"),
        Err(litchi_keynote::chart::data::Error::Limit {
            kind: ChartDataLimitKind::SelectorBytes,
            observed: 11,
            maximum: 10,
        })
    ));
    assert!(matches!(
        package.edit_slide_chart_title("123456", "12345"),
        Err(litchi_keynote::chart::title::Error::Limit {
            kind: ChartTitleLimitKind::SelectorBytes,
            observed: 11,
            maximum: 10,
        })
    ));
    Ok(())
}

#[test]
fn public_edit_entry_rejects_a_large_name_before_source_selection() -> TestResult {
    let package = package(1_024)?;
    let name = "x".repeat(64 * 1_024);
    let error = match package.edit_slide_chart_title(name.as_str(), ChartSelector::index(0)) {
        Ok(_) => panic!("the selector limit must run before chart graph selection"),
        Err(error) => error,
    };
    assert!(matches!(
        error,
        litchi_keynote::chart::title::Error::Limit {
            kind: ChartTitleLimitKind::SelectorBytes,
            observed,
            maximum: 1_024,
        } if observed == u64::try_from(name.len())?
    ));

    let refusal = match package.edit_slide_text(name.as_str(), Kind::SlideNumber) {
        Ok(_) => panic!("slide-number refusal must precede selector admission"),
        Err(TextError::Refused(refusal)) => refusal,
        Err(error) => panic!("unexpected slide-number error: {error:?}"),
    };
    assert!(matches!(
        refusal,
        TextRefusal::UnsupportedKind {
            kind: Kind::SlideNumber
        }
    ));
    Ok(())
}
