//! Construction/drop proof for deferred Numbers edit entries.
//!
//! Name selectors are caller-owned input.  Each edit admits and copies the
//! bounded selector without walking the package catalog; the first value
//! request, mutation, or commit performs the source-dependent resolution.

use litchi_numbers::shape::image::adjustments::{Error as ImageError, Refusal as ImageRefusal};
use litchi_numbers::sheet::chart::data::{Error as ChartDataError, Refusal as ChartDataRefusal};
use litchi_numbers::table::merge::Region;
use litchi_numbers::table::merge::transaction::{Error as MergeError, Refusal as MergeRefusal};
use litchi_numbers::table::title::{Error as TitleError, Refusal as TitleRefusal, Settings};
use litchi_numbers::{
    ChartData, ChartSelector, ImageAdjustments, ImageSelector, Package, SheetSelector,
    TableSelector,
};

type TestResult = Result<(), Box<dyn std::error::Error>>;

#[test]
fn name_selector_edit_entries_are_deferred_and_droppable() -> TestResult {
    let package = Package::new();

    let chart = package.edit_sheet_chart_data(
        SheetSelector::name("missing sheet"),
        ChartSelector::index(usize::MAX),
    )?;
    assert!(format!("{chart:?}").contains("selection_loaded: false"));
    drop(chart);

    let image = package.edit_sheet_image_adjustments(
        SheetSelector::name("missing sheet"),
        ImageSelector::index(usize::MAX),
    )?;
    assert!(format!("{image:?}").contains("loaded: false"));
    drop(image);

    let title = package.edit_table_title(
        SheetSelector::name("missing sheet"),
        TableSelector::name("missing table"),
    )?;
    assert!(format!("{title:?}").contains("resolved: false"));
    drop(title);

    let merges = package.edit_table_merges(
        SheetSelector::name("missing sheet"),
        TableSelector::name("missing table"),
    )?;
    assert!(format!("{merges:?}").contains("loaded: false"));
    drop(merges);

    let mut chart = package.edit_sheet_chart_data(
        SheetSelector::name("missing sheet"),
        ChartSelector::index(usize::MAX),
    )?;
    assert!(chart.before().is_err());

    let mut image = package.edit_sheet_image_adjustments(
        SheetSelector::name("missing sheet"),
        ImageSelector::index(usize::MAX),
    )?;
    assert!(image.before().is_err());

    let mut title = package.edit_table_title(
        SheetSelector::name("missing sheet"),
        TableSelector::name("missing table"),
    )?;
    assert!(title.path().is_err());

    let mut merges = package.edit_table_merges(
        SheetSelector::name("missing sheet"),
        TableSelector::name("missing table"),
    )?;
    assert!(merges.regions().is_err());
    Ok(())
}

#[test]
fn deferred_failures_preserve_budget_and_error_class_across_retries() -> TestResult {
    let package = Package::new();

    let mut chart = package.edit_sheet_chart_data(
        SheetSelector::name("missing sheet"),
        ChartSelector::index(0),
    )?;
    assert!(matches!(
        chart.before(),
        Err(ChartDataError::Refused(ChartDataRefusal::SheetNameNotFound))
    ));
    assert!(matches!(
        chart.sheet_position(),
        Err(ChartDataError::Refused(ChartDataRefusal::SheetNameNotFound))
    ));
    let replacement = ChartData::new(
        vec![String::from("row")],
        vec![String::from("column")],
        vec![vec![Some(1.0)]],
    )?;
    let chart = chart.set(replacement);
    assert!(matches!(
        chart.commit(),
        Err(ChartDataError::Refused(ChartDataRefusal::SheetNameNotFound))
    ));

    let mut image = package.edit_sheet_image_adjustments(
        SheetSelector::name("missing sheet"),
        ImageSelector::index(0),
    )?;
    assert!(matches!(
        image.before(),
        Err(ImageError::Refused(ImageRefusal::SheetNameNotFound))
    ));
    assert!(matches!(
        image.image_position(),
        Err(ImageError::Refused(ImageRefusal::SheetNameNotFound))
    ));
    let image = image.set(ImageAdjustments::default())?;
    assert!(matches!(
        image.commit(),
        Err(ImageError::Refused(ImageRefusal::SheetNameNotFound))
    ));

    let mut title = package.edit_table_title(
        SheetSelector::name("missing sheet"),
        TableSelector::name("missing table"),
    )?;
    assert!(matches!(
        title.path(),
        Err(TitleError::Refused(TitleRefusal::SheetNotFound))
    ));
    assert!(matches!(
        title.settings(),
        Err(TitleError::Refused(TitleRefusal::SheetNotFound))
    ));
    let title = title.set(Settings::default());
    assert!(matches!(
        title.commit(),
        Err(TitleError::Refused(TitleRefusal::SheetNotFound))
    ));

    let mut merges = package.edit_table_merges(
        SheetSelector::name("missing sheet"),
        TableSelector::name("missing table"),
    )?;
    assert!(matches!(
        merges.regions(),
        Err(MergeError::Refused(MergeRefusal::SheetNotFound))
    ));
    assert!(matches!(
        merges.merge(Region::new(0, 0, 1, 2)?),
        Err(MergeError::Refused(MergeRefusal::SheetNotFound))
    ));
    assert!(matches!(
        merges.commit(),
        Err(MergeError::Refused(MergeRefusal::SheetNotFound))
    ));
    Ok(())
}
