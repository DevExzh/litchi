//! Construction/drop proof for deferred Numbers edit entries.
//!
//! Name selectors are caller-owned input.  Each edit admits and copies the
//! bounded selector without walking the package catalog; the first value
//! request, mutation, or commit performs the source-dependent resolution.

use litchi_numbers::{ChartSelector, ImageSelector, Package, SheetSelector, TableSelector};

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
