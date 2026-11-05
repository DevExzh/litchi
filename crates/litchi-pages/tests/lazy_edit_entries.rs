//! Constructor/drop coverage for the selector-first Pages edit families.
//!
//! An edit entry retains only its immutable package reference and the public
//! selector/options.  The invalid selector below is intentional: constructors
//! must accept it and defer the rooted graph walk until an accessor or commit.

use litchi_pages::footnote::body::Selector as BodyFootnoteSelector;
use litchi_pages::table::dimension::{Dimension, Size};
use litchi_pages::{
    BodyChartSelector, BodyTableSelector, HeaderFooterSelector, Kind, Package, Template,
};

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn fixture() -> TestResult<Package> {
    let path = std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/pages/body-table-visible.pages");
    Ok(Package::open(path)?)
}

#[test]
fn edit_constructors_are_deferred_and_droppable() -> TestResult {
    let package = fixture()?;
    let table = BodyTableSelector::index(usize::MAX);
    let chart = BodyChartSelector::index(usize::MAX);
    let note = BodyFootnoteSelector::index(usize::MAX);
    let header = HeaderFooterSelector::index(usize::MAX, Template::First, Kind::Header, usize::MAX);

    drop(package.edit_body_table_values(table)?);
    drop(package.edit_body_table_sort_order(table)?);
    drop(package.edit_body_table_title(table)?);
    drop(package.edit_body_table_header_settings(table)?);
    drop(package.edit_body_table_dimension_size(table, Dimension::Row(0))?);
    drop(package.edit_body_table_name(table)?);
    drop(package.edit_body_table_appearance(table)?);
    drop(package.edit_body_table_hidden_axes(table)?);
    drop(package.edit_body_table_merges(table)?);
    drop(package.edit_body_chart_data(chart)?);
    drop(package.edit_body_drawable_order()?);
    drop(package.edit_header_footer_text(header)?);
    drop(package.edit_body_footnote_text(note)?);
    drop(package.edit_body_footnote(note)?);

    let mut title = package.edit_body_table_title(table)?;
    assert!(title.settings().is_err());
    let mut dimension = package.edit_body_table_dimension_size(table, Dimension::Row(0))?;
    assert!(dimension.size().is_err());
    let mut chart_edit = package.edit_body_chart_data(chart)?;
    assert!(chart_edit.position().is_err());
    let mut drawable = package.edit_body_drawable_order()?;
    let _ = drawable.before()?;
    let mut header_edit = package.edit_header_footer_text(header)?;
    assert!(header_edit.text().is_err());
    let mut footnote_text = package.edit_body_footnote_text(note)?;
    assert!(footnote_text.position().is_err());
    let mut footnote = package.edit_body_footnote(note)?;
    assert!(footnote.before().is_err());

    // Keep the staged-value paths exercised as well: a setter remains cheap,
    // and commit is the first operation that must resolve the invalid target.
    let dimension = package
        .edit_body_table_dimension_size(table, Dimension::Row(0))?
        .set(Size::Default);
    assert!(dimension.commit().is_err());
    let axes = package.edit_body_table_hidden_axes(table)?.clear();
    assert!(axes.commit().is_err());
    Ok(())
}
