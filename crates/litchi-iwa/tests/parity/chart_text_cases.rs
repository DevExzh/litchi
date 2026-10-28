//! Host-vs-leaf chart title and axis-title parity for W1-E slice 0375b.
//!
//! The host keeps the differential oracle: each case mutates one selector
//! route, reopens both outputs through the leaf, compares the semantic read,
//! and asks the leaf to apply the inverse patch. Caption authoring is
//! intentionally absent from this slice.

use litchi_iwa::numbers;
use litchi_iwa::pages;
use litchi_iwa_common::chart::axis::Axis;
use litchi_numbers::{ChartSelector, Package as NumbersPackage, SheetSelector};
use litchi_pages::{BodyChartSelector, Package as PagesPackage};

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::leaf_bytes;

const NUMBERS: &str = "iwork/numbers/chart-arrangement-native.numbers";
const PAGES: &str = "iwork/pages/chart-arrangement-native.pages";

#[derive(Clone, Copy, Debug)]
enum Property {
    Title,
    Axis(Axis),
}

fn numbers_target(bytes: &[u8]) -> TestResult<(u64, u64)> {
    let editor = numbers::Draft::from_bytes(bytes)?;
    let sheet = editor
        .sheets()?
        .first()
        .ok_or("native Numbers sheet missing")?;
    let chart = editor
        .sheet_charts(sheet.id())?
        .first()
        .ok_or("native Numbers chart missing")?
        .drawable_object_id;
    Ok((sheet.id(), chart))
}

fn pages_target(bytes: &[u8]) -> TestResult<u64> {
    pages::Draft::from_bytes(bytes)?
        .body_charts()?
        .first()
        .map(|chart| chart.drawable_object_id)
        .ok_or_else(|| "native Pages chart missing".into())
}

fn host_numbers(source: &[u8], property: Property) -> TestResult<Vec<u8>> {
    let (sheet, chart) = numbers_target(source)?;
    let mut editor = numbers::Draft::from_bytes(source)?;
    match property {
        Property::Title => editor.set_sheet_chart_title(sheet, chart, "W1-E title")?,
        Property::Axis(axis) => {
            editor.set_sheet_chart_axis_title(sheet, chart, axis, "W1-E axis title")?;
        },
    }
    Ok(editor.to_bytes()?)
}

fn leaf_numbers(source: &[u8], property: Property) -> TestResult<LeafRun> {
    let package = NumbersPackage::from_bytes(source)?;
    let (output, restored) = match property {
        Property::Title => {
            let commit = package
                .edit_sheet_chart_title(SheetSelector::index(0), ChartSelector::index(0))?
                .set("W1-E title")?
                .commit()?;
            let restored = commit
                .package()
                .apply_sheet_chart_title(&commit.patch().inverse())?;
            (
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            )
        },
        Property::Axis(axis) => {
            let commit = package
                .edit_sheet_chart_axis_title(
                    SheetSelector::index(0),
                    ChartSelector::index(0),
                    axis,
                )?
                .set("W1-E axis title")?
                .commit()?;
            let restored = commit
                .package()
                .apply_sheet_chart_axis_title(&commit.patch().inverse())?;
            (
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            )
        },
    };
    Ok(LeafRun::with_inverse(output, restored))
}

fn read_numbers(source: &[u8], property: Property) -> TestResult<Option<String>> {
    let package = NumbersPackage::from_bytes(source)?;
    Ok(match property {
        Property::Title => {
            package.sheet_chart_title(SheetSelector::index(0), ChartSelector::index(0))?
        },
        Property::Axis(axis) => package.sheet_chart_axis_title(
            SheetSelector::index(0),
            ChartSelector::index(0),
            axis,
        )?,
    })
}

fn host_pages(source: &[u8], property: Property) -> TestResult<Vec<u8>> {
    let chart = pages_target(source)?;
    let mut editor = pages::Draft::from_bytes(source)?;
    match property {
        Property::Title => editor.set_body_chart_title(chart, "W1-E title")?,
        Property::Axis(axis) => editor.set_body_chart_axis_title(chart, axis, "W1-E axis title")?,
    }
    Ok(editor.to_bytes()?)
}

fn leaf_pages(source: &[u8], property: Property) -> TestResult<LeafRun> {
    let package = PagesPackage::from_bytes(source)?;
    let (output, restored) = match property {
        Property::Title => {
            let commit = package
                .edit_body_chart_title(BodyChartSelector::index(0))?
                .set("W1-E title")?
                .commit()?;
            let restored = commit
                .package()
                .apply_body_chart_title(&commit.patch().inverse())?;
            (
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            )
        },
        Property::Axis(axis) => {
            let commit = package
                .edit_body_chart_axis_title(BodyChartSelector::index(0), axis)?
                .set("W1-E axis title")?
                .commit()?;
            let restored = commit
                .package()
                .apply_body_chart_axis_title(&commit.patch().inverse())?;
            (
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            )
        },
    };
    Ok(LeafRun::with_inverse(output, restored))
}

fn read_pages(source: &[u8], property: Property) -> TestResult<Option<String>> {
    let package = PagesPackage::from_bytes(source)?;
    Ok(match property {
        Property::Title => package.body_chart_title(BodyChartSelector::index(0))?,
        Property::Axis(axis) => package.body_chart_axis_title(BodyChartSelector::index(0), axis)?,
    })
}

fn host_numbers_remove(source: &[u8], property: Property) -> TestResult<Vec<u8>> {
    let (sheet, chart) = numbers_target(source)?;
    let mut editor = numbers::Draft::from_bytes(source)?;
    match property {
        Property::Title => {
            editor.set_sheet_chart_title(sheet, chart, "W1-E remove")?;
            let _ = editor.remove_sheet_chart_title(sheet, chart)?;
        },
        Property::Axis(axis) => {
            editor.set_sheet_chart_axis_title(sheet, chart, axis, "W1-E remove")?;
            let _ = editor.remove_sheet_chart_axis_title(sheet, chart, axis)?;
        },
    }
    Ok(editor.to_bytes()?)
}

fn leaf_numbers_remove(source: &[u8], property: Property) -> TestResult<LeafRun> {
    let package = NumbersPackage::from_bytes(source)?;
    match property {
        Property::Title => {
            let first = package
                .edit_sheet_chart_title(SheetSelector::index(0), ChartSelector::index(0))?
                .set("W1-E remove")?
                .commit()?;
            let second = first
                .package()
                .edit_sheet_chart_title(SheetSelector::index(0), ChartSelector::index(0))?
                .clear()?
                .commit()?;
            let restored_first = second
                .package()
                .apply_sheet_chart_title(&second.patch().inverse())?;
            let restored = restored_first
                .package()
                .apply_sheet_chart_title(&first.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(second.package()),
                leaf_bytes!(restored.package()),
            ))
        },
        Property::Axis(axis) => {
            let first = package
                .edit_sheet_chart_axis_title(
                    SheetSelector::index(0),
                    ChartSelector::index(0),
                    axis,
                )?
                .set("W1-E remove")?
                .commit()?;
            let second = first
                .package()
                .edit_sheet_chart_axis_title(
                    SheetSelector::index(0),
                    ChartSelector::index(0),
                    axis,
                )?
                .clear()?
                .commit()?;
            let restored_first = second
                .package()
                .apply_sheet_chart_axis_title(&second.patch().inverse())?;
            let restored = restored_first
                .package()
                .apply_sheet_chart_axis_title(&first.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(second.package()),
                leaf_bytes!(restored.package()),
            ))
        },
    }
}

fn host_pages_remove(source: &[u8], property: Property) -> TestResult<Vec<u8>> {
    let chart = pages_target(source)?;
    let mut editor = pages::Draft::from_bytes(source)?;
    match property {
        Property::Title => {
            editor.set_body_chart_title(chart, "W1-E remove")?;
            let _ = editor.remove_body_chart_title(chart)?;
        },
        Property::Axis(axis) => {
            editor.set_body_chart_axis_title(chart, axis, "W1-E remove")?;
            let _ = editor.remove_body_chart_axis_title(chart, axis)?;
        },
    }
    Ok(editor.to_bytes()?)
}

fn leaf_pages_remove(source: &[u8], property: Property) -> TestResult<LeafRun> {
    let package = PagesPackage::from_bytes(source)?;
    match property {
        Property::Title => {
            let first = package
                .edit_body_chart_title(BodyChartSelector::index(0))?
                .set("W1-E remove")?
                .commit()?;
            let second = first
                .package()
                .edit_body_chart_title(BodyChartSelector::index(0))?
                .clear()?
                .commit()?;
            let restored_first = second
                .package()
                .apply_body_chart_title(&second.patch().inverse())?;
            let restored = restored_first
                .package()
                .apply_body_chart_title(&first.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(second.package()),
                leaf_bytes!(restored.package()),
            ))
        },
        Property::Axis(axis) => {
            let first = package
                .edit_body_chart_axis_title(BodyChartSelector::index(0), axis)?
                .set("W1-E remove")?
                .commit()?;
            let second = first
                .package()
                .edit_body_chart_axis_title(BodyChartSelector::index(0), axis)?
                .clear()?
                .commit()?;
            let restored_first = second
                .package()
                .apply_body_chart_axis_title(&second.patch().inverse())?;
            let restored = restored_first
                .package()
                .apply_body_chart_axis_title(&first.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(second.package()),
                leaf_bytes!(restored.package()),
            ))
        },
    }
}

fn run_numbers_reads(input: &Fixture, property: Property) -> TestResult {
    let source = read_numbers(&input.bytes, property)?;
    let host = host_numbers(&input.bytes, property)?;
    let leaf = leaf_numbers(&input.bytes, property)?.output;
    assert_eq!(
        read_numbers(&host, property)?,
        read_numbers(&leaf, property)?
    );
    assert_eq!(source, read_numbers(&input.bytes, property)?);
    Ok(())
}

fn run_pages_reads(input: &Fixture, property: Property) -> TestResult {
    let source = read_pages(&input.bytes, property)?;
    let host = host_pages(&input.bytes, property)?;
    let leaf = leaf_pages(&input.bytes, property)?.output;
    assert_eq!(read_pages(&host, property)?, read_pages(&leaf, property)?);
    assert_eq!(source, read_pages(&input.bytes, property)?);
    Ok(())
}

#[test]
fn numbers_chart_title_and_axis_title_set_parity() -> TestResult {
    let input = fixture(NUMBERS)?;
    for property in [
        Property::Title,
        Property::Axis(Axis::Category),
        Property::Axis(Axis::Value),
    ] {
        Case::new(
            "numbers_chart_text_title_axis_set_parity",
            Format::Numbers,
            Tier::A,
            Evidence::Queued,
            &input,
        )
        .host(|source| host_numbers(source, property))
        .leaf(|source| leaf_numbers(source, property))
        .read("numbers chart text", move |source| {
            read_numbers(source, property)
        })
        .run()?;
    }
    Ok(())
}

#[test]
fn numbers_chart_title_and_axis_title_read_parity() -> TestResult {
    let input = fixture(NUMBERS)?;
    for property in [
        Property::Title,
        Property::Axis(Axis::Category),
        Property::Axis(Axis::Value),
    ] {
        run_numbers_reads(&input, property)?;
    }
    Ok(())
}

#[test]
fn numbers_chart_title_and_axis_title_remove_parity() -> TestResult {
    let input = fixture(NUMBERS)?;
    for property in [
        Property::Title,
        Property::Axis(Axis::Category),
        Property::Axis(Axis::Value),
    ] {
        Case::new(
            "numbers_chart_text_title_axis_remove_parity",
            Format::Numbers,
            Tier::A,
            Evidence::Queued,
            &input,
        )
        .host(|source| host_numbers_remove(source, property))
        .leaf(|source| leaf_numbers_remove(source, property))
        .read("numbers chart text", move |source| {
            read_numbers(source, property)
        })
        .run()?;
    }
    Ok(())
}

#[test]
fn pages_chart_title_and_axis_title_set_parity() -> TestResult {
    let input = fixture(PAGES)?;
    for property in [
        Property::Title,
        Property::Axis(Axis::Category),
        Property::Axis(Axis::Value),
    ] {
        Case::new(
            "pages_chart_text_title_axis_set_parity",
            Format::Pages,
            Tier::A,
            Evidence::Queued,
            &input,
        )
        .host(|source| host_pages(source, property))
        .leaf(|source| leaf_pages(source, property))
        .read("pages chart text", move |source| {
            read_pages(source, property)
        })
        .run()?;
    }
    Ok(())
}

#[test]
fn pages_chart_title_and_axis_title_read_parity() -> TestResult {
    let input = fixture(PAGES)?;
    for property in [
        Property::Title,
        Property::Axis(Axis::Category),
        Property::Axis(Axis::Value),
    ] {
        run_pages_reads(&input, property)?;
    }
    Ok(())
}

#[test]
fn pages_chart_title_and_axis_title_remove_parity() -> TestResult {
    let input = fixture(PAGES)?;
    for property in [
        Property::Title,
        Property::Axis(Axis::Category),
        Property::Axis(Axis::Value),
    ] {
        Case::new(
            "pages_chart_text_title_axis_remove_parity",
            Format::Pages,
            Tier::A,
            Evidence::Queued,
            &input,
        )
        .host(|source| host_pages_remove(source, property))
        .leaf(|source| leaf_pages_remove(source, property))
        .read("pages chart text", move |source| {
            read_pages(source, property)
        })
        .run()?;
    }
    Ok(())
}
