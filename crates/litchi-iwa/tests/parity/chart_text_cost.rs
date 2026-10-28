//! Callgrind ABBA probes for the W1-E chart title/axis-title slice.
//!
//! Each ignored test is one ledger route: title or axis family, and read,
//! set or clear, for one format.  The host and leaf calls are deliberately
//! kept in named, non-inlined functions so a coordinator can collect one
//! full-facade route at a time with callgrind's `--toggle-collect` filter.
//! Read routes include open, selector resolution and the readback.  Set and
//! clear routes additionally include the candidate commit, reopen/readback
//! and serialization.

use std::hint::black_box;

use litchi_iwa::numbers;
use litchi_iwa::pages;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};
use litchi_iwa_common::chart::ChartData;
use litchi_iwa_common::chart::axis::Axis;
use litchi_iwa_common::chart::kind::Kind;
use litchi_numbers::{ChartSelector, Package as NumbersPackage, SheetSelector};
use litchi_pages::{BodyChartSelector, Package as PagesPackage};

use crate::harness::{Fixture, TestResult, fixture};
use crate::leaf_bytes;

const NUMBERS: &str = "iwork/numbers/chart-data-native.numbers";
const NUMBERS_CHART: usize = 1;
const AXES: [Axis; 2] = [Axis::Category, Axis::Value];

fn numbers_input() -> TestResult<(Fixture, u64, u64)> {
    let native = fixture(NUMBERS)?;
    let mut editor = numbers::Draft::from_bytes(&native.bytes)?;
    let sheet = editor
        .sheets()?
        .first()
        .ok_or("native Numbers sheet missing")?
        .id();
    let data = ChartData::new(
        vec!["Region 1".to_owned(), "Region 2".to_owned()],
        vec!["April".to_owned(), "May".to_owned(), "June".to_owned()],
        vec![
            vec![Some(17.0), Some(26.0), Some(53.0)],
            vec![Some(55.0), Some(43.0), Some(70.0)],
        ],
    )?;
    let chart = editor.add_sheet_chart(
        sheet,
        Kind::Column2d,
        data,
        DrawablePoint { x: 120.0, y: 140.0 },
        DrawableSize {
            width: 360.0,
            height: 240.0,
        },
    )?;
    Ok((
        Fixture {
            path: "chart-data-native.numbers + add_sheet_chart",
            bytes: editor.to_bytes()?,
        },
        sheet,
        chart.drawable_object_id,
    ))
}

fn numbers_input_with_text(axis: Option<Axis>) -> TestResult<(Fixture, u64, u64)> {
    let (input, sheet, chart) = numbers_input()?;
    let mut editor = numbers::Draft::from_bytes(&input.bytes)?;
    match axis {
        Some(_) => {
            for axis in AXES {
                editor.set_sheet_chart_axis_title(sheet, chart, axis, "cost seed")?;
            }
        },
        None => editor.set_sheet_chart_title(sheet, chart, "cost seed")?,
    }
    Ok((
        Fixture {
            path: "chart-data-native.numbers + add_sheet_chart + cost seed",
            bytes: editor.to_bytes()?,
        },
        sheet,
        chart,
    ))
}

fn pages_input() -> TestResult<(Fixture, u64)> {
    let mut editor = pages::Draft::create_with_text("W1-E chart text cost")?;
    let chart = editor
        .add_body_chart(
            "W1-E chart text cost".encode_utf16().count(),
            Kind::Column2d,
            ChartData::new(
                vec!["Region 1".to_owned(), "Region 2".to_owned()],
                vec!["April".to_owned(), "May".to_owned(), "June".to_owned()],
                vec![
                    vec![Some(17.0), Some(26.0), Some(53.0)],
                    vec![Some(55.0), Some(43.0), Some(70.0)],
                ],
            )?,
            DrawablePoint { x: 120.0, y: 140.0 },
            DrawableSize {
                width: 360.0,
                height: 240.0,
            },
        )?
        .drawable_object_id;
    Ok((
        Fixture {
            path: "source-built: Pages chart title and axis cost",
            bytes: editor.to_bytes()?,
        },
        chart,
    ))
}

fn pages_input_with_text(axis: Option<Axis>) -> TestResult<(Fixture, u64)> {
    let (input, chart) = pages_input()?;
    let mut editor = pages::Draft::from_bytes(&input.bytes)?;
    match axis {
        Some(_) => {
            for axis in AXES {
                editor.set_body_chart_axis_title(chart, axis, "cost seed")?;
            }
        },
        None => editor.set_body_chart_title(chart, "cost seed")?,
    }
    Ok((
        Fixture {
            path: "source-built: Pages chart title and axis cost + cost seed",
            bytes: editor.to_bytes()?,
        },
        chart,
    ))
}

#[inline(never)]
fn numbers_title_read_host(source: &[u8], sheet: u64, chart: u64) -> TestResult<()> {
    let editor = numbers::Draft::from_bytes(source)?;
    black_box(editor.sheet_chart_title(sheet, chart)?);
    Ok(())
}

#[inline(never)]
fn numbers_title_read_leaf(source: &[u8]) -> TestResult<()> {
    let package = NumbersPackage::from_bytes(source)?;
    black_box(
        package.sheet_chart_title(SheetSelector::index(0), ChartSelector::index(NUMBERS_CHART))?,
    );
    Ok(())
}

#[inline(never)]
fn numbers_axis_read_host(source: &[u8], sheet: u64, chart: u64) -> TestResult<()> {
    let editor = numbers::Draft::from_bytes(source)?;
    for axis in AXES {
        black_box(editor.sheet_chart_axis_title(sheet, chart, axis)?);
    }
    Ok(())
}

#[inline(never)]
fn numbers_axis_read_leaf(source: &[u8]) -> TestResult<()> {
    let package = NumbersPackage::from_bytes(source)?;
    for axis in AXES {
        black_box(package.sheet_chart_axis_title(
            SheetSelector::index(0),
            ChartSelector::index(NUMBERS_CHART),
            axis,
        )?);
    }
    Ok(())
}

#[inline(never)]
fn numbers_title_set_host(source: &[u8], sheet: u64, chart: u64) -> TestResult<Vec<u8>> {
    let mut editor = numbers::Draft::from_bytes(source)?;
    editor.set_sheet_chart_title(sheet, chart, "cost probe")?;
    Ok(editor.to_bytes()?)
}

#[inline(never)]
fn numbers_title_set_leaf(source: &[u8]) -> TestResult<Vec<u8>> {
    let package = NumbersPackage::from_bytes(source)?;
    let edit = package
        .edit_sheet_chart_title(SheetSelector::index(0), ChartSelector::index(NUMBERS_CHART))?
        .set("cost probe")?;
    let output = edit.commit()?.into_package();
    Ok(leaf_bytes!(&output))
}

#[inline(never)]
fn numbers_axis_set_host(source: &[u8], sheet: u64, chart: u64) -> TestResult<Vec<u8>> {
    let mut editor = numbers::Draft::from_bytes(source)?;
    for axis in AXES {
        editor.set_sheet_chart_axis_title(sheet, chart, axis, "cost probe")?;
    }
    Ok(editor.to_bytes()?)
}

#[inline(never)]
fn numbers_axis_set_leaf(source: &[u8]) -> TestResult<Vec<u8>> {
    let mut package = NumbersPackage::from_bytes(source)?;
    for axis in AXES {
        let edit = package
            .edit_sheet_chart_axis_title(
                SheetSelector::index(0),
                ChartSelector::index(NUMBERS_CHART),
                axis,
            )?
            .set("cost probe")?;
        package = edit.commit()?.into_package();
    }
    Ok(leaf_bytes!(&package))
}

#[inline(never)]
fn numbers_title_clear_host(source: &[u8], sheet: u64, chart: u64) -> TestResult<Vec<u8>> {
    let mut editor = numbers::Draft::from_bytes(source)?;
    black_box(editor.remove_sheet_chart_title(sheet, chart)?);
    Ok(editor.to_bytes()?)
}

#[inline(never)]
fn numbers_title_clear_leaf(source: &[u8]) -> TestResult<Vec<u8>> {
    let package = NumbersPackage::from_bytes(source)?;
    let edit = package
        .edit_sheet_chart_title(SheetSelector::index(0), ChartSelector::index(NUMBERS_CHART))?
        .clear()?;
    let output = edit.commit()?.into_package();
    Ok(leaf_bytes!(&output))
}

#[inline(never)]
fn numbers_axis_clear_host(source: &[u8], sheet: u64, chart: u64) -> TestResult<Vec<u8>> {
    let mut editor = numbers::Draft::from_bytes(source)?;
    for axis in AXES {
        black_box(editor.remove_sheet_chart_axis_title(sheet, chart, axis)?);
    }
    Ok(editor.to_bytes()?)
}

#[inline(never)]
fn numbers_axis_clear_leaf(source: &[u8]) -> TestResult<Vec<u8>> {
    let mut package = NumbersPackage::from_bytes(source)?;
    for axis in AXES {
        let edit = package
            .edit_sheet_chart_axis_title(
                SheetSelector::index(0),
                ChartSelector::index(NUMBERS_CHART),
                axis,
            )?
            .clear()?;
        package = edit.commit()?.into_package();
    }
    Ok(leaf_bytes!(&package))
}

#[inline(never)]
fn pages_title_read_host(source: &[u8], chart: u64) -> TestResult<()> {
    let editor = pages::Draft::from_bytes(source)?;
    black_box(editor.body_chart_title(chart)?);
    Ok(())
}

#[inline(never)]
fn pages_title_read_leaf(source: &[u8]) -> TestResult<()> {
    let package = PagesPackage::from_bytes(source)?;
    black_box(package.body_chart_title(BodyChartSelector::index(0))?);
    Ok(())
}

#[inline(never)]
fn pages_axis_read_host(source: &[u8], chart: u64) -> TestResult<()> {
    let editor = pages::Draft::from_bytes(source)?;
    for axis in AXES {
        black_box(editor.body_chart_axis_title(chart, axis)?);
    }
    Ok(())
}

#[inline(never)]
fn pages_axis_read_leaf(source: &[u8]) -> TestResult<()> {
    let package = PagesPackage::from_bytes(source)?;
    for axis in AXES {
        black_box(package.body_chart_axis_title(BodyChartSelector::index(0), axis)?);
    }
    Ok(())
}

#[inline(never)]
fn pages_title_set_host(source: &[u8], chart: u64) -> TestResult<Vec<u8>> {
    let mut editor = pages::Draft::from_bytes(source)?;
    editor.set_body_chart_title(chart, "cost probe")?;
    Ok(editor.to_bytes()?)
}

#[inline(never)]
fn pages_title_set_leaf(source: &[u8]) -> TestResult<Vec<u8>> {
    let package = PagesPackage::from_bytes(source)?;
    let edit = package
        .edit_body_chart_title(BodyChartSelector::index(0))?
        .set("cost probe")?;
    let output = edit.commit()?.into_package();
    Ok(leaf_bytes!(&output))
}

#[inline(never)]
fn pages_axis_set_host(source: &[u8], chart: u64) -> TestResult<Vec<u8>> {
    let mut editor = pages::Draft::from_bytes(source)?;
    for axis in AXES {
        editor.set_body_chart_axis_title(chart, axis, "cost probe")?;
    }
    Ok(editor.to_bytes()?)
}

#[inline(never)]
fn pages_axis_set_leaf(source: &[u8]) -> TestResult<Vec<u8>> {
    let mut package = PagesPackage::from_bytes(source)?;
    for axis in AXES {
        let edit = package
            .edit_body_chart_axis_title(BodyChartSelector::index(0), axis)?
            .set("cost probe")?;
        package = edit.commit()?.into_package();
    }
    Ok(leaf_bytes!(&package))
}

#[inline(never)]
fn pages_title_clear_host(source: &[u8], chart: u64) -> TestResult<Vec<u8>> {
    let mut editor = pages::Draft::from_bytes(source)?;
    black_box(editor.remove_body_chart_title(chart)?);
    Ok(editor.to_bytes()?)
}

#[inline(never)]
fn pages_title_clear_leaf(source: &[u8]) -> TestResult<Vec<u8>> {
    let package = PagesPackage::from_bytes(source)?;
    let edit = package
        .edit_body_chart_title(BodyChartSelector::index(0))?
        .clear()?;
    let output = edit.commit()?.into_package();
    Ok(leaf_bytes!(&output))
}

#[inline(never)]
fn pages_axis_clear_host(source: &[u8], chart: u64) -> TestResult<Vec<u8>> {
    let mut editor = pages::Draft::from_bytes(source)?;
    for axis in AXES {
        black_box(editor.remove_body_chart_axis_title(chart, axis)?);
    }
    Ok(editor.to_bytes()?)
}

#[inline(never)]
fn pages_axis_clear_leaf(source: &[u8]) -> TestResult<Vec<u8>> {
    let mut package = PagesPackage::from_bytes(source)?;
    for axis in AXES {
        let edit = package
            .edit_body_chart_axis_title(BodyChartSelector::index(0), axis)?
            .clear()?;
        package = edit.commit()?.into_package();
    }
    Ok(leaf_bytes!(&package))
}

#[test]
#[ignore = "callgrind cost probe; coordinator runs it after integration"]
fn numbers_chart_title_read_cost_abba() -> TestResult {
    let (input, sheet, chart) = numbers_input_with_text(None)?;
    black_box(numbers_title_read_host(&input.bytes, sheet, chart)?);
    black_box(numbers_title_read_leaf(&input.bytes)?);
    black_box(numbers_title_read_leaf(&input.bytes)?);
    black_box(numbers_title_read_host(&input.bytes, sheet, chart)?);
    Ok(())
}

#[test]
#[ignore = "callgrind cost probe; coordinator runs it after integration"]
fn numbers_chart_title_set_cost_abba() -> TestResult {
    let (input, sheet, chart) = numbers_input()?;
    black_box(numbers_title_set_host(&input.bytes, sheet, chart)?);
    black_box(numbers_title_set_leaf(&input.bytes)?);
    black_box(numbers_title_set_leaf(&input.bytes)?);
    black_box(numbers_title_set_host(&input.bytes, sheet, chart)?);
    Ok(())
}

#[test]
#[ignore = "callgrind cost probe; coordinator runs it after integration"]
fn numbers_chart_title_clear_cost_abba() -> TestResult {
    let (input, sheet, chart) = numbers_input_with_text(None)?;
    black_box(numbers_title_clear_host(&input.bytes, sheet, chart)?);
    black_box(numbers_title_clear_leaf(&input.bytes)?);
    black_box(numbers_title_clear_leaf(&input.bytes)?);
    black_box(numbers_title_clear_host(&input.bytes, sheet, chart)?);
    Ok(())
}

#[test]
#[ignore = "callgrind cost probe; coordinator runs it after integration"]
fn numbers_chart_axis_read_cost_abba() -> TestResult {
    let (input, sheet, chart) = numbers_input_with_text(Some(Axis::Category))?;
    black_box(numbers_axis_read_host(&input.bytes, sheet, chart)?);
    black_box(numbers_axis_read_leaf(&input.bytes)?);
    black_box(numbers_axis_read_leaf(&input.bytes)?);
    black_box(numbers_axis_read_host(&input.bytes, sheet, chart)?);
    Ok(())
}

#[test]
#[ignore = "callgrind cost probe; coordinator runs it after integration"]
fn numbers_chart_axis_set_cost_abba() -> TestResult {
    let (input, sheet, chart) = numbers_input()?;
    black_box(numbers_axis_set_host(&input.bytes, sheet, chart)?);
    black_box(numbers_axis_set_leaf(&input.bytes)?);
    black_box(numbers_axis_set_leaf(&input.bytes)?);
    black_box(numbers_axis_set_host(&input.bytes, sheet, chart)?);
    Ok(())
}

#[test]
#[ignore = "callgrind cost probe; coordinator runs it after integration"]
fn numbers_chart_axis_clear_cost_abba() -> TestResult {
    let (input, sheet, chart) = numbers_input_with_text(Some(Axis::Category))?;
    black_box(numbers_axis_clear_host(&input.bytes, sheet, chart)?);
    black_box(numbers_axis_clear_leaf(&input.bytes)?);
    black_box(numbers_axis_clear_leaf(&input.bytes)?);
    black_box(numbers_axis_clear_host(&input.bytes, sheet, chart)?);
    Ok(())
}

#[test]
#[ignore = "callgrind cost probe; coordinator runs it after integration"]
fn pages_chart_title_read_cost_abba() -> TestResult {
    let (input, chart) = pages_input_with_text(None)?;
    black_box(pages_title_read_host(&input.bytes, chart)?);
    black_box(pages_title_read_leaf(&input.bytes)?);
    black_box(pages_title_read_leaf(&input.bytes)?);
    black_box(pages_title_read_host(&input.bytes, chart)?);
    Ok(())
}

#[test]
#[ignore = "callgrind cost probe; coordinator runs it after integration"]
fn pages_chart_title_set_cost_abba() -> TestResult {
    let (input, chart) = pages_input()?;
    black_box(pages_title_set_host(&input.bytes, chart)?);
    black_box(pages_title_set_leaf(&input.bytes)?);
    black_box(pages_title_set_leaf(&input.bytes)?);
    black_box(pages_title_set_host(&input.bytes, chart)?);
    Ok(())
}

#[test]
#[ignore = "callgrind cost probe; coordinator runs it after integration"]
fn pages_chart_title_clear_cost_abba() -> TestResult {
    let (input, chart) = pages_input_with_text(None)?;
    black_box(pages_title_clear_host(&input.bytes, chart)?);
    black_box(pages_title_clear_leaf(&input.bytes)?);
    black_box(pages_title_clear_leaf(&input.bytes)?);
    black_box(pages_title_clear_host(&input.bytes, chart)?);
    Ok(())
}

#[test]
#[ignore = "callgrind cost probe; coordinator runs it after integration"]
fn pages_chart_axis_read_cost_abba() -> TestResult {
    let (input, chart) = pages_input_with_text(Some(Axis::Category))?;
    black_box(pages_axis_read_host(&input.bytes, chart)?);
    black_box(pages_axis_read_leaf(&input.bytes)?);
    black_box(pages_axis_read_leaf(&input.bytes)?);
    black_box(pages_axis_read_host(&input.bytes, chart)?);
    Ok(())
}

#[test]
#[ignore = "callgrind cost probe; coordinator runs it after integration"]
fn pages_chart_axis_set_cost_abba() -> TestResult {
    let (input, chart) = pages_input()?;
    black_box(pages_axis_set_host(&input.bytes, chart)?);
    black_box(pages_axis_set_leaf(&input.bytes)?);
    black_box(pages_axis_set_leaf(&input.bytes)?);
    black_box(pages_axis_set_host(&input.bytes, chart)?);
    Ok(())
}

#[test]
#[ignore = "callgrind cost probe; coordinator runs it after integration"]
fn pages_chart_axis_clear_cost_abba() -> TestResult {
    let (input, chart) = pages_input_with_text(Some(Axis::Category))?;
    black_box(pages_axis_clear_host(&input.bytes, chart)?);
    black_box(pages_axis_clear_leaf(&input.bytes)?);
    black_box(pages_axis_clear_leaf(&input.bytes)?);
    black_box(pages_axis_clear_host(&input.bytes, chart)?);
    Ok(())
}
