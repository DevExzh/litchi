//! Callgrind ABBA probes for the W1-E chart title/axis-title slice.
//!
//! The probe includes open, selector resolution, one mutation, commit and
//! serialization on both sides. It is intentionally ignored until the
//! coordinator's release-mode pinned-CPU gate runs.

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

#[inline(never)]
fn chart_text_host_numbers(
    source: &[u8],
    sheet: u64,
    chart: u64,
    axis: Option<Axis>,
) -> TestResult<Vec<u8>> {
    let mut editor = numbers::Draft::from_bytes(source)?;
    match axis {
        Some(axis) => editor.set_sheet_chart_axis_title(sheet, chart, axis, "cost probe")?,
        None => editor.set_sheet_chart_title(sheet, chart, "cost probe")?,
    }
    Ok(editor.to_bytes()?)
}

#[inline(never)]
fn chart_text_leaf_numbers(source: &[u8], axis: Option<Axis>) -> TestResult<Vec<u8>> {
    let package = NumbersPackage::from_bytes(source)?;
    let output = match axis {
        Some(axis) => {
            let mut edit = package.edit_sheet_chart_axis_title(
                SheetSelector::index(0),
                ChartSelector::index(NUMBERS_CHART),
                axis,
            )?;
            edit = edit.set("cost probe")?;
            edit.commit()?.into_package()
        },
        None => {
            let mut edit = package.edit_sheet_chart_title(
                SheetSelector::index(0),
                ChartSelector::index(NUMBERS_CHART),
            )?;
            edit = edit.set("cost probe")?;
            edit.commit()?.into_package()
        },
    };
    Ok(leaf_bytes!(&output))
}

#[inline(never)]
fn chart_text_host_pages(source: &[u8], chart: u64, axis: Option<Axis>) -> TestResult<Vec<u8>> {
    let mut editor = pages::Draft::from_bytes(source)?;
    match axis {
        Some(axis) => editor.set_body_chart_axis_title(chart, axis, "cost probe")?,
        None => editor.set_body_chart_title(chart, "cost probe")?,
    }
    Ok(editor.to_bytes()?)
}

#[inline(never)]
fn chart_text_leaf_pages(source: &[u8], axis: Option<Axis>) -> TestResult<Vec<u8>> {
    let package = PagesPackage::from_bytes(source)?;
    let output = match axis {
        Some(axis) => {
            let mut edit = package.edit_body_chart_axis_title(BodyChartSelector::index(0), axis)?;
            edit = edit.set("cost probe")?;
            edit.commit()?.into_package()
        },
        None => {
            let mut edit = package.edit_body_chart_title(BodyChartSelector::index(0))?;
            edit = edit.set("cost probe")?;
            edit.commit()?.into_package()
        },
    };
    Ok(leaf_bytes!(&output))
}

#[test]
#[ignore = "callgrind cost probe; coordinator runs it after integration"]
fn chart_text_cost_abba() -> TestResult {
    let (numbers, sheet, chart) = numbers_input()?;
    let (pages, page_chart) = pages_input()?;
    for axis in [None, Some(Axis::Category), Some(Axis::Value)] {
        black_box(chart_text_host_numbers(&numbers.bytes, sheet, chart, axis)?);
        black_box(chart_text_leaf_numbers(&numbers.bytes, axis)?);
        black_box(chart_text_leaf_numbers(&numbers.bytes, axis)?);
        black_box(chart_text_host_numbers(&numbers.bytes, sheet, chart, axis)?);
        black_box(chart_text_host_pages(&pages.bytes, page_chart, axis)?);
        black_box(chart_text_leaf_pages(&pages.bytes, axis)?);
        black_box(chart_text_leaf_pages(&pages.bytes, axis)?);
        black_box(chart_text_host_pages(&pages.bytes, page_chart, axis)?);
    }
    Ok(())
}
