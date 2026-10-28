//! Callgrind ABBA probes for the W1-E chart title/axis-title slice.
//!
//! The probe includes open, selector resolution, one mutation, commit and
//! serialization on both sides. It is intentionally ignored until the
//! coordinator's release-mode pinned-CPU gate runs.

use std::hint::black_box;

use litchi_iwa::numbers;
use litchi_iwa::pages;
use litchi_iwa_common::chart::axis::Axis;
use litchi_numbers::{ChartSelector, Package as NumbersPackage, SheetSelector};
use litchi_pages::{BodyChartSelector, Package as PagesPackage};

use crate::harness::{TestResult, fixture};
use crate::leaf_bytes;

#[inline(never)]
fn chart_text_host_numbers(source: &[u8], axis: Option<Axis>) -> TestResult<Vec<u8>> {
    let mut editor = numbers::Draft::from_bytes(source)?;
    let sheet = editor
        .sheets()?
        .first()
        .ok_or("native Numbers sheet missing")?;
    let chart = editor
        .sheet_charts(sheet.id())?
        .first()
        .ok_or("native Numbers chart missing")?
        .drawable_object_id;
    match axis {
        Some(axis) => editor.set_sheet_chart_axis_title(sheet.id(), chart, axis, "cost probe")?,
        None => editor.set_sheet_chart_title(sheet.id(), chart, "cost probe")?,
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
                ChartSelector::index(0),
                axis,
            )?;
            edit = edit.set("cost probe")?;
            edit.commit()?.into_package()
        },
        None => {
            let mut edit =
                package.edit_sheet_chart_title(SheetSelector::index(0), ChartSelector::index(0))?;
            edit = edit.set("cost probe")?;
            edit.commit()?.into_package()
        },
    };
    Ok(leaf_bytes!(&output))
}

#[inline(never)]
fn chart_text_host_pages(source: &[u8], axis: Option<Axis>) -> TestResult<Vec<u8>> {
    let mut editor = pages::Draft::from_bytes(source)?;
    let chart = editor
        .body_charts()?
        .first()
        .ok_or("native Pages chart missing")?
        .drawable_object_id;
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
    let numbers = fixture("iwork/numbers/chart-arrangement-native.numbers")?;
    let pages = fixture("iwork/pages/chart-arrangement-native.pages")?;
    for axis in [None, Some(Axis::Category), Some(Axis::Value)] {
        black_box(chart_text_host_numbers(&numbers.bytes, axis)?);
        black_box(chart_text_leaf_numbers(&numbers.bytes, axis)?);
        black_box(chart_text_leaf_numbers(&numbers.bytes, axis)?);
        black_box(chart_text_host_numbers(&numbers.bytes, axis)?);
        black_box(chart_text_host_pages(&pages.bytes, axis)?);
        black_box(chart_text_leaf_pages(&pages.bytes, axis)?);
        black_box(chart_text_leaf_pages(&pages.bytes, axis)?);
        black_box(chart_text_host_pages(&pages.bytes, axis)?);
    }
    Ok(())
}
