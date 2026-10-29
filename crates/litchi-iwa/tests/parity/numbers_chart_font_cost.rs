//! Full-facade host-vs-leaf cost probes for the six Numbers chart-font rows.
//!
//! The named route functions include package open, selector resolution, the
//! read or staged edit, candidate reopen/readback for writes, and package
//! serialization.  Run the ignored test from a release build under the
//! coordinator's pinned CPU/callgrind procedure; this module intentionally
//! performs no wall-clock timing.

use std::hint::black_box;

use litchi_iwa::charts::{ChartFont as HostFont, ChartFontSize};
use litchi_iwa::numbers;
use litchi_iwa_common::chart::kind::Kind;
use litchi_iwa_style::chart_font::{Change, ChartFont};
use litchi_iwa_text::character::TextPointSize;
use litchi_iwa_text::font::Font;
use litchi_numbers::Package;
use litchi_numbers::sheet::chart::font::{Property, PropertyKind};

use crate::chart_style_cases::{Target, input};
use crate::harness::{Format, TestResult};
use crate::leaf_bytes;

const CHART: usize = 1;

#[derive(Clone, Copy)]
enum Route {
    FontRead,
    FontSet,
    FontReset,
    SizeRead,
    SizeSet,
    SizeReset,
}

impl Route {
    const ALL: [Self; 6] = [
        Self::FontRead,
        Self::FontSet,
        Self::FontReset,
        Self::SizeRead,
        Self::SizeSet,
        Self::SizeReset,
    ];

    const fn kind(self) -> PropertyKind {
        match self {
            Self::FontRead | Self::FontSet | Self::FontReset => PropertyKind::Font,
            Self::SizeRead | Self::SizeSet | Self::SizeReset => PropertyKind::Size,
        }
    }

    const fn is_reset(self) -> bool {
        matches!(self, Self::FontReset | Self::SizeReset)
    }
}

fn avenir() -> TestResult<Change> {
    Ok(Change::Font(ChartFont {
        font: Font::named("AvenirNext-DemiBold")?,
        bold: true,
        italic: false,
    }))
}

fn size() -> TestResult<Change> {
    Ok(Change::Size(TextPointSize::from_points(17.0)?))
}

fn host_font(font: &ChartFont) -> TestResult<HostFont> {
    let base = match font.font.name() {
        Some(name) => HostFont::named(name)?,
        None => HostFont::default(),
    };
    Ok(base.with_bold(font.bold).with_italic(font.italic))
}

#[inline(never)]
fn numbers_chart_font_reset_setup(
    source: &[u8],
    sheet: u64,
    chart: u64,
    route: Route,
) -> TestResult<Vec<u8>> {
    let mut editor = numbers::Draft::from_bytes(source)?;
    match route {
        Route::FontRead => {
            black_box(editor.sheet_chart_font(sheet, chart)?);
        },
        Route::FontSet => {
            let Change::Font(value) = avenir()? else {
                return Err("font probe did not construct a font".into());
            };
            editor.set_sheet_chart_font(sheet, chart, host_font(&value)?)?;
        },
        Route::FontReset => {
            editor.reset_sheet_chart_font(sheet, chart)?;
        },
        Route::SizeRead => {
            black_box(editor.sheet_chart_font_size(sheet, chart)?);
        },
        Route::SizeSet => {
            let Change::Size(value) = size()? else {
                return Err("size probe did not construct a size".into());
            };
            editor.set_sheet_chart_font_size(sheet, chart, ChartFontSize::from(value))?;
        },
        Route::SizeReset => {
            editor.reset_sheet_chart_font_size(sheet, chart)?;
        },
    }
    Ok(editor.to_bytes()?)
}

/// Counted host wrapper. Reset seeding calls the implementation directly so
/// setup does not add a third selected host call.
#[inline(never)]
fn numbers_chart_font_host_route(
    source: &[u8],
    sheet: u64,
    chart: u64,
    route: Route,
) -> TestResult<Vec<u8>> {
    numbers_chart_font_reset_setup(source, sheet, chart, route)
}

#[inline(never)]
fn numbers_chart_font_leaf_route(source: &[u8], route: Route) -> TestResult<Vec<u8>> {
    let package = Package::from_bytes(source)?;
    if matches!(route, Route::FontRead | Route::SizeRead) {
        black_box(package.sheet_chart_font(0, CHART, route.kind())?);
        return Ok(leaf_bytes!(&package));
    }
    let mut edit = package.edit_sheet_chart_font(0, CHART)?;
    match route {
        Route::FontSet => {
            let Change::Font(value) = avenir()? else {
                return Err("font probe did not construct a font".into());
            };
            edit.set(Property::Font(value))?;
        },
        Route::FontReset => {
            edit.reset(PropertyKind::Font)?;
        },
        Route::SizeSet => {
            let Change::Size(value) = size()? else {
                return Err("size probe did not construct a size".into());
            };
            edit.set(Property::Size(value))?;
        },
        Route::SizeReset => {
            edit.reset(PropertyKind::Size)?;
        },
        Route::FontRead | Route::SizeRead => unreachable!(),
    }
    Ok(leaf_bytes!(edit.commit()?.package()))
}

/// ABBA probe for all six full Numbers facade routes.  Set `W2A_FONT_ROUTE`
/// to `font_read`, `font_set`, `font_reset`, `size_read`, `size_set`, or
/// `size_reset` when collecting one callgrind symbol at a time.
#[test]
#[ignore = "full-facade release cost probe; run explicitly under callgrind"]
fn numbers_chart_font_facade_cost_abba() -> TestResult {
    let filter = std::env::var("W2A_FONT_ROUTE").ok();
    for route in Route::ALL {
        let label = match route {
            Route::FontRead => "font_read",
            Route::FontSet => "font_set",
            Route::FontReset => "font_reset",
            Route::SizeRead => "size_read",
            Route::SizeSet => "size_set",
            Route::SizeReset => "size_reset",
        };
        if filter.as_deref().is_some_and(|selected| selected != label) {
            continue;
        }
        let mut input = input(Format::Numbers, Kind::Column2d)?;
        let Target::Numbers { sheet } = input.target else {
            return Err("Numbers input did not provide a sheet target".into());
        };
        let chart = input.chart;
        if route.is_reset() {
            let seed = if route.kind() == PropertyKind::Font {
                Route::FontSet
            } else {
                Route::SizeSet
            };
            input.fixture.bytes =
                numbers_chart_font_reset_setup(&input.fixture.bytes, sheet, chart, seed)?;
        }
        let source = input.fixture.bytes.as_slice();
        black_box(numbers_chart_font_host_route(source, sheet, chart, route)?);
        black_box(numbers_chart_font_leaf_route(source, route)?);
        black_box(numbers_chart_font_leaf_route(source, route)?);
        black_box(numbers_chart_font_host_route(source, sheet, chart, route)?);
    }
    Ok(())
}
