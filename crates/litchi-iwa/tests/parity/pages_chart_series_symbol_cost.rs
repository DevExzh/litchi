//! Full-facade Pages chart-series-symbol ABBA cost probes (W2-B).
//!
//! The counted wrappers include package open, selector resolution, semantic
//! read or edit, candidate reopen/readback and serialization.  They are
//! intentionally ignored and never perform wall-clock timing.

use std::hint::black_box;

use litchi_iwa::charts::{ChartSeriesSymbol, ChartSeriesSymbolShape, ChartSeriesSymbolSize};
use litchi_iwa::pages;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};
use litchi_iwa_common::chart::ChartData;
use litchi_iwa_common::chart::kind::Kind;
use litchi_pages::chart::series::{Property, Symbol, SymbolShape, SymbolSize};
use litchi_pages::{BodyChartSelector, Package};

use crate::harness::{Fixture, TestResult, fixture};
use crate::leaf_bytes;

const NATIVE: &str = "iwork/pages/chart-data-native.pages";

#[derive(Clone, Copy)]
enum Route {
    Read,
    Set,
    Hide,
}

impl Route {
    const ALL: [Self; 3] = [Self::Read, Self::Set, Self::Hide];

    const fn label(self) -> &'static str {
        match self {
            Self::Read => "symbols_read",
            Self::Set => "symbols_set",
            Self::Hide => "symbols_hide",
        }
    }
}

const VALUES: [Option<(SymbolShape, Option<f32>)>; 2] = [
    Some((SymbolShape::Diamond, Some(9.0))),
    Some((SymbolShape::Square, None)),
];
const HIDDEN: [Option<(SymbolShape, Option<f32>)>; 2] =
    [Some((SymbolShape::Square, Some(12.0))), None];

fn data() -> TestResult<ChartData> {
    Ok(ChartData::new(
        vec!["Region 1".to_owned(), "Region 2".to_owned()],
        vec!["April".to_owned(), "May".to_owned(), "June".to_owned()],
        vec![
            vec![Some(17.0), Some(26.0), Some(53.0)],
            vec![Some(55.0), Some(43.0), Some(70.0)],
        ],
    )?)
}

fn input() -> TestResult<Fixture> {
    let native = fixture(NATIVE)?;
    let mut editor = pages::Draft::from_bytes(&native.bytes)?;
    editor.add_body_chart(
        0,
        Kind::Line2d,
        data()?,
        DrawablePoint { x: 120.0, y: 140.0 },
        DrawableSize {
            width: 360.0,
            height: 240.0,
        },
    )?;
    Ok(Fixture {
        path: "chart-data-native.pages + host-added Line2d chart",
        bytes: editor.to_bytes()?,
    })
}

fn host_values(
    values: &[Option<(SymbolShape, Option<f32>)>],
) -> TestResult<Vec<Option<ChartSeriesSymbol>>> {
    values
        .iter()
        .map(|value| {
            value
                .map(|(shape, size)| {
                    let shape = match shape {
                        SymbolShape::Circle => ChartSeriesSymbolShape::Circle,
                        SymbolShape::Triangle => ChartSeriesSymbolShape::Triangle,
                        SymbolShape::InvertedTriangle => ChartSeriesSymbolShape::InvertedTriangle,
                        SymbolShape::Diamond => ChartSeriesSymbolShape::Diamond,
                        SymbolShape::Square => ChartSeriesSymbolShape::Square,
                    };
                    match size {
                        Some(points) => Ok(ChartSeriesSymbol::sized(
                            shape,
                            ChartSeriesSymbolSize::new(points)?,
                        )),
                        None => Ok(ChartSeriesSymbol::automatic(shape)),
                    }
                })
                .transpose()
        })
        .collect()
}

fn leaf_values(values: &[Option<(SymbolShape, Option<f32>)>]) -> TestResult<Vec<Option<Symbol>>> {
    values
        .iter()
        .map(|value| {
            value
                .map(|(shape, size)| {
                    let size = size.map(SymbolSize::new).transpose()?;
                    Ok(match size {
                        Some(size) => Symbol::sized(shape, size),
                        None => Symbol::automatic(shape),
                    })
                })
                .transpose()
        })
        .collect()
}

fn host_id(editor: &pages::Draft) -> TestResult<u64> {
    editor
        .body_charts()?
        .first()
        .map(|chart| chart.drawable_object_id)
        .ok_or_else(|| "host-added Pages chart missing".into())
}

#[inline(never)]
fn pages_symbols_host_route(source: &[u8], route: Route) -> TestResult<Vec<u8>> {
    let mut editor = pages::Draft::from_bytes(source)?;
    let id = host_id(&editor)?;
    match route {
        Route::Read => {
            let _ = black_box(editor.body_chart_series_symbols(id)?);
        },
        Route::Set => editor.set_body_chart_series_symbols(id, &host_values(&VALUES)?)?,
        Route::Hide => editor.set_body_chart_series_symbols(id, &host_values(&HIDDEN)?)?,
    }
    Ok(editor.to_bytes()?)
}

#[inline(never)]
fn pages_symbols_leaf_route(source: &[u8], route: Route) -> TestResult<Vec<u8>> {
    let package = Package::from_bytes(source)?;
    if matches!(route, Route::Read) {
        black_box(package.body_chart_series(BodyChartSelector::index(0))?);
        return Ok(leaf_bytes!(&package));
    }
    let values = if matches!(route, Route::Set) {
        &VALUES
    } else {
        &HIDDEN
    };
    let mut edit = package.edit_body_chart_series(BodyChartSelector::index(0))?;
    edit = edit.set(Property::Symbols(leaf_values(values)?))?;
    Ok(leaf_bytes!(edit.commit()?.package()))
}

/// Full-facade Pages symbol routes.  Set `W2B_SYMBOL_ROUTE` to
/// `symbols_read`, `symbols_set`, or `symbols_hide` to isolate one symbol.
#[test]
#[ignore = "full-facade release cost probe; run explicitly under callgrind"]
fn pages_chart_series_symbols_facade_cost_abba() -> TestResult {
    let filter = std::env::var("W2B_SYMBOL_ROUTE").ok();
    let mut source = input()?.bytes;
    if filter.as_deref() == Some("symbols_hide") {
        source = pages_symbols_host_route(&source, Route::Set)?;
    }
    for route in Route::ALL {
        if filter
            .as_deref()
            .is_some_and(|selected| selected != route.label())
        {
            continue;
        }
        black_box(pages_symbols_host_route(&source, route)?);
        black_box(pages_symbols_leaf_route(&source, route)?);
        black_box(pages_symbols_leaf_route(&source, route)?);
        black_box(pages_symbols_host_route(&source, route)?);
    }
    Ok(())
}
