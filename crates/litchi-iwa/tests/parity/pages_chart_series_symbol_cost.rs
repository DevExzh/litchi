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
use litchi_pages::chart::series::{Property, PropertyKind, Symbol, SymbolShape, SymbolSize};
use litchi_pages::{BodyChartSelector, Package};

use crate::harness::{Fixture, TestResult, fixture};
use crate::leaf_bytes;

const NATIVE: &str = "iwork/pages/chart-data-native.pages";

#[derive(Clone, Copy)]
enum Route {
    AllRead,
    AllSet,
    SingleRead,
    SingleSet,
}

impl Route {
    const ALL: [Self; 4] = [
        Self::AllRead,
        Self::AllSet,
        Self::SingleRead,
        Self::SingleSet,
    ];

    const fn label(self) -> &'static str {
        match self {
            Self::AllRead => "all_symbols_read",
            Self::AllSet => "all_symbols_set",
            Self::SingleRead => "single_symbol_read",
            Self::SingleSet => "single_symbol_set",
        }
    }
}

const VALUES: [Option<(SymbolShape, Option<f32>)>; 2] = [
    Some((SymbolShape::Diamond, Some(9.0))),
    Some((SymbolShape::Square, None)),
];
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
        Route::AllRead => {
            let _ = black_box(editor.body_chart_series_symbols(id)?);
        },
        Route::AllSet => editor.set_body_chart_series_symbols(id, &host_values(&VALUES)?)?,
        Route::SingleRead => {
            let _ = black_box(
                editor
                    .body_chart_series_symbol(id, litchi_iwa::charts::Index::from_zero_based(0))?,
            );
        },
        Route::SingleSet => editor.set_body_chart_series_symbol(
            id,
            litchi_iwa::charts::Index::from_zero_based(0),
            host_values(&VALUES)?[0],
        )?,
    }
    Ok(editor.to_bytes()?)
}

#[inline(never)]
fn pages_symbols_leaf_route(source: &[u8], route: Route) -> TestResult<Vec<u8>> {
    let package = Package::from_bytes(source)?;
    match route {
        Route::AllRead => {
            black_box(
                package.body_chart_series(BodyChartSelector::index(0), PropertyKind::Symbols)?,
            );
            Ok(leaf_bytes!(&package))
        },
        Route::SingleRead => {
            let values = match package
                .body_chart_series(BodyChartSelector::index(0), PropertyKind::Symbols)?
            {
                Property::Symbols(values) => values,
                _ => return Err("unexpected Pages series property".into()),
            };
            black_box(values.first().copied());
            Ok(leaf_bytes!(&package))
        },
        Route::AllSet => {
            let mut edit = package.edit_body_chart_series(BodyChartSelector::index(0))?;
            edit = edit.set(Property::Symbols(leaf_values(&VALUES)?))?;
            Ok(leaf_bytes!(edit.commit()?.package()))
        },
        Route::SingleSet => {
            let current = match package
                .body_chart_series(BodyChartSelector::index(0), PropertyKind::Symbols)?
            {
                Property::Symbols(values) => values,
                _ => return Err("unexpected Pages series property".into()),
            };
            let mut values = current;
            values[0] = leaf_values(&VALUES)?[0];
            let mut edit = package.edit_body_chart_series(BodyChartSelector::index(0))?;
            edit = edit.set(Property::Symbols(values))?;
            Ok(leaf_bytes!(edit.commit()?.package()))
        },
    }
}

#[inline(never)]
fn pages_symbols_host_all_read(source: &[u8]) -> TestResult<Vec<u8>> {
    pages_symbols_host_route(source, Route::AllRead)
}

#[inline(never)]
fn pages_symbols_host_all_set(source: &[u8]) -> TestResult<Vec<u8>> {
    pages_symbols_host_route(source, Route::AllSet)
}

#[inline(never)]
fn pages_symbols_host_single_read(source: &[u8]) -> TestResult<Vec<u8>> {
    pages_symbols_host_route(source, Route::SingleRead)
}

#[inline(never)]
fn pages_symbols_host_single_set(source: &[u8]) -> TestResult<Vec<u8>> {
    pages_symbols_host_route(source, Route::SingleSet)
}

#[inline(never)]
fn pages_symbols_leaf_all_read(source: &[u8]) -> TestResult<Vec<u8>> {
    pages_symbols_leaf_route(source, Route::AllRead)
}

#[inline(never)]
fn pages_symbols_leaf_all_set(source: &[u8]) -> TestResult<Vec<u8>> {
    pages_symbols_leaf_route(source, Route::AllSet)
}

#[inline(never)]
fn pages_symbols_leaf_single_read(source: &[u8]) -> TestResult<Vec<u8>> {
    pages_symbols_leaf_route(source, Route::SingleRead)
}

#[inline(never)]
fn pages_symbols_leaf_single_set(source: &[u8]) -> TestResult<Vec<u8>> {
    pages_symbols_leaf_route(source, Route::SingleSet)
}

fn host_route(source: &[u8], route: Route) -> TestResult<Vec<u8>> {
    match route {
        Route::AllRead => pages_symbols_host_all_read(source),
        Route::AllSet => pages_symbols_host_all_set(source),
        Route::SingleRead => pages_symbols_host_single_read(source),
        Route::SingleSet => pages_symbols_host_single_set(source),
    }
}

fn leaf_route(source: &[u8], route: Route) -> TestResult<Vec<u8>> {
    match route {
        Route::AllRead => pages_symbols_leaf_all_read(source),
        Route::AllSet => pages_symbols_leaf_all_set(source),
        Route::SingleRead => pages_symbols_leaf_single_read(source),
        Route::SingleSet => pages_symbols_leaf_single_set(source),
    }
}

/// Full-facade Pages symbol routes. Set `W2B_SYMBOL_ROUTE` to
/// `all_symbols_read`, `all_symbols_set`, `single_symbol_read`, or
/// `single_symbol_set` to isolate one route. Hide remains a parity control,
/// outside the measured wrappers.
#[test]
#[ignore = "full-facade release cost probe; run explicitly under callgrind"]
fn pages_chart_series_symbols_facade_cost_abba() -> TestResult {
    let filter = std::env::var("W2B_SYMBOL_ROUTE").ok();
    let source = input()?.bytes;
    for route in Route::ALL {
        if filter
            .as_deref()
            .is_some_and(|selected| selected != route.label())
        {
            continue;
        }
        black_box(host_route(&source, route)?);
        black_box(leaf_route(&source, route)?);
        black_box(leaf_route(&source, route)?);
        black_box(host_route(&source, route)?);
    }
    Ok(())
}
