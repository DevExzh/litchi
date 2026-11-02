//! Pages body chart series-style parity (wave W4-C, record 0414): the
//! host `set_/reset_sheet_chart_series_{stroke,fill,symbol_fill,
//! symbol_outline}` routes (all-series, one-series and reset) against
//! `litchi_pages::chart::series::Edit::{set, reset}` on W1-E's
//! series-style write rules.

use litchi_iwa::charts::{
    ChartSeriesStroke, ChartSeriesStrokePattern, ChartSeriesSymbolFill, Index,
};
use litchi_iwa::pages;
use litchi_iwa::shapes::ShapeFill;
use litchi_iwa_common::chart::kind::Kind;
use litchi_iwa_common::chart::series::{Stroke, StrokePattern, SymbolFill};
use litchi_iwa_common::color::{RgbColorSpace, Rgba};
use litchi_iwa_common::shape::fill::{Angle, Fill, Gradient};
use litchi_iwa_common::shape::stroke::Width;
use litchi_pages::Package;
use litchi_pages::chart::series::{Property, PropertyKind};

use crate::chart_style_cases::input;
use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier};
use crate::leaf_bytes;

fn rgb(red: f32, green: f32, blue: f32) -> TestResult<Rgba> {
    Ok(Rgba::new(red, green, blue, 1.0, RgbColorSpace::Srgb)?)
}

fn stroke(red: f32, width: f32, pattern: StrokePattern) -> TestResult<Stroke> {
    Ok(Stroke::new(
        rgb(red, 0.2, 0.3)?,
        Width::new(width)?,
        pattern,
    ))
}

#[derive(Clone, Debug)]
enum Values {
    Strokes(Vec<Option<Stroke>>),
    Fills(Vec<Fill>),
    SymbolFills(Vec<SymbolFill>),
    Outlines(Vec<Option<Stroke>>),
}

/// One write: every series, one series, or one series' reset.
#[derive(Clone, Copy, Debug)]
enum Route {
    All,
    One(usize),
    Reset(usize),
}

impl Values {
    const fn chart(&self) -> Kind {
        match self {
            Self::Fills(_) => Kind::Column2d,
            _ => Kind::Line2d,
        }
    }

    const fn kind(&self) -> PropertyKind {
        match self {
            Self::Strokes(_) => PropertyKind::Strokes,
            Self::Fills(_) => PropertyKind::Fills,
            Self::SymbolFills(_) => PropertyKind::SymbolFills,
            Self::Outlines(_) => PropertyKind::SymbolOutlines,
        }
    }

    fn leaf(&self) -> Property {
        match self {
            Self::Strokes(v) => Property::Strokes(v.clone()),
            Self::Fills(v) => Property::Fills(v.clone()),
            Self::SymbolFills(v) => Property::SymbolFills(v.clone()),
            Self::Outlines(v) => Property::SymbolOutlines(v.clone()),
        }
    }
}

fn host_stroke(stroke: Stroke) -> ChartSeriesStroke {
    let pattern = match stroke.pattern {
        StrokePattern::Solid => ChartSeriesStrokePattern::Solid,
        StrokePattern::MediumDash => ChartSeriesStrokePattern::MediumDash,
        StrokePattern::RoundedDash => ChartSeriesStrokePattern::RoundedDash,
    };
    ChartSeriesStroke::new(stroke.color, stroke.width, pattern)
}

fn host_fill(fill: &Fill) -> TestResult<ShapeFill> {
    Ok(match fill {
        Fill::None => ShapeFill::None,
        Fill::Solid(color) => ShapeFill::Solid(*color),
        Fill::Gradient(gradient) => ShapeFill::Gradient(gradient.clone()),
        Fill::Image(_) => return Err("image fills are not covered".into()),
    })
}

fn host_symbol_fill(fill: &SymbolFill) -> TestResult<ChartSeriesSymbolFill> {
    Ok(match fill {
        SymbolFill::SeriesFill => ChartSeriesSymbolFill::SeriesFill,
        SymbolFill::SeriesStroke => ChartSeriesSymbolFill::SeriesStroke,
        SymbolFill::Custom(fill) => ChartSeriesSymbolFill::Custom(host_fill(fill)?),
    })
}

fn host_write(
    source: &[u8],
    __sheet: (),
    chart: u64,
    values: &Values,
    route: Route,
) -> TestResult<Vec<u8>> {
    let mut e = pages::Draft::from_bytes(source)?;
    let at = Index::from_zero_based;
    match (values, route) {
        (Values::Strokes(v), Route::All) => e.set_body_chart_series_strokes(
            chart,
            &v.iter().map(|s| s.map(host_stroke)).collect::<Vec<_>>(),
        )?,
        (Values::Strokes(v), Route::One(i)) => {
            e.set_body_chart_series_stroke(chart, at(i), v[i].map(host_stroke))?;
        },
        (Values::Strokes(_), Route::Reset(i)) => {
            e.reset_body_chart_series_stroke(chart, at(i))?;
        },
        (Values::Outlines(v), Route::All) => e.set_body_chart_series_symbol_outlines(
            chart,
            &v.iter().map(|s| s.map(host_stroke)).collect::<Vec<_>>(),
        )?,
        (Values::Outlines(v), Route::One(i)) => {
            e.set_body_chart_series_symbol_outline(chart, at(i), v[i].map(host_stroke))?;
        },
        (Values::Outlines(_), Route::Reset(i)) => {
            e.reset_body_chart_series_symbol_outline(chart, at(i))?;
        },
        (Values::Fills(v), Route::All) => e.set_body_chart_series_fills(
            chart,
            &v.iter().map(host_fill).collect::<TestResult<Vec<_>>>()?,
        )?,
        (Values::Fills(v), Route::One(i)) => {
            e.set_body_chart_series_fill(chart, at(i), &host_fill(&v[i])?)?;
        },
        (Values::Fills(_), Route::Reset(i)) => {
            e.reset_body_chart_series_fill(chart, at(i))?;
        },
        (Values::SymbolFills(v), Route::All) => e.set_body_chart_series_symbol_fills(
            chart,
            &v.iter()
                .map(host_symbol_fill)
                .collect::<TestResult<Vec<_>>>()?,
        )?,
        (Values::SymbolFills(v), Route::One(i)) => {
            e.set_body_chart_series_symbol_fill(chart, at(i), host_symbol_fill(&v[i])?)?;
        },
        (Values::SymbolFills(_), Route::Reset(i)) => {
            e.reset_body_chart_series_symbol_fill(chart, at(i))?;
        },
    }
    Ok(e.to_bytes()?)
}

fn host_read(bytes: &[u8], __sheet: (), chart: u64, values: &Values) -> TestResult<String> {
    let e = pages::Draft::from_bytes(bytes)?;
    Ok(match values {
        Values::Strokes(_) => format!("{:?}", e.body_chart_series_strokes(chart)?),
        Values::Fills(_) => format!("{:?}", e.body_chart_series_fills(chart)?),
        Values::SymbolFills(_) => format!("{:?}", e.body_chart_series_symbol_fills(chart)?),
        Values::Outlines(_) => format!("{:?}", e.body_chart_series_symbol_outlines(chart)?),
    })
}

/// The host-added chart is the body's only discoverable chart.
fn last_chart(_bytes: &[u8]) -> TestResult<usize> {
    Ok(0)
}

fn replace_one(current: &mut Property, wanted: &Property, i: usize) -> TestResult {
    match (current, wanted) {
        (Property::Strokes(c), Property::Strokes(w))
        | (Property::SymbolOutlines(c), Property::SymbolOutlines(w)) => c[i] = w[i],
        (Property::Fills(c), Property::Fills(w)) => c[i] = w[i].clone(),
        (Property::SymbolFills(c), Property::SymbolFills(w)) => c[i] = w[i].clone(),
        _ => return Err("property mismatch".into()),
    }
    Ok(())
}

fn leaf_write(source: &[u8], values: &Values, route: Route) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let chart = last_chart(source)?;
    let edit = package.edit_body_chart_series(chart)?;
    let edit = match route {
        Route::All => edit.set(values.leaf())?,
        Route::One(i) => {
            let mut current = package.body_chart_series(chart, values.kind())?;
            replace_one(&mut current, &values.leaf(), i)?;
            edit.set(current)?
        },
        Route::Reset(i) => edit.reset(values.kind(), i)?,
    };
    let commit = edit.commit()?;
    let restored = commit
        .package()
        .apply_body_chart_series(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

fn run(name: &'static str, before: Option<&Values>, values: &Values, route: Route) -> TestResult {
    let mut input = input(Format::Pages, values.chart())?;
    let sheet = ();
    let chart = input.chart;
    if let Some(before) = before {
        input.fixture.bytes = host_write(&input.fixture.bytes, sheet, chart, before, Route::All)?;
    }
    let fixture: &Fixture = &input.fixture;
    Case::new(name, Format::Pages, Tier::A, Evidence::Queued, fixture)
        .host(|source| host_write(source, sheet, chart, values, route))
        .leaf(|source| leaf_write(source, values, route))
        .read("host series style", move |bytes| {
            host_read(bytes, sheet, chart, values)
        })
        .run()
        .map_err(|error| format!("{values:?} {route:?}: {error}"))?;
    Ok(())
}

fn strokes() -> TestResult<(Values, Values)> {
    Ok((
        Values::Strokes(vec![
            Some(stroke(0.8, 2.5, StrokePattern::MediumDash)?),
            None,
        ]),
        Values::Strokes(vec![
            Some(stroke(0.1, 1.0, StrokePattern::Solid)?),
            Some(stroke(0.6, 3.0, StrokePattern::RoundedDash)?),
        ]),
    ))
}

fn fills() -> TestResult<(Values, Values)> {
    Ok((
        Values::Fills(vec![
            Fill::Solid(rgb(0.1, 0.4, 0.9)?),
            Fill::Gradient(Gradient::linear(
                rgb(0.9, 0.9, 0.1)?,
                rgb(0.1, 0.6, 0.3)?,
                Angle::from_degrees(90.0)?,
            )),
        ]),
        Values::Fills(vec![
            Fill::Gradient(Gradient::linear(
                rgb(0.2, 0.2, 0.2)?,
                rgb(0.8, 0.8, 0.8)?,
                Angle::from_degrees(45.0)?,
            )),
            Fill::Solid(rgb(0.7, 0.1, 0.1)?),
        ]),
    ))
}

fn symbol_fills() -> TestResult<(Values, Values)> {
    Ok((
        Values::SymbolFills(vec![
            SymbolFill::SeriesFill,
            SymbolFill::Custom(Fill::Solid(rgb(0.2, 0.7, 0.2)?)),
        ]),
        Values::SymbolFills(vec![
            SymbolFill::Custom(Fill::Solid(rgb(0.9, 0.5, 0.1)?)),
            SymbolFill::SeriesStroke,
        ]),
    ))
}

fn outlines() -> TestResult<(Values, Values)> {
    Ok((
        Values::Outlines(vec![Some(stroke(0.3, 1.5, StrokePattern::Solid)?), None]),
        Values::Outlines(vec![
            None,
            Some(stroke(0.9, 2.0, StrokePattern::RoundedDash)?),
        ]),
    ))
}

fn family(name: &'static str, (a, b): (Values, Values), stored: usize) -> TestResult {
    run(name, None, &a, Route::All)?;
    run(name, Some(&a), &b, Route::All)?;
    run(name, Some(&a), &b, Route::One(1))?;
    run(name, Some(&a), &a, Route::Reset(0))?;
    run(name, Some(&b), &b, Route::Reset(stored))?;
    Ok(())
}

#[test]
fn pages_body_chart_series_strokes_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    family("pages_body_chart_series_strokes_parity", strokes()?, 1)
}

#[test]
fn pages_body_chart_series_fills_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    family("pages_body_chart_series_fills_parity", fills()?, 1)
}

#[test]
fn pages_body_chart_series_symbol_fills_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    family(
        "pages_body_chart_series_symbol_fills_parity",
        symbol_fills()?,
        0,
    )
}

#[test]
fn pages_body_chart_series_symbol_outlines_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    family(
        "pages_body_chart_series_symbol_outlines_parity",
        outlines()?,
        1,
    )
}

#[test]
fn pages_body_chart_series_style_read_parity() -> TestResult {
    for (values, _) in [strokes()?, fills()?, symbol_fills()?, outlines()?] {
        let mut input = input(Format::Pages, values.chart())?;
        let sheet = ();
        input.fixture.bytes = host_write(
            &input.fixture.bytes,
            sheet,
            input.chart,
            &values,
            Route::All,
        )?;
        let chart = last_chart(&input.fixture.bytes)?;
        let read =
            Package::from_bytes(&input.fixture.bytes)?.body_chart_series(chart, values.kind())?;
        assert_eq!(read, values.leaf(), "{values:?}");
    }
    let _ = (Tier::A, Evidence::Queued);
    Ok(())
}
