//! Keynote slide-chart data-symbol facade parity (W2-C, record 0401).
//!
//! The shared series graph is covered by record 0375.  These cases exercise
//! the selector-first Keynote package boundary on the same host-authored
//! chart, including direct host-vs-leaf reads, all-series and one-series host
//! writes, exact inverse/no-op behavior, and terminal selector refusals.

use litchi_iwa::charts::{ChartSeriesSymbol, ChartSeriesSymbolShape, ChartSeriesSymbolSize};
use litchi_iwa::keynote;
use litchi_iwa_common::chart::kind::Kind;
use litchi_iwa_common::chart::series::{Symbol, SymbolShape, SymbolSize};
use litchi_keynote::chart::series::{Error, Property, Refusal};
use litchi_keynote::{ChartSelector, Package, SlideSelector};

use crate::chart_style_cases::{Input, input};
use crate::harness::{Case, Evidence, Format, LeafRun, TestResult, Tier};
use crate::leaf_bytes;

const SLIDE: usize = 0;
const LEAF_SLIDE: SlideSelector<'static> = SlideSelector::index(SLIDE);
const LEAF_CHART: ChartSelector<'static> = ChartSelector::index(1);

#[derive(Clone, Debug)]
struct Values(Vec<Option<(SymbolShape, Option<f32>)>>);

impl Values {
    fn leaf(&self) -> TestResult<Vec<Option<Symbol>>> {
        self.0
            .iter()
            .map(|value| {
                value
                    .map(|(shape, points)| {
                        Ok(Symbol {
                            shape,
                            size: points.map(SymbolSize::new).transpose()?,
                        })
                    })
                    .transpose()
            })
            .collect()
    }

    fn host(&self) -> TestResult<Vec<Option<ChartSeriesSymbol>>> {
        self.0
            .iter()
            .map(|value| {
                Ok(match value {
                    None => None,
                    Some((shape, None)) => Some(ChartSeriesSymbol::automatic(host_shape(*shape))),
                    Some((shape, Some(points))) => Some(ChartSeriesSymbol::sized(
                        host_shape(*shape),
                        ChartSeriesSymbolSize::new(*points)?,
                    )),
                })
            })
            .collect()
    }
}

const fn host_shape(shape: SymbolShape) -> ChartSeriesSymbolShape {
    match shape {
        SymbolShape::Circle => ChartSeriesSymbolShape::Circle,
        SymbolShape::Triangle => ChartSeriesSymbolShape::Triangle,
        SymbolShape::InvertedTriangle => ChartSeriesSymbolShape::InvertedTriangle,
        SymbolShape::Diamond => ChartSeriesSymbolShape::Diamond,
        SymbolShape::Square => ChartSeriesSymbolShape::Square,
    }
}

fn host_read(source: &[u8], chart: u64) -> TestResult<Vec<Option<Symbol>>> {
    let editor = keynote::Draft::from_bytes(source)?;
    editor
        .slide_chart_series_symbols(SLIDE, chart)?
        .into_iter()
        .map(|value| match value {
            None => Ok(None),
            Some(value) => Ok(Some(Symbol {
                shape: match value.shape {
                    ChartSeriesSymbolShape::Circle => SymbolShape::Circle,
                    ChartSeriesSymbolShape::Triangle => SymbolShape::Triangle,
                    ChartSeriesSymbolShape::InvertedTriangle => SymbolShape::InvertedTriangle,
                    ChartSeriesSymbolShape::Diamond => SymbolShape::Diamond,
                    ChartSeriesSymbolShape::Square => SymbolShape::Square,
                },
                size: value
                    .size
                    .map(|size| SymbolSize::new(size.points()))
                    .transpose()?,
            })),
        })
        .collect()
}

fn leaf_read(source: &[u8]) -> TestResult<Vec<Option<Symbol>>> {
    let package = Package::from_bytes(source)?;
    let Property::Symbols(values) = package.slide_chart_series_symbols(LEAF_SLIDE, LEAF_CHART)?
    else {
        return Err("unexpected chart-series property".into());
    };
    Ok(values)
}

fn host_write(
    source: &[u8],
    chart: u64,
    values: &Values,
    one: Option<usize>,
) -> TestResult<Vec<u8>> {
    let mut editor = keynote::Draft::from_bytes(source)?;
    let host_values = values.host()?;
    match one {
        None => editor.set_slide_chart_series_symbols(SLIDE, chart, &host_values)?,
        Some(index) => editor.set_slide_chart_series_symbol(
            SLIDE,
            chart,
            litchi_iwa::charts::Index::from_zero_based(index),
            host_values[index],
        )?,
    }
    Ok(editor.to_bytes()?)
}

fn leaf_write(source: &[u8], values: &Values, one: Option<usize>) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let mut symbols = match one {
        None => values.leaf()?,
        Some(_index) => leaf_read(source)?,
    };
    if let Some(index) = one {
        symbols[index] = values.leaf()?[index];
    }
    let mut edit = package.edit_slide_chart_series(LEAF_SLIDE, LEAF_CHART)?;
    edit.set(symbols)?;
    let commit = edit.commit()?;
    let restored = commit
        .package()
        .apply_slide_chart_series(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

fn direct_read_proof(source: &[u8], chart: u64) -> TestResult {
    let host = host_read(source, chart)?;
    let leaf = leaf_read(source)?;
    assert_eq!(host, leaf, "host and leaf readers must agree directly");
    Ok(())
}

fn input_with_values(values: &Values) -> TestResult<(Input, u64)> {
    let mut input = input(Format::Keynote, Kind::Line2d)?;
    let chart = input.chart;
    input.fixture.bytes = host_write(&input.fixture.bytes, chart, values, None)?;
    direct_read_proof(&input.fixture.bytes, chart)?;
    Ok((input, chart))
}

fn run_case(name: &'static str, values: Values, one: Option<usize>) -> TestResult {
    let input = input(Format::Keynote, Kind::Line2d)?;
    let chart = input.chart;
    let values_for_host = values.clone();
    let values_for_leaf = values.clone();
    Case::new(
        name,
        Format::Keynote,
        Tier::B,
        Evidence::Queued,
        &input.fixture,
    )
    .host(move |bytes| host_write(bytes, chart, &values_for_host, one))
    .leaf(move |bytes| leaf_write(bytes, &values_for_leaf, one))
    .read("host/leaf chart-series symbols", move |bytes| {
        host_read(bytes, chart)
    })
    .leaf_invalidates_root_previews()
    .leaf_structural_metadata()
    .run()?;
    Ok(())
}

#[test]
fn keynote_slide_chart_series_symbols_allocate_facade_parity() -> TestResult {
    let _tier = Tier::B;
    let _evidence = Evidence::Queued;
    run_case(
        "keynote_slide_chart_series_symbols_allocate_facade_parity",
        Values(vec![
            Some((SymbolShape::Diamond, Some(9.0))),
            Some((SymbolShape::Square, None)),
        ]),
        None,
    )
}

#[test]
fn keynote_slide_chart_series_symbol_allocate_facade_parity() -> TestResult {
    let _tier = Tier::B;
    let _evidence = Evidence::Queued;
    run_case(
        "keynote_slide_chart_series_symbol_allocate_facade_parity",
        Values(vec![Some((SymbolShape::Diamond, Some(11.0))), None]),
        Some(0),
    )
}

#[test]
fn keynote_slide_chart_series_symbols_read_facade_parity() -> TestResult {
    let _tier = Tier::A;
    let _evidence = Evidence::Queued;
    let values = Values(vec![
        Some((SymbolShape::Diamond, Some(9.0))),
        Some((SymbolShape::Square, None)),
    ]);
    let (input, chart) = input_with_values(&values)?;
    direct_read_proof(&input.fixture.bytes, chart)?;
    assert_eq!(host_read(&input.fixture.bytes, chart)?, values.leaf()?);
    Ok(())
}

#[test]
fn keynote_slide_chart_series_symbol_read_facade_parity() -> TestResult {
    let _tier = Tier::A;
    let _evidence = Evidence::Queued;
    let values = Values(vec![Some((SymbolShape::Triangle, Some(8.0))), None]);
    let (input, chart) = input_with_values(&values)?;
    direct_read_proof(&input.fixture.bytes, chart)?;
    assert_eq!(host_read(&input.fixture.bytes, chart)?, values.leaf()?);
    Ok(())
}

#[test]
fn keynote_slide_chart_series_noop_inverse_and_refusal_controls() -> TestResult {
    let values = Values(vec![
        Some((SymbolShape::Diamond, Some(9.0))),
        Some((SymbolShape::Square, None)),
    ]);
    let (input, _chart) = input_with_values(&values)?;
    let package = Package::from_bytes(&input.fixture.bytes)?;
    let current = values.leaf()?;
    let mut edit = package.edit_slide_chart_series(LEAF_SLIDE, LEAF_CHART)?;
    edit.set(current)?;
    let noop = edit.commit()?;
    assert!(noop.patch().is_noop());
    assert!(!noop.diagnostics().changed());
    assert_eq!(leaf_bytes!(noop.package()), input.fixture.bytes);

    let mut changed = package.edit_slide_chart_series(LEAF_SLIDE, LEAF_CHART)?;
    changed.set(Values(vec![None, None]).leaf()?)?;
    let changed = changed.commit()?;
    let restored = changed
        .package()
        .apply_slide_chart_series(&changed.patch().inverse())?;
    assert_eq!(leaf_bytes!(restored.package()), input.fixture.bytes);

    assert!(matches!(
        package.slide_chart_series_symbols(LEAF_SLIDE, ChartSelector::index(99)),
        Err(Error::Refused(Refusal::ChartPositionNotFound { .. }))
    ));
    assert!(matches!(
        package
            .edit_slide_chart_series(LEAF_SLIDE, LEAF_CHART)?
            .set(Vec::new()),
        Err(Error::Refused(Refusal::SeriesCount))
    ));
    Ok(())
}
