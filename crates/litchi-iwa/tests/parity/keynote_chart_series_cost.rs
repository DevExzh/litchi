//! Ignored release ABBA probes for the Keynote chart-series symbol facade
//! (W2-C, record 0401).
//!
//! The setup writes are outside the measured wrappers.  A coordinator may
//! capture one route at a time with the same fixture, release profile and
//! pinned CPU used by the other Keynote facade records:
//!
//! ```text
//! CARGO_PROFILE_RELEASE_DEBUG=1 cargo test --release -p litchi-iwa --test parity --no-run
//! taskset -c 3 valgrind --tool=callgrind target/release/deps/parity-<hash> --ignored --exact \
//!   keynote_chart_series_cost::keynote_chart_series_symbols_set_cost_abba --test-threads=1
//! ```

use std::hint::black_box;

use litchi_iwa::charts::{ChartSeriesSymbol, ChartSeriesSymbolShape, ChartSeriesSymbolSize, Index};
use litchi_iwa::keynote;
use litchi_iwa_common::chart::kind::Kind;
use litchi_iwa_common::chart::series::{Symbol, SymbolShape, SymbolSize};
use litchi_keynote::chart::series::Property;
use litchi_keynote::{ChartSelector, Package, SlideSelector};

use crate::chart_style_cases::input;
use crate::harness::{Format, TestResult};
use crate::leaf_bytes;

const SLIDE: usize = 0;
const CHART: ChartSelector<'static> = ChartSelector::index(1);
const LEAF_SLIDE: SlideSelector<'static> = SlideSelector::index(SLIDE);

fn host_symbols() -> Vec<Option<ChartSeriesSymbol>> {
    vec![
        Some(ChartSeriesSymbol::sized(
            ChartSeriesSymbolShape::Diamond,
            ChartSeriesSymbolSize::new(9.0).expect("finite positive size"),
        )),
        Some(ChartSeriesSymbol::automatic(ChartSeriesSymbolShape::Square)),
    ]
}

fn leaf_symbols() -> Vec<Option<Symbol>> {
    vec![
        Some(Symbol {
            shape: SymbolShape::Diamond,
            size: Some(SymbolSize::new(9.0).expect("finite positive size")),
        }),
        Some(Symbol {
            shape: SymbolShape::Square,
            size: None,
        }),
    ]
}

fn setup_source() -> TestResult<(Vec<u8>, u64)> {
    let input = input(Format::Keynote, Kind::Line2d)?;
    Ok((input.fixture.bytes, input.chart))
}

fn setup_populated_source() -> TestResult<(Vec<u8>, u64)> {
    let (source, chart) = setup_source()?;
    let mut editor = keynote::Draft::from_bytes(&source)?;
    editor.set_slide_chart_series_symbols(SLIDE, chart, &host_symbols())?;
    Ok((editor.to_bytes()?, chart))
}

#[inline(never)]
fn host_symbols_read_cost(source: &[u8], chart: u64) -> TestResult<Vec<Option<ChartSeriesSymbol>>> {
    let editor = keynote::Draft::from_bytes(source)?;
    Ok(editor.slide_chart_series_symbols(SLIDE, chart)?)
}

#[inline(never)]
fn leaf_symbols_read_cost(source: &[u8]) -> TestResult<Vec<Option<Symbol>>> {
    let package = Package::from_bytes(source)?;
    let Property::Symbols(values) = package.slide_chart_series(LEAF_SLIDE, CHART)? else {
        return Err("unexpected chart-series property".into());
    };
    Ok(values)
}

#[inline(never)]
fn host_symbol_read_cost(
    source: &[u8],
    chart: u64,
    index: usize,
) -> TestResult<Option<ChartSeriesSymbol>> {
    let editor = keynote::Draft::from_bytes(source)?;
    Ok(editor.slide_chart_series_symbol(SLIDE, chart, Index::from_zero_based(index))?)
}

#[inline(never)]
fn leaf_symbol_read_cost(source: &[u8], index: usize) -> TestResult<Option<Symbol>> {
    let package = Package::from_bytes(source)?;
    let Property::Symbols(values) = package.slide_chart_series(LEAF_SLIDE, CHART)? else {
        return Err("unexpected chart-series property".into());
    };
    values
        .get(index)
        .cloned()
        .ok_or_else(|| "leaf series index was not present".into())
}

#[inline(never)]
fn host_symbols_set_cost(source: &[u8], chart: u64) -> TestResult<Vec<u8>> {
    let mut editor = keynote::Draft::from_bytes(source)?;
    editor.set_slide_chart_series_symbols(SLIDE, chart, &host_symbols())?;
    Ok(editor.to_bytes()?)
}

#[inline(never)]
fn leaf_symbols_set_cost(source: &[u8]) -> TestResult<Vec<u8>> {
    let package = Package::from_bytes(source)?;
    let mut edit = package.edit_slide_chart_series(LEAF_SLIDE, CHART)?;
    edit.set(leaf_symbols())?;
    Ok(leaf_bytes!(edit.commit()?.package()))
}

#[inline(never)]
fn host_symbol_set_cost(source: &[u8], chart: u64) -> TestResult<Vec<u8>> {
    let mut editor = keynote::Draft::from_bytes(source)?;
    editor.set_slide_chart_series_symbol(
        SLIDE,
        chart,
        Index::from_zero_based(0),
        host_symbols()[0],
    )?;
    Ok(editor.to_bytes()?)
}

#[inline(never)]
fn leaf_symbol_set_cost(source: &[u8]) -> TestResult<Vec<u8>> {
    let package = Package::from_bytes(source)?;
    let Property::Symbols(mut values) = package.slide_chart_series(LEAF_SLIDE, CHART)? else {
        return Err("unexpected chart-series property".into());
    };
    values[0] = leaf_symbols()[0];
    let mut edit = package.edit_slide_chart_series(LEAF_SLIDE, CHART)?;
    edit.set(values)?;
    Ok(leaf_bytes!(edit.commit()?.package()))
}

#[test]
#[ignore = "coordinator release ABBA cost probe"]
fn keynote_chart_series_symbols_read_cost_abba() -> TestResult {
    let (source, chart) = setup_populated_source()?;
    black_box(host_symbols_read_cost(&source, chart)?);
    black_box(leaf_symbols_read_cost(&source)?);
    black_box(leaf_symbols_read_cost(&source)?);
    black_box(host_symbols_read_cost(&source, chart)?);
    Ok(())
}

#[test]
#[ignore = "coordinator release ABBA cost probe"]
fn keynote_chart_series_symbols_set_cost_abba() -> TestResult {
    let (source, chart) = setup_source()?;
    black_box(host_symbols_set_cost(&source, chart)?);
    black_box(leaf_symbols_set_cost(&source)?);
    black_box(leaf_symbols_set_cost(&source)?);
    black_box(host_symbols_set_cost(&source, chart)?);
    Ok(())
}

#[test]
#[ignore = "coordinator release ABBA cost probe"]
fn keynote_chart_series_symbol_set_cost_abba() -> TestResult {
    let (source, chart) = setup_source()?;
    black_box(host_symbol_set_cost(&source, chart)?);
    black_box(leaf_symbol_set_cost(&source)?);
    black_box(leaf_symbol_set_cost(&source)?);
    black_box(host_symbol_set_cost(&source, chart)?);
    Ok(())
}

#[test]
#[ignore = "coordinator release ABBA cost probe"]
fn keynote_chart_series_symbol_read_cost_abba() -> TestResult {
    let (source, chart) = setup_populated_source()?;
    black_box(host_symbol_read_cost(&source, chart, 0)?);
    black_box(leaf_symbol_read_cost(&source, 0)?);
    black_box(leaf_symbol_read_cost(&source, 0)?);
    black_box(host_symbol_read_cost(&source, chart, 0)?);
    Ok(())
}
