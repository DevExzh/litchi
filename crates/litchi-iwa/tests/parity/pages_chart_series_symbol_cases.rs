//! Pages body-chart series-symbol facade parity (W2-B).
//!
//! The fixture is native Pages with one host-created Line2d chart.  The host
//! route remains the differential oracle; the leaf route exercises only the
//! selector-first `litchi_pages::Package` facade.

use litchi_iwa::charts::{ChartSeriesSymbol, ChartSeriesSymbolShape, ChartSeriesSymbolSize};
use litchi_iwa::pages;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};
use litchi_iwa_common::chart::ChartData;
use litchi_iwa_common::chart::kind::Kind;
use litchi_pages::chart::series::{Property, Symbol, SymbolShape, SymbolSize};
use litchi_pages::{BodyChartSelector, Package};

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::leaf_bytes;

const NATIVE: &str = "iwork/pages/chart-data-native.pages";

const POSITION: DrawablePoint = DrawablePoint { x: 120.0, y: 140.0 };
const SIZE: DrawableSize = DrawableSize {
    width: 360.0,
    height: 240.0,
};

type SymbolSpec = Option<(SymbolShape, Option<f32>)>;

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
    editor.add_body_chart(0, Kind::Line2d, data()?, POSITION, SIZE)?;
    Ok(Fixture {
        path: "chart-data-native.pages + host-added Line2d chart",
        bytes: editor.to_bytes()?,
    })
}

fn host_id(editor: &pages::Draft) -> TestResult<u64> {
    editor
        .body_charts()?
        .first()
        .map(|chart| chart.drawable_object_id)
        .ok_or_else(|| "host-added Pages chart missing".into())
}

fn host_symbols(source: &[u8]) -> TestResult<Vec<Option<ChartSeriesSymbol>>> {
    let editor = pages::Draft::from_bytes(source)?;
    Ok(editor.body_chart_series_symbols(host_id(&editor)?)?)
}

fn leaf_symbols(source: &[u8]) -> TestResult<Vec<Option<Symbol>>> {
    let package = Package::from_bytes(source)?;
    match package.body_chart_series(BodyChartSelector::index(0))? {
        Property::Symbols(values) => Ok(values),
        _ => Err("unexpected Pages series property".into()),
    }
}

fn to_host(values: &[SymbolSpec]) -> TestResult<Vec<Option<ChartSeriesSymbol>>> {
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

fn to_leaf(values: &[SymbolSpec]) -> TestResult<Vec<Option<Symbol>>> {
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

fn host_route(source: &[u8], values: &[SymbolSpec]) -> TestResult<Vec<u8>> {
    let mut editor = pages::Draft::from_bytes(source)?;
    let id = host_id(&editor)?;
    editor.set_body_chart_series_symbols(id, &to_host(values)?)?;
    Ok(editor.to_bytes()?)
}

fn leaf_route(source: &[u8], values: &[SymbolSpec]) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let mut edit = package.edit_body_chart_series(BodyChartSelector::index(0))?;
    edit = edit.set(Property::Symbols(to_leaf(values)?))?;
    let commit = edit.commit()?;
    let restored = commit
        .package()
        .apply_body_chart_series(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

fn direct_read_agrees(bytes: &[u8]) -> TestResult<String> {
    let host = host_symbols(bytes)?;
    let host = host
        .iter()
        .map(|value| {
            value.map(|value| {
                let shape = match value.shape {
                    ChartSeriesSymbolShape::Circle => SymbolShape::Circle,
                    ChartSeriesSymbolShape::Triangle => SymbolShape::Triangle,
                    ChartSeriesSymbolShape::InvertedTriangle => SymbolShape::InvertedTriangle,
                    ChartSeriesSymbolShape::Diamond => SymbolShape::Diamond,
                    ChartSeriesSymbolShape::Square => SymbolShape::Square,
                };
                let size = value
                    .size
                    .map(|size| SymbolSize::new(size.points()))
                    .transpose()
                    .expect("host point size is finite and positive");
                Symbol { shape, size }
            })
        })
        .collect::<Vec<_>>();
    let leaf = leaf_symbols(bytes)?;
    if host != leaf {
        return Err(format!("Pages host/leaf symbol reads differ: {host:?} != {leaf:?}").into());
    }
    Ok(format!("{leaf:?}"))
}

fn run_case(
    name: &'static str,
    fixture: &Fixture,
    values: &'static [SymbolSpec],
    tier: Tier,
    evidence: Evidence,
) -> TestResult {
    Case::new(name, Format::Pages, tier, evidence, fixture)
        .host(|source| host_route(source, values))
        .leaf(|source| leaf_route(source, values))
        .read("direct host-vs-leaf symbol read", direct_read_agrees)
        .leaf_structural_metadata()
        .run()?;
    Ok(())
}

const DIAMONDS: [SymbolSpec; 2] = [
    Some((SymbolShape::Diamond, Some(9.0))),
    Some((SymbolShape::Square, None)),
];

const SQUARE_HIDDEN: [SymbolSpec; 2] = [Some((SymbolShape::Square, Some(12.0))), None];

const NO_SYMBOLS: [SymbolSpec; 2] = [None, None];

#[test]
fn pages_body_chart_series_symbols_allocate_parity() -> TestResult {
    run_case(
        "pages_body_chart_series_symbols_allocate_parity",
        &input()?,
        &DIAMONDS,
        Tier::B,
        Evidence::Queued,
    )
}

#[test]
fn pages_body_chart_series_symbols_patch_parity() -> TestResult {
    let source = input()?;
    let seeded = Fixture {
        path: "host-seeded Pages Line2d symbols",
        bytes: host_route(&source.bytes, &DIAMONDS)?,
    };
    run_case(
        "pages_body_chart_series_symbols_patch_parity",
        &seeded,
        &SQUARE_HIDDEN,
        Tier::A,
        Evidence::Queued,
    )
}

#[test]
fn pages_body_chart_series_symbols_remove_parity() -> TestResult {
    let source = input()?;
    let seeded = Fixture {
        path: "host-seeded Pages Line2d symbols",
        bytes: host_route(&source.bytes, &DIAMONDS)?,
    };
    run_case(
        "pages_body_chart_series_symbols_remove_parity",
        &seeded,
        &NO_SYMBOLS,
        Tier::A,
        Evidence::Queued,
    )
}

#[test]
fn pages_body_chart_series_symbols_read_parity() -> TestResult {
    let source = input()?;
    let seeded = Fixture {
        path: "host-seeded Pages Line2d symbols",
        bytes: host_route(&source.bytes, &DIAMONDS)?,
    };
    let before = seeded.bytes.clone();
    Case::new(
        "pages_body_chart_series_symbols_read_parity",
        Format::Pages,
        Tier::A,
        Evidence::Queued,
        &seeded,
    )
    .host(|source| Ok(source.to_vec()))
    .leaf(|source| {
        let package = Package::from_bytes(source)?;
        Ok(LeafRun::output(leaf_bytes!(&package)))
    })
    .read("direct host-vs-leaf symbol read", direct_read_agrees)
    .read_only()
    .run()?;
    if seeded.bytes != before {
        return Err("read fixture changed before parity run".into());
    }
    Ok(())
}
