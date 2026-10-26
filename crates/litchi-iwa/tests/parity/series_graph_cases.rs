//! The style-variation graph engine of per-series chart overrides (wave
//! W1-E, record 0375a): connection lines, wedge explosions and label
//! distances through the host editors and through
//! `litchi_iwa_style::series::set_values` over a `litchi_iwa_style::graph`
//! workspace of the same input — override allocation (styled overrides in
//! the chart component, unstyled ones in a new object container),
//! stylesheet registration, removal, container release, the package
//! metadata registry and the identifier watermark. The leaf route
//! republishes the workspace's changes through `litchi_iwa_archive`'s
//! exact-source reassembly; selection belongs to the families.

use litchi_iwa::charts::{
    ChartPieLabelDistance, ChartPieWedgeExplosion, ChartSeriesConnectionLine, ChartSeriesSymbol,
    ChartSeriesSymbolShape, ChartSeriesSymbolSize, ChartSeriesTrendline,
    ChartSeriesTrendlineMovingAveragePeriod, ChartSeriesTrendlinePolynomialOrder,
};
use litchi_iwa::{keynote, numbers, pages};
use litchi_iwa_archive::Limits as ArchiveLimits;
use litchi_iwa_archive::iwa::{Limits as CoreLimits, SnappyLimits, SnappyStream};
use litchi_iwa_archive::unstable::{EntryEdit, EntryInsertion};
use litchi_iwa_common::chart::kind::Kind;
use litchi_iwa_common::chart::pie::{LabelDistance, WedgeExplosion};
use litchi_iwa_common::chart::series::{
    ConnectionLine, ConnectionLineFamily, MovingAveragePeriod, PolynomialOrder, Symbol,
    SymbolFamily, SymbolShape, SymbolSize, Trendline,
};
use litchi_iwa_style::graph::{Changes, Workspace};
use litchi_iwa_style::series::{
    self,
    property::{ConnectionLines, LabelDistances, Trendlines, WedgeExplosions},
};

use crate::chart_style_cases::{Components, Target, input};
use crate::harness::{Case, Evidence, Format, LeafRun, TestResult, Tier};

/// One per-series property write.
#[derive(Clone, Debug)]
enum Values {
    Lines(Kind, Vec<ConnectionLine>),
    /// Wedge explosions, in percent.
    Wedges(Vec<f32>),
    /// Label distances, in percent.
    Labels(Vec<f32>),
    /// Data symbols: `None` hides; `(shape, size in points)`.
    Symbols(Kind, Vec<Option<(SymbolShape, Option<f32>)>>),
    Trendlines(Vec<Trend>),
}

/// A trendline configuration, spelled once for both vocabularies.
#[derive(Clone, Copy, Debug)]
enum Trend {
    None,
    /// Linear, named "Fit" in the legend, with its equation.
    NamedLinear,
    /// Third-order polynomial with R².
    Polynomial,
    /// Four-point moving average.
    MovingAverage,
}

impl Trend {
    fn host(self) -> TestResult<ChartSeriesTrendline> {
        Ok(match self {
            Self::None => ChartSeriesTrendline::none(),
            Self::NamedLinear => ChartSeriesTrendline::linear()
                .with_legend_name("Fit")?
                .with_legend_visibility(true)?
                .with_equation_visibility(true)?,
            Self::Polynomial => {
                ChartSeriesTrendline::polynomial(ChartSeriesTrendlinePolynomialOrder::new(3)?)
                    .with_r_squared_visibility(true)?
            },
            Self::MovingAverage => ChartSeriesTrendline::moving_average(
                ChartSeriesTrendlineMovingAveragePeriod::new(4)?,
            ),
        })
    }

    fn common(self) -> TestResult<Trendline> {
        Ok(match self {
            Self::None => Trendline::none(),
            Self::NamedLinear => Trendline::linear()
                .with_legend_name("Fit")?
                .with_legend_visibility(true)?
                .with_equation_visibility(true)?,
            Self::Polynomial => {
                Trendline::polynomial(PolynomialOrder::new(3)?).with_r_squared_visibility(true)?
            },
            Self::MovingAverage => Trendline::moving_average(MovingAveragePeriod::new(4)?),
        })
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

impl Values {
    const fn kind(&self) -> Kind {
        match self {
            Self::Lines(kind, _) | Self::Symbols(kind, _) => *kind,
            Self::Wedges(_) | Self::Labels(_) => Kind::Pie2d,
            Self::Trendlines(_) => Kind::Scatter2d,
        }
    }
}

const fn host_line(value: ConnectionLine) -> ChartSeriesConnectionLine {
    match value {
        ConnectionLine::Hidden => ChartSeriesConnectionLine::Hidden,
        ConnectionLine::Straight => ChartSeriesConnectionLine::Straight,
        ConnectionLine::Curved => ChartSeriesConnectionLine::Curved,
    }
}

fn host_write(source: &[u8], target: Target, chart: u64, values: &Values) -> TestResult<Vec<u8>> {
    let lines = |values: &[ConnectionLine]| -> Vec<ChartSeriesConnectionLine> {
        values.iter().copied().map(host_line).collect()
    };
    let wedges = |values: &[f32]| -> TestResult<Vec<ChartPieWedgeExplosion>> {
        Ok(values
            .iter()
            .map(|percent| ChartPieWedgeExplosion::from_percent(*percent))
            .collect::<Result<_, _>>()?)
    };
    let labels = |values: &[f32]| -> TestResult<Vec<ChartPieLabelDistance>> {
        Ok(values
            .iter()
            .map(|percent| ChartPieLabelDistance::from_percent(*percent))
            .collect::<Result<_, _>>()?)
    };
    let symbols = |values: &[Option<(SymbolShape, Option<f32>)>]| -> TestResult<Vec<Option<ChartSeriesSymbol>>> {
        values
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
    };
    let trends = |values: &[Trend]| -> TestResult<Vec<ChartSeriesTrendline>> {
        values.iter().map(|trend| trend.host()).collect()
    };
    macro_rules! write {
        ($editor:expr, $($scope:expr,)? ; $lines:ident, $wedges:ident, $labels:ident, $symbols:ident, $trends:ident) => {{
            let mut editor = $editor;
            match values {
                Values::Lines(_, values) => editor.$lines($($scope,)? chart, &lines(values))?,
                Values::Wedges(values) => editor.$wedges($($scope,)? chart, &wedges(values)?)?,
                Values::Labels(values) => editor.$labels($($scope,)? chart, &labels(values)?)?,
                Values::Symbols(_, values) => editor.$symbols($($scope,)? chart, &symbols(values)?)?,
                Values::Trendlines(values) => editor.$trends($($scope,)? chart, &trends(values)?)?,
            }
            editor.to_bytes()?
        }};
    }
    Ok(match target {
        Target::Keynote { slide } => write!(
            keynote::Draft::from_bytes(source)?, slide, ;
            set_slide_chart_series_connection_lines, set_slide_chart_pie_wedge_explosions,
            set_slide_chart_pie_label_distances, set_slide_chart_series_symbols,
            set_slide_chart_series_trendlines
        ),
        Target::Numbers { sheet } => write!(
            numbers::Draft::from_bytes(source)?, sheet, ;
            set_sheet_chart_series_connection_lines, set_sheet_chart_pie_wedge_explosions,
            set_sheet_chart_pie_label_distances, set_sheet_chart_series_symbols,
            set_sheet_chart_series_trendlines
        ),
        Target::Pages => write!(
            pages::Draft::from_bytes(source)?, ;
            set_body_chart_series_connection_lines, set_body_chart_pie_wedge_explosions,
            set_body_chart_pie_label_distances, set_body_chart_series_symbols,
            set_body_chart_series_trendlines
        ),
    })
}

/// The host's read of the same property, as comparable text.
fn host_read(bytes: &[u8], target: Target, chart: u64, values: &Values) -> TestResult<String> {
    macro_rules! read {
        ($editor:expr, $($scope:expr,)? ; $lines:ident, $wedges:ident, $labels:ident, $symbols:ident, $trends:ident) => {{
            let editor = $editor;
            match values {
                Values::Lines(..) => format!("{:?}", editor.$lines($($scope,)? chart)?),
                Values::Wedges(_) => format!("{:?}", editor.$wedges($($scope,)? chart)?),
                Values::Labels(_) => format!("{:?}", editor.$labels($($scope,)? chart)?),
                Values::Symbols(..) => format!("{:?}", editor.$symbols($($scope,)? chart)?),
                Values::Trendlines(_) => format!("{:?}", editor.$trends($($scope,)? chart)?),
            }
        }};
    }
    Ok(match target {
        Target::Keynote { slide } => read!(
            keynote::Draft::from_bytes(bytes)?, slide, ;
            slide_chart_series_connection_lines, slide_chart_pie_wedge_explosions,
            slide_chart_pie_label_distances, slide_chart_series_symbols,
            slide_chart_series_trendlines
        ),
        Target::Numbers { sheet } => read!(
            numbers::Draft::from_bytes(bytes)?, sheet, ;
            sheet_chart_series_connection_lines, sheet_chart_pie_wedge_explosions,
            sheet_chart_pie_label_distances, sheet_chart_series_symbols,
            sheet_chart_series_trendlines
        ),
        Target::Pages => read!(
            pages::Draft::from_bytes(bytes)?, ;
            body_chart_series_connection_lines, body_chart_pie_wedge_explosions,
            body_chart_pie_label_distances, body_chart_series_symbols,
            body_chart_series_trendlines
        ),
    })
}

/// Republish a finished workspace's changes once through exact-source
/// reassembly.
pub(crate) fn publish_changes(
    source: &[u8],
    components: &Components,
    changes: Changes,
) -> TestResult<Vec<u8>> {
    if changes.is_empty() {
        return Ok(source.to_vec());
    }
    let compress = |archive: &litchi_iwa_archive::iwa::Archive| -> TestResult<Vec<u8>> {
        let bytes = archive.to_bytes(CoreLimits::default())?;
        Ok(SnappyStream::compress(&bytes, SnappyLimits::default())?)
    };
    let edited = changes
        .edits
        .iter()
        .map(|(name, archive)| Ok((name.clone(), compress(archive)?)))
        .collect::<TestResult<Vec<_>>>()?;
    let names: Vec<&str> = components
        .catalog
        .iter()
        .map(|entry| entry.name())
        .collect();
    let inserted = changes
        .insertions
        .iter()
        .map(|insertion| {
            let position = match &insertion.before {
                Some(before) => Some(
                    names
                        .iter()
                        .position(|name| name == before)
                        .ok_or("an insertion anchor is missing")?,
                ),
                None => None,
            };
            Ok((
                insertion.name.clone(),
                compress(&insertion.archive)?,
                position,
            ))
        })
        .collect::<TestResult<Vec<_>>>()?;
    let edits: Vec<EntryEdit<'_>> = edited
        .iter()
        .map(|(name, data)| EntryEdit::new(name.as_str(), data.as_slice()))
        .collect();
    let insertions: Vec<EntryInsertion<'_>> = inserted
        .iter()
        .map(|(name, data, position)| match position {
            Some(position) => EntryInsertion::at(*position, name.as_str(), data.as_slice()),
            None => EntryInsertion::new(name.as_str(), data.as_slice()),
        })
        .collect();
    let deletions: Vec<&str> = changes.deletions.iter().map(String::as_str).collect();
    let prepared = components.catalog.prepare_reassembly_with_changes(
        &insertions,
        &edits,
        &deletions,
        ArchiveLimits::default(),
    )?;
    let requirements = prepared.execution_requirements();
    Ok(prepared.execute(requirements.exact_limits())?)
}

/// The engine write over a workspace of the source components.
fn engine_write(source: &[u8], chart: u64, values: &Values) -> TestResult<Vec<u8>> {
    let components = Components::read(source)?;
    let mut workspace = Workspace::new(&components.archives, CoreLimits::default())?;
    match values {
        Values::Lines(kind, lines) => {
            let family = ConnectionLineFamily::for_chart_kind(*kind).ok_or("no line family")?;
            series::set_values(&mut workspace, chart, &ConnectionLines(family), lines)?;
        },
        Values::Wedges(percents) => {
            let wedges = percents
                .iter()
                .map(|percent| WedgeExplosion::from_percent(*percent))
                .collect::<Result<Vec<_>, _>>()?;
            series::set_values(&mut workspace, chart, &WedgeExplosions, &wedges)?;
        },
        Values::Labels(percents) => {
            let labels = percents
                .iter()
                .map(|percent| LabelDistance::from_percent(*percent))
                .collect::<Result<Vec<_>, _>>()?;
            series::set_values(&mut workspace, chart, &LabelDistances, &labels)?;
        },
        Values::Symbols(kind, values) => {
            let family = SymbolFamily::for_chart_kind(*kind).ok_or("no symbol family")?;
            let symbols = values
                .iter()
                .map(|value| {
                    value
                        .map(|(shape, size)| {
                            Ok::<_, Box<dyn std::error::Error>>(Symbol {
                                shape,
                                size: size.map(SymbolSize::new).transpose()?,
                            })
                        })
                        .transpose()
                })
                .collect::<Result<Vec<_>, _>>()?;
            series::set_symbols(&mut workspace, chart, family, &symbols)?;
        },
        Values::Trendlines(values) => {
            let trends = values
                .iter()
                .map(|trend| trend.common())
                .collect::<TestResult<Vec<_>>>()?;
            series::set_values(&mut workspace, chart, &Trendlines, &trends)?;
        },
    }
    let changes = workspace.finish();
    publish_changes(source, &components, changes)
}

/// Write `after` on a host-added chart, first written with each of
/// `before` through the host (so the engine meets stored overrides: in-place
/// patches, removals, container release).
fn graph_case(
    name: &'static str,
    format: Format,
    tier: Tier,
    before: &[Values],
    after: Values,
) -> TestResult {
    let mut input = input(format, after.kind())?;
    let (target, chart) = (input.target, input.chart);
    for values in before {
        input.fixture.bytes = host_write(&input.fixture.bytes, target, chart, values)?;
    }
    let after_ref = &after;
    let mut case = Case::new(name, input.format, tier, Evidence::Queued, &input.fixture)
        .host(|source| host_write(source, target, chart, after_ref))
        .leaf(|source| Ok(LeafRun::output(engine_write(source, chart, after_ref)?)))
        .read("host series read", move |bytes| {
            host_read(bytes, target, chart, after_ref)
        });
    if tier == Tier::B {
        case = case.leaf_structural_metadata();
    }
    case.run()?;
    Ok(())
}

fn lines(kind: Kind, values: &[ConnectionLine]) -> Values {
    Values::Lines(kind, values.to_vec())
}

const CURVED_FIRST: [ConnectionLine; 2] = [ConnectionLine::Curved, ConnectionLine::Straight];
const ALL_CURVED: [ConnectionLine; 2] = [ConnectionLine::Curved, ConnectionLine::Curved];
const ALL_STRAIGHT: [ConnectionLine; 2] = [ConnectionLine::Straight, ConnectionLine::Straight];
const SCATTER: [ConnectionLine; 1] = [ConnectionLine::Curved];
const SCATTER_HIDDEN: [ConnectionLine; 1] = [ConnectionLine::Hidden];

fn symbols(kind: Kind, values: &[Option<(SymbolShape, Option<f32>)>]) -> Values {
    Values::Symbols(kind, values.to_vec())
}

const DIAMONDS: [Option<(SymbolShape, Option<f32>)>; 2] = [
    Some((SymbolShape::Diamond, Some(9.0))),
    Some((SymbolShape::Square, None)),
];
const SQUARE_HIDDEN: [Option<(SymbolShape, Option<f32>)>; 2] =
    [Some((SymbolShape::Square, Some(12.0))), None];
const NO_SYMBOLS: [Option<(SymbolShape, Option<f32>)>; 2] = [None, None];

macro_rules! graph_cases {
    ($($name:ident: $format:ident, $tier:ident, [$($before:expr),*], $after:expr;)*) => {$(
        #[test]
        fn $name() -> TestResult {
            graph_case(stringify!($name), Format::$format, Tier::$tier, &[$($before),*], $after)
        }
    )*};
}

graph_cases! {
    keynote_chart_series_connection_lines_allocate_parity: Keynote, B, [], lines(Kind::Line2d, &CURVED_FIRST);
    numbers_chart_series_connection_lines_allocate_parity: Numbers, B, [], lines(Kind::Line2d, &CURVED_FIRST);
    pages_chart_series_connection_lines_allocate_parity: Pages, B, [], lines(Kind::Line2d, &CURVED_FIRST);
    keynote_chart_series_connection_lines_patch_parity: Keynote, B, [lines(Kind::Line2d, &CURVED_FIRST)], lines(Kind::Line2d, &[ConnectionLine::Straight, ConnectionLine::Curved]);
    numbers_chart_series_connection_lines_remove_parity: Numbers, A, [lines(Kind::Line2d, &ALL_CURVED)], lines(Kind::Line2d, &ALL_STRAIGHT);
    pages_chart_series_connection_lines_remove_parity: Pages, A, [lines(Kind::Line2d, &ALL_CURVED)], lines(Kind::Line2d, &ALL_STRAIGHT);
    keynote_chart_series_connection_lines_remove_parity: Keynote, A, [lines(Kind::Line2d, &ALL_CURVED)], lines(Kind::Line2d, &ALL_STRAIGHT);
    keynote_chart_series_scatter_lines_allocate_parity: Keynote, B, [], lines(Kind::Scatter2d, &SCATTER);
    pages_chart_series_scatter_lines_allocate_parity: Pages, B, [], lines(Kind::Scatter2d, &SCATTER);
    numbers_chart_series_scatter_lines_remove_parity: Numbers, A, [lines(Kind::Scatter2d, &SCATTER)], lines(Kind::Scatter2d, &SCATTER_HIDDEN);
    pages_chart_series_scatter_lines_patch_parity: Pages, A, [lines(Kind::Scatter2d, &SCATTER)], lines(Kind::Scatter2d, &[ConnectionLine::Straight]);
    numbers_chart_series_radar_lines_allocate_parity: Numbers, B, [], lines(Kind::Radar2d, &ALL_CURVED);
    keynote_chart_pie_wedge_explosions_allocate_parity: Keynote, B, [], Values::Wedges(vec![25.0, 0.0, 10.0]);
    numbers_chart_pie_wedge_explosions_allocate_parity: Numbers, B, [], Values::Wedges(vec![25.0, 0.0, 10.0]);
    pages_chart_pie_wedge_explosions_allocate_parity: Pages, B, [], Values::Wedges(vec![25.0, 0.0, 10.0]);
    keynote_chart_pie_wedge_explosions_partial_release_parity: Keynote, A, [Values::Wedges(vec![25.0, 0.0, 10.0])], Values::Wedges(vec![0.0, 0.0, 10.0]);
    numbers_chart_pie_wedge_explosions_release_parity: Numbers, A, [Values::Wedges(vec![25.0, 0.0, 10.0])], Values::Wedges(vec![0.0, 0.0, 0.0]);
    pages_chart_pie_wedge_explosions_release_parity: Pages, A, [Values::Wedges(vec![25.0, 0.0, 10.0])], Values::Wedges(vec![0.0, 0.0, 0.0]);
    keynote_chart_pie_label_distances_allocate_parity: Keynote, B, [], Values::Labels(vec![120.0, 67.0, 40.0]);
    numbers_chart_pie_label_distances_patch_parity: Numbers, A, [Values::Wedges(vec![25.0, 0.0, 10.0])], Values::Labels(vec![120.0, 67.0, 40.0]);
    pages_chart_pie_label_distances_patch_parity: Pages, A, [Values::Wedges(vec![25.0, 0.0, 10.0])], Values::Labels(vec![120.0, 67.0, 40.0]);
    keynote_chart_series_symbols_allocate_parity: Keynote, B, [], symbols(Kind::Line2d, &DIAMONDS);
    numbers_chart_series_symbols_allocate_parity: Numbers, B, [], symbols(Kind::Line2d, &DIAMONDS);
    pages_chart_series_symbols_allocate_parity: Pages, B, [], symbols(Kind::Line2d, &DIAMONDS);
    keynote_chart_series_symbols_patch_parity: Keynote, A, [symbols(Kind::Line2d, &DIAMONDS)], symbols(Kind::Line2d, &SQUARE_HIDDEN);
    numbers_chart_series_symbols_remove_parity: Numbers, A, [symbols(Kind::Line2d, &DIAMONDS)], symbols(Kind::Line2d, &NO_SYMBOLS);
    pages_chart_series_symbols_patch_parity: Pages, A, [symbols(Kind::Line2d, &DIAMONDS)], symbols(Kind::Line2d, &SQUARE_HIDDEN);
    keynote_chart_series_trendlines_allocate_parity: Keynote, B, [], Values::Trendlines(vec![Trend::NamedLinear]);
    numbers_chart_series_trendlines_allocate_parity: Numbers, B, [], Values::Trendlines(vec![Trend::Polynomial]);
    pages_chart_series_trendlines_allocate_parity: Pages, B, [], Values::Trendlines(vec![Trend::MovingAverage]);
    keynote_chart_series_trendlines_patch_parity: Keynote, A, [Values::Trendlines(vec![Trend::NamedLinear])], Values::Trendlines(vec![Trend::Polynomial]);
    numbers_chart_series_trendlines_patch_parity: Numbers, A, [Values::Trendlines(vec![Trend::Polynomial])], Values::Trendlines(vec![Trend::MovingAverage]);
    pages_chart_series_trendlines_remove_parity: Pages, A, [Values::Trendlines(vec![Trend::NamedLinear])], Values::Trendlines(vec![Trend::None]);
}

/// The leaf route a family runs: admit the source package, the graph
/// engine write, admit the candidate (strict reopen).
#[inline(never)]
fn leaf_series_graph_route(
    format: Format,
    source: &[u8],
    chart: u64,
    values: &Values,
) -> TestResult<Vec<u8>> {
    crate::series_style_cases::leaf_reopen(format, source)?;
    let output = engine_write(source, chart, values)?;
    format.reopen(&output)?;
    Ok(output)
}

#[inline(never)]
fn host_series_graph_route(
    source: &[u8],
    target: Target,
    chart: u64,
    values: &Values,
) -> TestResult<Vec<u8>> {
    host_write(source, target, chart, values)
}

/// ABBA cost probe of the graph-engine writes with source and candidate
/// admission (record 0375a §5): a styled allocation, an unstyled
/// allocation (new object container) and a container release per format;
/// toggle `*series_graph_cases::host_series_graph_route*` or
/// `*series_graph_cases::leaf_series_graph_route*`.
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn series_graph_write_cost_abba() -> TestResult {
    for format in [Format::Keynote, Format::Numbers, Format::Pages] {
        let runs = [
            (None, lines(Kind::Line2d, &CURVED_FIRST)),
            (None, Values::Wedges(vec![25.0, 0.0, 10.0])),
            (
                Some(Values::Wedges(vec![25.0, 0.0, 10.0])),
                Values::Wedges(vec![0.0, 0.0, 0.0]),
            ),
        ];
        for (before, values) in runs {
            let mut input = input(format, values.kind())?;
            let (target, chart) = (input.target, input.chart);
            if let Some(before) = &before {
                input.fixture.bytes = host_write(&input.fixture.bytes, target, chart, before)?;
            }
            let source = input.fixture.bytes.as_slice();
            std::hint::black_box(host_series_graph_route(source, target, chart, &values)?);
            std::hint::black_box(leaf_series_graph_route(format, source, chart, &values)?);
            std::hint::black_box(leaf_series_graph_route(format, source, chart, &values)?);
            std::hint::black_box(host_series_graph_route(source, target, chart, &values)?);
        }
    }
    Ok(())
}
