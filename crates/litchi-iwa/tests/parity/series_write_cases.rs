//! Shared chart series style writes (wave W1-E, record 0375a): per-series
//! strokes, fills, data-symbol fills and data-symbol outlines through the
//! host editors and through the engine — `litchi_iwa_common::chart::series`
//! values, the leaf `drawing_native` encoders, the
//! `litchi_iwa_protos::chart_style_codec::property` write rules (the stroke
//! merge included) — on each series' effective style. The host route writes
//! the series styles of a chart it added (exclusive theme styles); the
//! payload write is the engine's, selection and exclusivity the families'.

use litchi_iwa::charts::{ChartSeriesStroke, ChartSeriesStrokePattern, ChartSeriesSymbolFill};
use litchi_iwa::shapes::ShapeFill;
use litchi_iwa::{keynote, numbers, pages};
use litchi_iwa_common::chart::series::{
    FillFamily, Stroke, StrokeFamily, StrokePattern, SymbolFill, SymbolFillFamily,
    SymbolOutlineFamily,
};
use litchi_iwa_common::color::{RgbColorSpace, Rgba};
use litchi_iwa_common::shape::fill::{Angle, Fill, Gradient};
use litchi_iwa_common::shape::stroke::Width;
use litchi_iwa_protos::chart_style_codec::{
    self as codec, Archive, Limits, Write,
    property::{self, PropertyWrite},
};
use litchi_iwa_style::drawing as plumbing;

use crate::chart_style_cases::{Components, Target, input, publish};
use crate::harness::{Case, Evidence, Format, LeafRun, TestResult, Tier};
use crate::series_style_cases::{Property, fill_at, flag, host_read, inherited, stroke_at};

const CHART_MESSAGE_TYPE: u32 = 5_021;

fn rgb(red: f32, green: f32, blue: f32) -> TestResult<Rgba> {
    Ok(Rgba::new(red, green, blue, 1.0, RgbColorSpace::Srgb)?)
}

/// Per-series values of one property, in the shared vocabulary.
#[derive(Clone, Debug)]
enum Values {
    Strokes(Vec<Option<Stroke>>),
    Fills(Vec<Fill>),
    SymbolFills(Vec<SymbolFill>),
    SymbolOutlines(Vec<Option<Stroke>>),
}

impl Values {
    const fn property(&self) -> Property {
        match self {
            Self::Strokes(_) => Property::Strokes,
            Self::Fills(_) => Property::Fills,
            Self::SymbolFills(_) => Property::SymbolFills,
            Self::SymbolOutlines(_) => Property::SymbolOutlines,
        }
    }

    fn len(&self) -> usize {
        match self {
            Self::Strokes(values) | Self::SymbolOutlines(values) => values.len(),
            Self::Fills(values) => values.len(),
            Self::SymbolFills(values) => values.len(),
        }
    }
}

fn to_host_stroke(stroke: Stroke) -> ChartSeriesStroke {
    let pattern = match stroke.pattern {
        StrokePattern::Solid => ChartSeriesStrokePattern::Solid,
        StrokePattern::MediumDash => ChartSeriesStrokePattern::MediumDash,
        StrokePattern::RoundedDash => ChartSeriesStrokePattern::RoundedDash,
    };
    ChartSeriesStroke::new(stroke.color, stroke.width, pattern)
}

fn to_host_fill(fill: &Fill) -> TestResult<ShapeFill> {
    Ok(match fill {
        Fill::None => ShapeFill::None,
        Fill::Solid(color) => ShapeFill::Solid(*color),
        Fill::Gradient(gradient) => ShapeFill::Gradient(gradient.clone()),
        Fill::Image(_) => return Err("image fills need the data-reference graph engine".into()),
    })
}

fn host_write(source: &[u8], target: Target, chart: u64, values: &Values) -> TestResult<Vec<u8>> {
    let strokes = |values: &[Option<Stroke>]| -> Vec<Option<ChartSeriesStroke>> {
        values
            .iter()
            .map(|value| value.map(to_host_stroke))
            .collect()
    };
    let fills = |values: &[Fill]| -> TestResult<Vec<ShapeFill>> {
        values.iter().map(to_host_fill).collect()
    };
    let symbol_fills = |values: &[SymbolFill]| -> TestResult<Vec<ChartSeriesSymbolFill>> {
        values
            .iter()
            .map(|value| {
                Ok(match value {
                    SymbolFill::SeriesFill => ChartSeriesSymbolFill::SeriesFill,
                    SymbolFill::SeriesStroke => ChartSeriesSymbolFill::SeriesStroke,
                    SymbolFill::Custom(fill) => ChartSeriesSymbolFill::Custom(to_host_fill(fill)?),
                })
            })
            .collect()
    };
    macro_rules! write {
        ($editor:expr, $($scope:expr,)? ; $strokes:ident, $fills:ident, $symbol_fills:ident, $outlines:ident) => {{
            let mut editor = $editor;
            match values {
                Values::Strokes(values) => editor.$strokes($($scope,)? chart, &strokes(values))?,
                Values::Fills(values) => editor.$fills($($scope,)? chart, &fills(values)?)?,
                Values::SymbolFills(values) => editor.$symbol_fills($($scope,)? chart, &symbol_fills(values)?)?,
                Values::SymbolOutlines(values) => editor.$outlines($($scope,)? chart, &strokes(values))?,
            }
            editor.to_bytes()?
        }};
    }
    Ok(match target {
        Target::Keynote { slide } => write!(
            keynote::Draft::from_bytes(source)?, slide, ;
            set_slide_chart_series_strokes, set_slide_chart_series_fills,
            set_slide_chart_series_symbol_fills, set_slide_chart_series_symbol_outlines
        ),
        Target::Numbers { sheet } => write!(
            numbers::Draft::from_bytes(source)?, sheet, ;
            set_sheet_chart_series_strokes, set_sheet_chart_series_fills,
            set_sheet_chart_series_symbol_fills, set_sheet_chart_series_symbol_outlines
        ),
        Target::Pages => write!(
            pages::Draft::from_bytes(source)?, ;
            set_body_chart_series_strokes, set_body_chart_series_fills,
            set_body_chart_series_symbol_fills, set_body_chart_series_symbol_outlines
        ),
    })
}

/// The engine write: for each series whose effective value differs, the
/// library write rule applied to its effective style.
fn engine_write(source: &[u8], chart: u64, values: &Values) -> TestResult<Vec<u8>> {
    let components = Components::read(source)?;
    let drawable = components.locate(chart, CHART_MESSAGE_TYPE)?;
    let references = codec::references(components.payload(chart, drawable)?, Limits::DEFAULT)?;
    let kind = values.property().kind();
    let mut replacements = Vec::new();
    for index in 0..values.len() {
        let style = references.series_style(index).ok_or("no series style")?;
        let location = components.locate(style, Archive::SeriesStyle.message_type())?;
        let payload = components.payload(style, location)?;
        // The effective value, the encoded replacement and its write rule.
        let encoded: Vec<u8>;
        let write: PropertyWrite<'_> = match values {
            Values::Strokes(strokes) => {
                let family = StrokeFamily::for_chart_kind(kind).ok_or("no stroke family")?;
                let current =
                    inherited(&components, style, |p| stroke_at(p, family.field()))?.flatten();
                if current == strokes[index] {
                    continue;
                }
                encoded = plumbing::encode_stroke(strokes[index].map(Stroke::to_shape_stroke));
                property::series_stroke(family.field(), Some(&encoded))
            },
            Values::SymbolOutlines(strokes) => {
                let family =
                    SymbolOutlineFamily::for_chart_kind(kind).ok_or("no outline family")?;
                let current =
                    inherited(&components, style, |p| stroke_at(p, family.field()))?.flatten();
                if current == strokes[index] {
                    continue;
                }
                encoded = plumbing::encode_stroke(strokes[index].map(Stroke::to_shape_stroke));
                property::series_message(family.field(), Some(&encoded))
            },
            Values::Fills(fills) => {
                let family = FillFamily::for_chart_kind(kind).ok_or("no fill family")?;
                let current = match inherited(&components, style, |p| fill_at(p, family.field()))? {
                    Some(fill) => Some(fill),
                    None => inherited(&components, style, |p| {
                        fill_at(p, FillFamily::DEFAULT_FIELD)
                    })?,
                };
                if current.as_ref() == Some(&fills[index]) {
                    continue;
                }
                encoded = plumbing::encode_fill(&fills[index]);
                property::series_message(family.field(), Some(&encoded))
            },
            Values::SymbolFills(fills) => {
                let family =
                    SymbolFillFamily::for_chart_kind(kind).ok_or("no symbol fill family")?;
                let current = inherited(&components, style, |p| {
                    let use_fill = match family.use_series_fill_field() {
                        Some(field) => flag(p, field)?,
                        None => None,
                    };
                    let use_stroke = flag(p, family.use_series_stroke_field())?;
                    Ok(family.read(use_fill, use_stroke, fill_at(p, family.fill_field())?)?)
                })?
                .unwrap_or_else(|| family.default_fill());
                if current == fills[index] {
                    continue;
                }
                let value = match &fills[index] {
                    SymbolFill::Custom(fill) => {
                        encoded = plumbing::encode_fill(fill);
                        property::SymbolFill::Custom(&encoded)
                    },
                    SymbolFill::SeriesFill => property::SymbolFill::SeriesFill,
                    SymbolFill::SeriesStroke => property::SymbolFill::SeriesStroke,
                };
                property::symbol_fill(
                    family.fill_field(),
                    family.use_series_fill_field(),
                    family.use_series_stroke_field(),
                    value,
                )
                .ok_or("the family cannot follow the series fill")?
            },
        };
        let writes: Vec<Write<'_>> = write.writes().to_vec();
        let rewritten = codec::rewrite(
            payload,
            write.archive(),
            &writes,
            write.empty(),
            Limits::DEFAULT,
        )?;
        if rewritten.changed() {
            replacements.push((
                style,
                Archive::SeriesStyle.message_type(),
                rewritten.into_bytes().into_owned(),
            ));
        }
    }
    publish(source, components, replacements)
}

/// Write `after` on a host-added chart, first styled with `before` when
/// given (so the write meets stored values: the stroke merge, in-place
/// replacement, removal).
fn write_case(
    name: &'static str,
    format: Format,
    tier: Tier,
    evidence: Evidence,
    before: Option<Values>,
    after: Values,
) -> TestResult {
    let mut input = input(format, after.property().kind())?;
    let (target, chart) = (input.target, input.chart);
    if let Some(before) = &before {
        input.fixture.bytes = host_write(&input.fixture.bytes, target, chart, before)?;
    }
    let after_ref = &after;
    let property = after.property();
    Case::new(name, input.format, tier, evidence, &input.fixture)
        .host(|source| host_write(source, target, chart, after_ref))
        .leaf(|source| Ok(LeafRun::output(engine_write(source, chart, after_ref)?)))
        .read("host series read", move |bytes| {
            host_read(bytes, target, chart, property)
        })
        .run()?;
    Ok(())
}

fn stroke(red: f32, width: f32, pattern: StrokePattern) -> TestResult<Stroke> {
    Ok(Stroke::new(
        rgb(red, 0.2, 0.3)?,
        Width::new(width)?,
        pattern,
    ))
}

fn strokes_a() -> TestResult<Vec<Option<Stroke>>> {
    Ok(vec![
        Some(stroke(0.8, 2.5, StrokePattern::MediumDash)?),
        None,
    ])
}

fn strokes_b() -> TestResult<Vec<Option<Stroke>>> {
    Ok(vec![
        Some(stroke(0.1, 1.0, StrokePattern::Solid)?),
        Some(stroke(0.6, 3.0, StrokePattern::RoundedDash)?),
    ])
}

fn fills_a() -> TestResult<Values> {
    Ok(Values::Fills(vec![
        Fill::Solid(rgb(0.1, 0.4, 0.9)?),
        Fill::Gradient(Gradient::linear(
            rgb(0.9, 0.9, 0.1)?,
            rgb(0.1, 0.6, 0.3)?,
            Angle::from_degrees(90.0)?,
        )),
    ]))
}

fn fills_b() -> TestResult<Values> {
    Ok(Values::Fills(vec![
        Fill::Gradient(Gradient::linear(
            rgb(0.2, 0.2, 0.2)?,
            rgb(0.8, 0.8, 0.8)?,
            Angle::from_degrees(45.0)?,
        )),
        Fill::Solid(rgb(0.7, 0.1, 0.1)?),
    ]))
}

fn symbol_fills_a() -> TestResult<Values> {
    Ok(Values::SymbolFills(vec![
        SymbolFill::SeriesFill,
        SymbolFill::Custom(Fill::Solid(rgb(0.2, 0.7, 0.2)?)),
    ]))
}

fn symbol_fills_b() -> TestResult<Values> {
    Ok(Values::SymbolFills(vec![
        SymbolFill::Custom(Fill::Solid(rgb(0.9, 0.5, 0.1)?)),
        SymbolFill::SeriesStroke,
    ]))
}

macro_rules! case {
    ($name:ident, $format:ident, $before:expr, $after:expr) => {
        #[test]
        fn $name() -> TestResult {
            write_case(
                stringify!($name),
                Format::$format,
                Tier::A,
                Evidence::Queued,
                $before,
                $after,
            )
        }
    };
}

case!(
    keynote_chart_series_strokes_write_parity,
    Keynote,
    None,
    Values::Strokes(strokes_a()?)
);
case!(
    numbers_chart_series_strokes_write_parity,
    Numbers,
    None,
    Values::Strokes(strokes_a()?)
);
case!(
    pages_chart_series_strokes_write_parity,
    Pages,
    None,
    Values::Strokes(strokes_a()?)
);
case!(
    keynote_chart_series_strokes_rewrite_parity,
    Keynote,
    Some(Values::Strokes(strokes_a()?)),
    Values::Strokes(strokes_b()?)
);
case!(
    numbers_chart_series_strokes_rewrite_parity,
    Numbers,
    Some(Values::Strokes(strokes_a()?)),
    Values::Strokes(strokes_b()?)
);
case!(
    pages_chart_series_strokes_rewrite_parity,
    Pages,
    Some(Values::Strokes(strokes_a()?)),
    Values::Strokes(strokes_b()?)
);
case!(
    keynote_chart_series_fills_write_parity,
    Keynote,
    None,
    fills_a()?
);
case!(
    numbers_chart_series_fills_write_parity,
    Numbers,
    None,
    fills_a()?
);
case!(
    pages_chart_series_fills_write_parity,
    Pages,
    None,
    fills_a()?
);
case!(
    keynote_chart_series_fills_rewrite_parity,
    Keynote,
    Some(fills_a()?),
    fills_b()?
);
case!(
    numbers_chart_series_fills_rewrite_parity,
    Numbers,
    Some(fills_a()?),
    fills_b()?
);
case!(
    pages_chart_series_fills_rewrite_parity,
    Pages,
    Some(fills_a()?),
    fills_b()?
);
case!(
    keynote_chart_series_symbol_fills_write_parity,
    Keynote,
    None,
    symbol_fills_a()?
);
case!(
    numbers_chart_series_symbol_fills_write_parity,
    Numbers,
    None,
    symbol_fills_a()?
);
case!(
    pages_chart_series_symbol_fills_write_parity,
    Pages,
    None,
    symbol_fills_a()?
);
case!(
    keynote_chart_series_symbol_fills_rewrite_parity,
    Keynote,
    Some(symbol_fills_a()?),
    symbol_fills_b()?
);
case!(
    numbers_chart_series_symbol_fills_rewrite_parity,
    Numbers,
    Some(symbol_fills_a()?),
    symbol_fills_b()?
);
case!(
    pages_chart_series_symbol_fills_rewrite_parity,
    Pages,
    Some(symbol_fills_a()?),
    symbol_fills_b()?
);
case!(
    keynote_chart_series_symbol_outlines_write_parity,
    Keynote,
    None,
    Values::SymbolOutlines(strokes_a()?)
);
case!(
    numbers_chart_series_symbol_outlines_write_parity,
    Numbers,
    None,
    Values::SymbolOutlines(strokes_a()?)
);
case!(
    pages_chart_series_symbol_outlines_write_parity,
    Pages,
    None,
    Values::SymbolOutlines(strokes_a()?)
);
case!(
    keynote_chart_series_symbol_outlines_rewrite_parity,
    Keynote,
    Some(Values::SymbolOutlines(strokes_a()?)),
    Values::SymbolOutlines(strokes_b()?)
);
case!(
    numbers_chart_series_symbol_outlines_rewrite_parity,
    Numbers,
    Some(Values::SymbolOutlines(strokes_a()?)),
    Values::SymbolOutlines(strokes_b()?)
);
case!(
    pages_chart_series_symbol_outlines_rewrite_parity,
    Pages,
    Some(Values::SymbolOutlines(strokes_a()?)),
    Values::SymbolOutlines(strokes_b()?)
);

/// The leaf route a family runs: admit the source package, the engine
/// write, admit the candidate (strict reopen).
#[inline(never)]
fn leaf_series_write_route(
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
fn host_series_write_route(
    source: &[u8],
    target: Target,
    chart: u64,
    values: &Values,
) -> TestResult<Vec<u8>> {
    host_write(source, target, chart, values)
}

/// ABBA cost probe of the series-style writes with source and candidate
/// admission (record 0375a §5); toggle
/// `*series_write_cases::host_series_write_route*` or
/// `*series_write_cases::leaf_series_write_route*`.
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn series_style_write_cost_abba() -> TestResult {
    for format in [Format::Keynote, Format::Numbers, Format::Pages] {
        for values in [Values::Strokes(strokes_a()?), fills_a()?, symbol_fills_a()?] {
            let input = input(format, values.property().kind())?;
            let (target, chart) = (input.target, input.chart);
            let source = input.fixture.bytes.as_slice();
            std::hint::black_box(host_series_write_route(source, target, chart, &values)?);
            std::hint::black_box(leaf_series_write_route(format, source, chart, &values)?);
            std::hint::black_box(leaf_series_write_route(format, source, chart, &values)?);
            std::hint::black_box(host_series_write_route(source, target, chart, &values)?);
        }
    }
    Ok(())
}
