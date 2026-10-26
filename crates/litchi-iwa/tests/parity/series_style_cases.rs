//! Shared chart series style and shape fill reads (wave W1-E, record 0375):
//! the host editors' per-series strokes, fills, connection lines, data
//! symbols, symbol fills, symbol outlines and trendlines, and a shape's
//! image fill, against the engine reads — `litchi_iwa_common::chart::series`
//! and `shape::{fill, native}` over `litchi_iwa_protos::{chart_style_codec,
//! shape_style_codec::native}`, joined by the leaf `drawing_native` field
//! pass — on a chart (or shape) the host added and styled in a native-saved
//! fixture of each format. Read-only cases: both routes leave the bytes as
//! they are and the reads carry the comparison.

use litchi_iwa::charts::{
    ChartSeriesConnectionLine, ChartSeriesStroke, ChartSeriesStrokePattern, ChartSeriesSymbol,
    ChartSeriesSymbolFill, ChartSeriesSymbolShape, ChartSeriesSymbolSize, ChartSeriesTrendline,
    ChartSeriesTrendlinePolynomialOrder,
};
use litchi_iwa::shapes::{
    DrawablePoint, DrawableSize, RgbColorSpace, RgbaColor, ShapeFill, ShapeImageDataIdentifier,
    ShapeImageFill, ShapeImageFillTechnique,
};
use litchi_iwa::{keynote, numbers, pages};
use litchi_iwa_common::chart::kind::Kind;
use litchi_iwa_common::chart::series::{
    ConnectionLine, ConnectionLineFamily, FillFamily, PolynomialOrder, Stroke, StrokeFamily,
    StrokePattern, Symbol, SymbolFamily, SymbolFill, SymbolFillFamily, SymbolOutlineFamily,
    SymbolShape, SymbolSize, Trendline, trendline_field,
};
use litchi_iwa_common::shape::fill::{Fill, Image, ImageData, ImageTechnique};
use litchi_iwa_common::shape::path::Preset;
use litchi_iwa_common::shape::stroke::Width;
use litchi_iwa_protos::chart_style_codec::{self as codec, Archive, Limits, Properties};
use litchi_iwa_protos::shape_style_codec::{self as shape_codec, native as stored};
use litchi_keynote::drawing_native as plumbing;

use crate::chart_style_cases::{Components, Target, input};
use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier};
use crate::leaf_bytes;

const CHART_MESSAGE_TYPE: u32 = 5_021;
const SHAPE_INFO_MESSAGE_TYPE: u32 = 2_011;
const SHAPE_STYLE_MESSAGE_TYPE: u32 = 2_025;
const MAX_DEPTH: usize = 64;

fn color(red: f32, green: f32, blue: f32) -> TestResult<RgbaColor> {
    Ok(RgbaColor::new(red, green, blue, 1.0, RgbColorSpace::Srgb)?)
}

/// The series property a case styles and reads.
#[derive(Clone, Copy, Debug)]
pub(crate) enum Property {
    Strokes,
    Fills,
    ConnectionLines,
    Symbols,
    SymbolFills,
    SymbolOutlines,
    Trendlines,
}

impl Property {
    pub(crate) const fn kind(self) -> Kind {
        match self {
            Self::Fills => Kind::Column2d,
            _ => Kind::Line2d,
        }
    }
}

fn host_stroke(width: f32, pattern: ChartSeriesStrokePattern) -> TestResult<ChartSeriesStroke> {
    Ok(ChartSeriesStroke::new(
        color(0.8, 0.1, 0.2)?,
        Width::new(width)?,
        pattern,
    ))
}

/// Style the host chart with the case's values (the read-only input).
fn host_style(
    source: &[u8],
    target: Target,
    chart: u64,
    property: Property,
) -> TestResult<Vec<u8>> {
    let strokes = vec![
        Some(host_stroke(2.5, ChartSeriesStrokePattern::MediumDash)?),
        None,
    ];
    let fills = vec![
        ShapeFill::Solid(color(0.1, 0.4, 0.9)?),
        ShapeFill::Gradient(litchi_iwa_common::shape::fill::Gradient::linear(
            color(0.9, 0.9, 0.1)?,
            color(0.1, 0.6, 0.3)?,
            litchi_iwa_common::shape::fill::Angle::from_degrees(90.0)?,
        )),
    ];
    let lines = vec![
        ChartSeriesConnectionLine::Curved,
        ChartSeriesConnectionLine::Straight,
    ];
    let symbols = vec![
        Some(ChartSeriesSymbol::sized(
            ChartSeriesSymbolShape::Diamond,
            ChartSeriesSymbolSize::new(9.0)?,
        )),
        None,
    ];
    let symbol_fills = vec![
        ChartSeriesSymbolFill::SeriesFill,
        ChartSeriesSymbolFill::Custom(ShapeFill::Solid(color(0.2, 0.7, 0.2)?)),
    ];
    let outlines = vec![
        Some(host_stroke(1.5, ChartSeriesStrokePattern::RoundedDash)?),
        None,
    ];
    let trendlines = vec![
        ChartSeriesTrendline::linear()
            .with_legend_name("Fit")?
            .with_equation_visibility(true)?,
        ChartSeriesTrendline::polynomial(ChartSeriesTrendlinePolynomialOrder::new(3)?)
            .with_r_squared_visibility(true)?,
    ];
    macro_rules! style {
        ($editor:expr, $($scope:expr,)? ; $strokes:ident, $fills:ident, $lines:ident, $symbols:ident, $symbol_fills:ident, $outlines:ident, $trendlines:ident) => {{
            let mut editor = $editor;
            match property {
                Property::Strokes => editor.$strokes($($scope,)? chart, &strokes)?,
                Property::Fills => editor.$fills($($scope,)? chart, &fills)?,
                Property::ConnectionLines => editor.$lines($($scope,)? chart, &lines)?,
                Property::Symbols => editor.$symbols($($scope,)? chart, &symbols)?,
                Property::SymbolFills => editor.$symbol_fills($($scope,)? chart, &symbol_fills)?,
                Property::SymbolOutlines => editor.$outlines($($scope,)? chart, &outlines)?,
                Property::Trendlines => editor.$trendlines($($scope,)? chart, &trendlines)?,
            }
            editor.to_bytes()?
        }};
    }
    Ok(match target {
        Target::Keynote { slide } => style!(
            keynote::Draft::from_bytes(source)?, slide, ;
            set_slide_chart_series_strokes, set_slide_chart_series_fills,
            set_slide_chart_series_connection_lines, set_slide_chart_series_symbols,
            set_slide_chart_series_symbol_fills, set_slide_chart_series_symbol_outlines,
            set_slide_chart_series_trendlines
        ),
        Target::Numbers { sheet } => style!(
            numbers::Draft::from_bytes(source)?, sheet, ;
            set_sheet_chart_series_strokes, set_sheet_chart_series_fills,
            set_sheet_chart_series_connection_lines, set_sheet_chart_series_symbols,
            set_sheet_chart_series_symbol_fills, set_sheet_chart_series_symbol_outlines,
            set_sheet_chart_series_trendlines
        ),
        Target::Pages => style!(
            pages::Draft::from_bytes(source)?, ;
            set_body_chart_series_strokes, set_body_chart_series_fills,
            set_body_chart_series_connection_lines, set_body_chart_series_symbols,
            set_body_chart_series_symbol_fills, set_body_chart_series_symbol_outlines,
            set_body_chart_series_trendlines
        ),
    })
}

// Host values rendered in the shared vocabulary, so both reads compare as
// the same `Debug` text.

fn from_host_stroke(stroke: ChartSeriesStroke) -> Stroke {
    let pattern = match stroke.pattern {
        ChartSeriesStrokePattern::Solid => StrokePattern::Solid,
        ChartSeriesStrokePattern::MediumDash => StrokePattern::MediumDash,
        ChartSeriesStrokePattern::RoundedDash => StrokePattern::RoundedDash,
    };
    Stroke::new(stroke.color, stroke.width, pattern)
}

fn from_host_technique(technique: ShapeImageFillTechnique) -> ImageTechnique {
    match technique {
        ShapeImageFillTechnique::OriginalSize => ImageTechnique::OriginalSize,
        ShapeImageFillTechnique::Stretch => ImageTechnique::Stretch,
        ShapeImageFillTechnique::Tile => ImageTechnique::Tile,
        ShapeImageFillTechnique::ScaleToFill => ImageTechnique::ScaleToFill,
        ShapeImageFillTechnique::ScaleToFit => ImageTechnique::ScaleToFit,
    }
}

fn from_host_fill(fill: ShapeFill) -> TestResult<Fill> {
    Ok(match fill {
        ShapeFill::None => Fill::None,
        ShapeFill::Solid(color) => Fill::Solid(color),
        ShapeFill::Gradient(gradient) => Fill::Gradient(gradient),
        ShapeFill::Image(image) => {
            let technique = from_host_technique(image.technique());
            let mut value = match image.data_identifier() {
                Some(data) => {
                    Image::embedded(ImageData::new(data.get())?, technique, image.fill_size())?
                },
                None => Image::placeholder(technique, image.fill_size())?,
            };
            if let Some(tint) = image.tint() {
                value = value.with_tint(tint);
            }
            if let Some(reference) = image.reference_color() {
                value = value.with_reference_color(reference);
            }
            Fill::Image(
                value.with_generic_untagged_image(image.interprets_untagged_image_as_generic()),
            )
        },
    })
}

fn from_host_trendline(trendline: &ChartSeriesTrendline) -> TestResult<Trendline> {
    use litchi_iwa::charts::ChartSeriesTrendlineType as Host;
    let mut value = match trendline.trendline_type() {
        Host::None => Trendline::none(),
        Host::Linear => Trendline::linear(),
        Host::Logarithmic => Trendline::logarithmic(),
        Host::Polynomial(order) => Trendline::polynomial(PolynomialOrder::new(order.get())?),
        Host::Power => Trendline::power(),
        Host::Exponential => Trendline::exponential(),
        Host::MovingAverage(period) => Trendline::moving_average(
            litchi_iwa_common::chart::series::MovingAveragePeriod::new(period.get())?,
        ),
        Host::Unsupported(native) => Trendline::unsupported(native)?,
        _ => return Err("unknown host trendline type".into()),
    };
    if let Some(name) = trendline.custom_name() {
        value = value.with_legend_name(name)?;
    }
    Ok(value
        .with_legend_visibility(trendline.shows_name_in_legend())?
        .with_equation_visibility(trendline.shows_equation())?
        .with_r_squared_visibility(trendline.shows_r_squared())?)
}

pub(crate) fn host_read(
    bytes: &[u8],
    target: Target,
    chart: u64,
    property: Property,
) -> TestResult<String> {
    macro_rules! read {
        ($editor:expr, $($scope:expr,)? ; $strokes:ident, $fills:ident, $lines:ident, $symbols:ident, $symbol_fills:ident, $outlines:ident, $trendlines:ident) => {{
            let editor = $editor;
            match property {
                Property::Strokes => format!("{:?}", editor.$strokes($($scope,)? chart)?.into_iter().map(|s| s.map(from_host_stroke)).collect::<Vec<_>>()),
                Property::Fills => format!("{:?}", editor.$fills($($scope,)? chart)?.into_iter().map(from_host_fill).collect::<TestResult<Vec<_>>>()?),
                Property::ConnectionLines => format!("{:?}", editor.$lines($($scope,)? chart)?.into_iter().map(|line| match line {
                    ChartSeriesConnectionLine::Hidden => ConnectionLine::Hidden,
                    ChartSeriesConnectionLine::Straight => ConnectionLine::Straight,
                    ChartSeriesConnectionLine::Curved => ConnectionLine::Curved,
                }).collect::<Vec<_>>()),
                Property::Symbols => format!("{:?}", editor.$symbols($($scope,)? chart)?.into_iter().map(|symbol| symbol.map(|symbol| -> TestResult<Symbol> {
                    let shape = match symbol.shape {
                        ChartSeriesSymbolShape::Circle => SymbolShape::Circle,
                        ChartSeriesSymbolShape::Triangle => SymbolShape::Triangle,
                        ChartSeriesSymbolShape::InvertedTriangle => SymbolShape::InvertedTriangle,
                        ChartSeriesSymbolShape::Diamond => SymbolShape::Diamond,
                        ChartSeriesSymbolShape::Square => SymbolShape::Square,
                    };
                    Ok(Symbol { shape, size: symbol.size.map(|size| SymbolSize::new(size.points())).transpose()? })
                }).transpose()).collect::<TestResult<Vec<_>>>()?),
                Property::SymbolFills => format!("{:?}", editor.$symbol_fills($($scope,)? chart)?.into_iter().map(|fill| Ok(match fill {
                    ChartSeriesSymbolFill::SeriesFill => SymbolFill::SeriesFill,
                    ChartSeriesSymbolFill::SeriesStroke => SymbolFill::SeriesStroke,
                    ChartSeriesSymbolFill::Custom(fill) => SymbolFill::Custom(from_host_fill(fill)?),
                })).collect::<TestResult<Vec<_>>>()?),
                Property::SymbolOutlines => format!("{:?}", editor.$outlines($($scope,)? chart)?.into_iter().map(|s| s.map(from_host_stroke)).collect::<Vec<_>>()),
                Property::Trendlines => format!("{:?}", editor.$trendlines($($scope,)? chart)?.iter().map(from_host_trendline).collect::<TestResult<Vec<_>>>()?),
            }
        }};
    }
    Ok(match target {
        Target::Keynote { slide } => read!(
            keynote::Draft::from_bytes(bytes)?, slide, ;
            slide_chart_series_strokes, slide_chart_series_fills,
            slide_chart_series_connection_lines, slide_chart_series_symbols,
            slide_chart_series_symbol_fills, slide_chart_series_symbol_outlines,
            slide_chart_series_trendlines
        ),
        Target::Numbers { sheet } => read!(
            numbers::Draft::from_bytes(bytes)?, sheet, ;
            sheet_chart_series_strokes, sheet_chart_series_fills,
            sheet_chart_series_connection_lines, sheet_chart_series_symbols,
            sheet_chart_series_symbol_fills, sheet_chart_series_symbol_outlines,
            sheet_chart_series_trendlines
        ),
        Target::Pages => read!(
            pages::Draft::from_bytes(bytes)?, ;
            body_chart_series_strokes, body_chart_series_fills,
            body_chart_series_connection_lines, body_chart_series_symbols,
            body_chart_series_symbol_fills, body_chart_series_symbol_outlines,
            body_chart_series_trendlines
        ),
    })
}

/// Resolve a series-style property through the parent chain: `read`
/// returns `Some` when the style stores the property.
pub(crate) fn inherited<T>(
    components: &Components,
    first: u64,
    read: impl Fn(&Properties<'_>) -> TestResult<Option<T>>,
) -> TestResult<Option<T>> {
    let mut style = first;
    for _ in 0..MAX_DEPTH {
        let location = components.locate(style, Archive::SeriesStyle.message_type())?;
        let payload = components.payload(style, location)?;
        let properties = codec::decode(payload, Archive::SeriesStyle, Limits::DEFAULT)?;
        if let Some(value) = read(&properties)? {
            return Ok(Some(value));
        }
        match codec::parent(payload, Limits::DEFAULT)? {
            Some(parent) => style = parent,
            None => return Ok(None),
        }
    }
    Err("series style chain too deep".into())
}

pub(crate) fn stroke_at(
    properties: &Properties<'_>,
    field: u32,
) -> TestResult<Option<Option<Stroke>>> {
    let Some(bytes) = properties.message(field)? else {
        return Ok(None);
    };
    let stroke = plumbing::stroke(&stored::stroke(bytes, Limits::DEFAULT)?)?;
    Ok(Some(stroke.map(Stroke::from_shape_stroke).transpose()?))
}

pub(crate) fn fill_at(properties: &Properties<'_>, field: u32) -> TestResult<Option<Fill>> {
    let Some(bytes) = properties.message(field)? else {
        return Ok(None);
    };
    Ok(Some(plumbing::fill(&stored::fill(
        bytes,
        Limits::DEFAULT,
    )?)?))
}

pub(crate) fn flag(properties: &Properties<'_>, field: u32) -> TestResult<Option<bool>> {
    Ok(properties.varint(field)?.map(|value| value != 0))
}

/// The engine read of the case's property for every series.
pub(crate) fn engine_read(
    bytes: &[u8],
    chart: u64,
    property: Property,
    series: usize,
) -> TestResult<String> {
    let components = Components::read(bytes)?;
    let drawable = components.locate(chart, CHART_MESSAGE_TYPE)?;
    let references = codec::references(components.payload(chart, drawable)?, Limits::DEFAULT)?;
    let kind = property.kind();
    let style = |index: usize| references.series_style(index).ok_or("no series style");
    let non_style = |index: usize| -> TestResult<Option<Vec<u8>>> {
        let Some(object) = references
            .series_non_styles()
            .and_then(|sparse| sparse.get(u32::try_from(index).ok()?))
        else {
            return Ok(None);
        };
        let location = components.locate(object, Archive::SeriesNonStyle.message_type())?;
        Ok(Some(components.payload(object, location)?.to_vec()))
    };
    let mut rendered = Vec::new();
    for index in 0..series {
        rendered.push(match property {
            Property::Strokes => {
                let family = StrokeFamily::for_chart_kind(kind).ok_or("no stroke family")?;
                let value =
                    inherited(&components, style(index)?, |p| stroke_at(p, family.field()))?;
                format!("{:?}", value.flatten())
            },
            Property::SymbolOutlines => {
                let family =
                    SymbolOutlineFamily::for_chart_kind(kind).ok_or("no outline family")?;
                let value =
                    inherited(&components, style(index)?, |p| stroke_at(p, family.field()))?;
                format!("{:?}", value.flatten())
            },
            Property::Fills => {
                let family = FillFamily::for_chart_kind(kind).ok_or("no fill family")?;
                let value =
                    match inherited(&components, style(index)?, |p| fill_at(p, family.field()))? {
                        Some(fill) => fill,
                        None => inherited(&components, style(index)?, |p| {
                            fill_at(p, FillFamily::DEFAULT_FIELD)
                        })?
                        .ok_or("no inherited fill")?,
                    };
                format!("{value:?}")
            },
            Property::SymbolFills => {
                let family =
                    SymbolFillFamily::for_chart_kind(kind).ok_or("no symbol fill family")?;
                let value = inherited(&components, style(index)?, |p| {
                    let use_fill = match family.use_series_fill_field() {
                        Some(field) => flag(p, field)?,
                        None => None,
                    };
                    let use_stroke = flag(p, family.use_series_stroke_field())?;
                    Ok(family.read(use_fill, use_stroke, fill_at(p, family.fill_field())?)?)
                })?
                .unwrap_or_else(|| family.default_fill());
                format!("{value:?}")
            },
            Property::ConnectionLines => {
                let family = ConnectionLineFamily::for_chart_kind(kind).ok_or("no line family")?;
                let value = match non_style(index)? {
                    Some(payload) => {
                        let p = codec::decode(&payload, Archive::SeriesNonStyle, Limits::DEFAULT)?;
                        family.read(
                            flag(&p, ConnectionLineFamily::SCATTER_SHOW_FIELD)?,
                            p.varint(family.field())?.map(|value| value as i32),
                        )?
                    },
                    None => family.read(None, None)?,
                };
                format!("{value:?}")
            },
            Property::Symbols => {
                let family = SymbolFamily::for_chart_kind(kind).ok_or("no symbol family")?;
                let (visible, shape) = match non_style(index)? {
                    Some(payload) => {
                        let p = codec::decode(&payload, Archive::SeriesNonStyle, Limits::DEFAULT)?;
                        (
                            flag(&p, family.show_field())?.unwrap_or(false),
                            SymbolShape::from_native(p.int32(family.type_field())?.unwrap_or(0))?,
                        )
                    },
                    None => (false, SymbolShape::Circle),
                };
                let value = if visible {
                    let size = inherited(&components, style(index)?, |p| {
                        Ok(p.float(family.size_field())?
                            .map(SymbolSize::new)
                            .transpose()?)
                    })?;
                    Some(Symbol { shape, size })
                } else {
                    None
                };
                format!("{value:?}")
            },
            Property::Trendlines => {
                let value = match non_style(index)? {
                    Some(payload) => {
                        let p = codec::decode(&payload, Archive::SeriesNonStyle, Limits::DEFAULT)?;
                        let label = p
                            .bytes(trendline_field::LABEL)?
                            .map(|bytes| std::str::from_utf8(bytes).map(str::to_owned))
                            .transpose()?;
                        Trendline::from_native(
                            p.varint(trendline_field::SHOW)?,
                            p.varint(trendline_field::TYPE)?,
                            p.varint(trendline_field::POLYNOMIAL_ORDER)?,
                            p.varint(trendline_field::MOVING_AVERAGE_PERIOD)?,
                            p.varint(trendline_field::SHOW_LABEL)?,
                            p.varint(trendline_field::SHOW_EQUATION)?,
                            p.varint(trendline_field::SHOW_R_SQUARED)?,
                            label,
                        )?
                    },
                    None => Trendline::none(),
                };
                format!("{value:?}")
            },
        });
    }
    Ok(format!("[{}]", rendered.join(", ")))
}

/// Re-render the host read in the engine's element format.
fn host_rendered(
    bytes: &[u8],
    target: Target,
    chart: u64,
    property: Property,
) -> TestResult<String> {
    host_read(bytes, target, chart, property)
}

pub(crate) fn leaf_reopen(format: Format, source: &[u8]) -> TestResult<Vec<u8>> {
    Ok(match format {
        Format::Numbers => leaf_bytes!(litchi_numbers::Package::from_bytes(source)?),
        Format::Pages => leaf_bytes!(litchi_pages::Package::from_bytes(source)?),
        Format::Keynote => leaf_bytes!(litchi_keynote::Package::from_bytes(source)?),
    })
}

fn series_case(
    name: &'static str,
    format: Format,
    tier: Tier,
    evidence: Evidence,
    property: Property,
) -> TestResult {
    let mut input = input(format, property.kind())?;
    let (target, chart) = (input.target, input.chart);
    input.fixture.bytes = host_style(&input.fixture.bytes, target, chart, property)?;
    // Both sides render a list of per-series values; the engine joins its
    // elements the way `Debug` renders a `Vec`.
    let normalize = |text: String| text.replace(['\n', ' '], "");
    Case::new(name, format, tier, evidence, &input.fixture)
        .read_only()
        .host(|bytes| Ok(litchi_iwa::Package::from_bytes(bytes)?.to_bytes()?))
        .leaf(move |bytes| Ok(LeafRun::output(leaf_reopen(format, bytes)?)))
        .read("host and engine series reads agree", move |bytes| {
            let host = normalize(host_rendered(bytes, target, chart, property)?);
            let engine = normalize(engine_read(bytes, chart, property, 2)?);
            if host != engine {
                return Err(format!("host {host}\nengine {engine}").into());
            }
            Ok(host)
        })
        .run()?;
    Ok(())
}

#[test]
fn keynote_chart_series_strokes_read_parity() -> TestResult {
    series_case(
        "keynote_chart_series_strokes_read_parity",
        Format::Keynote,
        Tier::A,
        Evidence::Queued,
        Property::Strokes,
    )
}

#[test]
fn numbers_chart_series_strokes_read_parity() -> TestResult {
    series_case(
        "numbers_chart_series_strokes_read_parity",
        Format::Numbers,
        Tier::A,
        Evidence::Queued,
        Property::Strokes,
    )
}

#[test]
fn pages_chart_series_strokes_read_parity() -> TestResult {
    series_case(
        "pages_chart_series_strokes_read_parity",
        Format::Pages,
        Tier::A,
        Evidence::Queued,
        Property::Strokes,
    )
}

#[test]
fn keynote_chart_series_fills_read_parity() -> TestResult {
    series_case(
        "keynote_chart_series_fills_read_parity",
        Format::Keynote,
        Tier::A,
        Evidence::Queued,
        Property::Fills,
    )
}

#[test]
fn numbers_chart_series_fills_read_parity() -> TestResult {
    series_case(
        "numbers_chart_series_fills_read_parity",
        Format::Numbers,
        Tier::A,
        Evidence::Queued,
        Property::Fills,
    )
}

#[test]
fn pages_chart_series_fills_read_parity() -> TestResult {
    series_case(
        "pages_chart_series_fills_read_parity",
        Format::Pages,
        Tier::A,
        Evidence::Queued,
        Property::Fills,
    )
}

#[test]
fn keynote_chart_series_connection_lines_read_parity() -> TestResult {
    series_case(
        "keynote_chart_series_connection_lines_read_parity",
        Format::Keynote,
        Tier::A,
        Evidence::Queued,
        Property::ConnectionLines,
    )
}

#[test]
fn numbers_chart_series_connection_lines_read_parity() -> TestResult {
    series_case(
        "numbers_chart_series_connection_lines_read_parity",
        Format::Numbers,
        Tier::A,
        Evidence::Queued,
        Property::ConnectionLines,
    )
}

#[test]
fn pages_chart_series_connection_lines_read_parity() -> TestResult {
    series_case(
        "pages_chart_series_connection_lines_read_parity",
        Format::Pages,
        Tier::A,
        Evidence::Queued,
        Property::ConnectionLines,
    )
}

#[test]
fn keynote_chart_series_symbols_read_parity() -> TestResult {
    series_case(
        "keynote_chart_series_symbols_read_parity",
        Format::Keynote,
        Tier::A,
        Evidence::Queued,
        Property::Symbols,
    )
}

#[test]
fn numbers_chart_series_symbols_read_parity() -> TestResult {
    series_case(
        "numbers_chart_series_symbols_read_parity",
        Format::Numbers,
        Tier::A,
        Evidence::Queued,
        Property::Symbols,
    )
}

#[test]
fn pages_chart_series_symbols_read_parity() -> TestResult {
    series_case(
        "pages_chart_series_symbols_read_parity",
        Format::Pages,
        Tier::A,
        Evidence::Queued,
        Property::Symbols,
    )
}

#[test]
fn keynote_chart_series_symbol_fills_read_parity() -> TestResult {
    series_case(
        "keynote_chart_series_symbol_fills_read_parity",
        Format::Keynote,
        Tier::A,
        Evidence::Queued,
        Property::SymbolFills,
    )
}

#[test]
fn numbers_chart_series_symbol_fills_read_parity() -> TestResult {
    series_case(
        "numbers_chart_series_symbol_fills_read_parity",
        Format::Numbers,
        Tier::A,
        Evidence::Queued,
        Property::SymbolFills,
    )
}

#[test]
fn pages_chart_series_symbol_fills_read_parity() -> TestResult {
    series_case(
        "pages_chart_series_symbol_fills_read_parity",
        Format::Pages,
        Tier::A,
        Evidence::Queued,
        Property::SymbolFills,
    )
}

#[test]
fn keynote_chart_series_symbol_outlines_read_parity() -> TestResult {
    series_case(
        "keynote_chart_series_symbol_outlines_read_parity",
        Format::Keynote,
        Tier::A,
        Evidence::Queued,
        Property::SymbolOutlines,
    )
}

#[test]
fn numbers_chart_series_symbol_outlines_read_parity() -> TestResult {
    series_case(
        "numbers_chart_series_symbol_outlines_read_parity",
        Format::Numbers,
        Tier::A,
        Evidence::Queued,
        Property::SymbolOutlines,
    )
}

#[test]
fn pages_chart_series_symbol_outlines_read_parity() -> TestResult {
    series_case(
        "pages_chart_series_symbol_outlines_read_parity",
        Format::Pages,
        Tier::A,
        Evidence::Queued,
        Property::SymbolOutlines,
    )
}

#[test]
fn keynote_chart_series_trendlines_read_parity() -> TestResult {
    series_case(
        "keynote_chart_series_trendlines_read_parity",
        Format::Keynote,
        Tier::A,
        Evidence::Queued,
        Property::Trendlines,
    )
}

#[test]
fn numbers_chart_series_trendlines_read_parity() -> TestResult {
    series_case(
        "numbers_chart_series_trendlines_read_parity",
        Format::Numbers,
        Tier::A,
        Evidence::Queued,
        Property::Trendlines,
    )
}

#[test]
fn pages_chart_series_trendlines_read_parity() -> TestResult {
    series_case(
        "pages_chart_series_trendlines_read_parity",
        Format::Pages,
        Tier::A,
        Evidence::Queued,
        Property::Trendlines,
    )
}

// ---- shape image fill -------------------------------------------------

/// A native fixture with media, a shape the host added and image-filled
/// with the first embedded image asset.
fn image_fill_input(format: Format) -> TestResult<(Fixture, Target, u64)> {
    let path = match format {
        Format::Keynote => "iwork/keynote/slide-movie-geometry-source-native.key",
        Format::Pages => "iwork/pages/image-adjustments-native.pages",
        Format::Numbers => "iwork/numbers/table-relocation-native.numbers",
    };
    let native = crate::harness::fixture(path)?;
    let media = litchi_iwa::media::Draft::from_bytes(&native.bytes)?;
    let asset = media
        .assets()
        .iter()
        .find(|asset| {
            asset.is_materialized() && asset.media_type == litchi_iwa_common::media::Type::Image
        })
        .ok_or("no embedded image asset")?;
    let identifier = ShapeImageDataIdentifier::new(asset.data_identifier.to_string().parse()?)?;
    let fill = ShapeFill::Image(
        ShapeImageFill::embedded(
            identifier,
            ShapeImageFillTechnique::ScaleToFit,
            DrawableSize {
                width: 320.0,
                height: 180.0,
            },
        )?
        .with_tint(RgbaColor::new(
            0.2,
            0.3,
            0.9,
            0.5,
            RgbColorSpace::DisplayP3,
        )?)
        .with_generic_untagged_image(true),
    );
    let (position, size) = (
        DrawablePoint { x: 80.0, y: 90.0 },
        DrawableSize {
            width: 200.0,
            height: 120.0,
        },
    );
    Ok(match format {
        Format::Keynote => {
            let mut editor = keynote::Draft::from_bytes(&native.bytes)?;
            let shape = editor.add_slide_shape(0, "", position, size, Preset::Rectangle)?;
            editor.set_slide_shape_fill(0, shape.drawable_object_id, &fill)?;
            (
                Fixture {
                    path,
                    bytes: editor.to_bytes()?,
                },
                Target::Keynote { slide: 0 },
                shape.drawable_object_id,
            )
        },
        Format::Numbers => {
            let mut editor = numbers::Draft::from_bytes(&native.bytes)?;
            let sheet = editor.sheets()?.first().ok_or("no sheet")?.id();
            let shape = editor.add_sheet_shape(sheet, "", position, size, Preset::Rectangle)?;
            editor.set_sheet_shape_fill(sheet, shape.drawable_object_id, &fill)?;
            (
                Fixture {
                    path,
                    bytes: editor.to_bytes()?,
                },
                Target::Numbers { sheet },
                shape.drawable_object_id,
            )
        },
        Format::Pages => {
            let mut editor = pages::Draft::from_bytes(&native.bytes)?;
            let shape = editor.add_body_shape(0, "", position, size, Preset::Rectangle)?;
            editor.set_body_shape_fill(shape.drawable_object_id, &fill)?;
            (
                Fixture {
                    path,
                    bytes: editor.to_bytes()?,
                },
                Target::Pages,
                shape.drawable_object_id,
            )
        },
    })
}

fn host_shape_fill(bytes: &[u8], target: Target, shape: u64) -> TestResult<String> {
    let fill = match target {
        Target::Keynote { slide } => {
            keynote::Draft::from_bytes(bytes)?.slide_shape_fill(slide, shape)?
        },
        Target::Numbers { sheet } => {
            numbers::Draft::from_bytes(bytes)?.sheet_shape_fill(sheet, shape)?
        },
        Target::Pages => pages::Draft::from_bytes(bytes)?.body_shape_fill(shape)?,
    };
    Ok(format!("{:?}", from_host_fill(fill)?))
}

/// The engine read of a shape's fill: its style chain (`TSWP.ShapeStyleArchive`).
fn engine_shape_fill(bytes: &[u8], shape: u64) -> TestResult<String> {
    let components = Components::read(bytes)?;
    let location = components.locate(shape, SHAPE_INFO_MESSAGE_TYPE)?;
    let mut style =
        shape_codec::style_reference(components.payload(shape, location)?, Limits::DEFAULT)?;
    for _ in 0..MAX_DEPTH {
        let Some(identifier) = style else {
            return Ok(format!("{:?}", Fill::None));
        };
        let location = components.locate(identifier, SHAPE_STYLE_MESSAGE_TYPE)?;
        let decoded =
            stored::shape_style(components.payload(identifier, location)?, Limits::DEFAULT)?;
        if let Some(fill) = decoded.fill {
            return Ok(format!(
                "{:?}",
                plumbing::fill(&stored::fill(fill, Limits::DEFAULT)?)?
            ));
        }
        style = decoded.parent;
    }
    Err("shape style chain too deep".into())
}

fn image_fill_case(
    name: &'static str,
    format: Format,
    tier: Tier,
    evidence: Evidence,
) -> TestResult {
    let (input, target, shape) = image_fill_input(format)?;
    Case::new(name, format, tier, evidence, &input)
        .read_only()
        .host(|bytes| Ok(litchi_iwa::Package::from_bytes(bytes)?.to_bytes()?))
        .leaf(move |bytes| Ok(LeafRun::output(leaf_reopen(format, bytes)?)))
        .read("host shape fill", move |bytes| {
            host_shape_fill(bytes, target, shape)
        })
        .read("engine shape fill", move |bytes| {
            engine_shape_fill(bytes, shape)
        })
        .run()?;
    Ok(())
}

#[test]
fn keynote_shape_image_fill_read_parity() -> TestResult {
    image_fill_case(
        "keynote_shape_image_fill_read_parity",
        Format::Keynote,
        Tier::A,
        Evidence::Queued,
    )
}

#[test]
fn numbers_shape_image_fill_read_parity() -> TestResult {
    image_fill_case(
        "numbers_shape_image_fill_read_parity",
        Format::Numbers,
        Tier::A,
        Evidence::Queued,
    )
}

#[test]
fn pages_shape_image_fill_read_parity() -> TestResult {
    image_fill_case(
        "pages_shape_image_fill_read_parity",
        Format::Pages,
        Tier::A,
        Evidence::Queued,
    )
}

#[inline(never)]
fn host_series_route(
    bytes: &[u8],
    target: Target,
    chart: u64,
    property: Property,
) -> TestResult<String> {
    host_read(bytes, target, chart, property)
}

#[inline(never)]
fn leaf_series_route(bytes: &[u8], chart: u64, property: Property) -> TestResult<String> {
    engine_read(bytes, chart, property, 2)
}

/// ABBA cost probe of the series-style reads (record 0375 §5); toggle
/// `*series_style_cases::host_series_route*` or
/// `*series_style_cases::leaf_series_route*`.
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn series_style_read_cost_abba() -> TestResult {
    for property in [Property::Strokes, Property::Fills, Property::Trendlines] {
        for format in [Format::Keynote, Format::Numbers, Format::Pages] {
            let mut input = input(format, property.kind())?;
            let (target, chart) = (input.target, input.chart);
            input.fixture.bytes = host_style(&input.fixture.bytes, target, chart, property)?;
            let bytes = input.fixture.bytes.as_slice();
            std::hint::black_box(host_series_route(bytes, target, chart, property)?);
            std::hint::black_box(leaf_series_route(bytes, chart, property)?);
            std::hint::black_box(leaf_series_route(bytes, chart, property)?);
            std::hint::black_box(host_series_route(bytes, target, chart, property)?);
        }
    }
    Ok(())
}
