//! Pages body chart per-series property parity (wave W4-C, record 0414):
//! the host `set_body_chart_*` all-series and one-series routes
//! (connection lines, wedge explosions, label distances, data symbols,
//! trendlines, error bars, pie labels, value-label visibility) against
//! `litchi_pages::Package::edit_body_chart_series` on W1-E's graph
//! engine.

use litchi_iwa::charts::{
    ChartPieLabelDistance, ChartPieWedgeExplosion, ChartPieWedgeIndex, ChartSeriesConnectionLine,
    ChartSeriesSymbol, ChartSeriesSymbolShape, ChartSeriesSymbolSize, ChartSeriesTrendline,
    ChartSeriesTrendlineMovingAveragePeriod, ChartSeriesTrendlinePolynomialOrder, Index,
};
use litchi_iwa::pages;
use litchi_iwa_common::chart::error_bar::{
    CustomValues, Direction as ErrorBarDirection, FixedValue, Percentage, Series as ErrorBarSeries,
};
use litchi_iwa_common::chart::kind::Kind;
use litchi_iwa_common::chart::number_format::{
    DecimalPlaces, LabelAffixes, NegativeStyle, NumberFormat,
};
use litchi_iwa_common::chart::pie::{LabelDistance, LabelVisibility, WedgeExplosion};
use litchi_iwa_common::chart::series::{
    ConnectionLine, MovingAveragePeriod, PolynomialOrder, Symbol, SymbolShape, SymbolSize,
    Trendline,
};
use litchi_iwa_common::chart::series_labels::Visibility;
use litchi_pages::Package;
use litchi_pages::chart::series::{Property, PropertyKind};

use crate::chart_style_cases::{Input, input};
use crate::harness::{Case, Evidence, Format, LeafRun, TestResult, Tier};
use crate::leaf_bytes;

/// A trendline configuration, spelled once for both vocabularies.
#[derive(Clone, Copy, Debug)]
enum Trend {
    None,
    NamedLinear,
    Polynomial,
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

    fn leaf(self) -> TestResult<Trendline> {
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

/// One per-series write: every series, in series order.
#[derive(Clone, Debug)]
enum Values {
    Lines(Kind, Vec<ConnectionLine>),
    Wedges(Vec<f32>),
    Labels(Vec<f32>),
    Symbols(Kind, Vec<Option<(SymbolShape, Option<f32>)>>),
    Trends(Vec<Trend>),
    Errors(Vec<ErrorBarSeries>),
    PieLabels(Vec<LabelVisibility>),
    ValueLabels(Vec<Visibility>),
    Affixes(Vec<LabelAffixes>),
    Formats(Vec<NumberFormat>),
}

impl Values {
    const fn chart(&self) -> Kind {
        match self {
            Self::Lines(kind, _) | Self::Symbols(kind, _) => *kind,
            Self::Wedges(_) | Self::Labels(_) | Self::PieLabels(_) => Kind::Pie2d,
            Self::Trends(_) => Kind::Scatter2d,
            Self::Errors(_) | Self::ValueLabels(_) | Self::Affixes(_) | Self::Formats(_) => {
                Kind::Column2d
            },
        }
    }

    const fn kind(&self) -> PropertyKind {
        match self {
            Self::Lines(..) => PropertyKind::ConnectionLines,
            Self::Wedges(_) => PropertyKind::WedgeExplosions,
            Self::Labels(_) => PropertyKind::LabelDistances,
            Self::Symbols(..) => PropertyKind::Symbols,
            Self::Trends(_) => PropertyKind::Trendlines,
            Self::Errors(_) => PropertyKind::ErrorBars,
            Self::PieLabels(_) => PropertyKind::PieLabels,
            Self::ValueLabels(_) => PropertyKind::ValueLabels,
            Self::Affixes(_) => PropertyKind::ValueLabelAffixes,
            Self::Formats(_) => PropertyKind::ValueLabelNumberFormats,
        }
    }

    fn leaf(&self) -> TestResult<Property> {
        Ok(match self {
            Self::Lines(_, values) => Property::ConnectionLines(values.clone()),
            Self::Wedges(values) => Property::WedgeExplosions(
                values
                    .iter()
                    .map(|percent| WedgeExplosion::from_percent(*percent))
                    .collect::<Result<_, _>>()?,
            ),
            Self::Labels(values) => Property::LabelDistances(
                values
                    .iter()
                    .map(|percent| LabelDistance::from_percent(*percent))
                    .collect::<Result<_, _>>()?,
            ),
            Self::Symbols(_, values) => Property::Symbols(
                values
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
                    .collect::<Result<_, _>>()?,
            ),
            Self::Trends(values) => Property::Trendlines(
                values
                    .iter()
                    .map(|trend| trend.leaf())
                    .collect::<Result<_, _>>()?,
            ),
            Self::Errors(values) => Property::ErrorBars(values.clone()),
            Self::PieLabels(values) => Property::PieLabels(values.clone()),
            Self::ValueLabels(values) => Property::ValueLabels(values.clone()),
            Self::Affixes(values) => Property::ValueLabelAffixes(values.clone()),
            Self::Formats(values) => Property::ValueLabelNumberFormats(values.clone()),
        })
    }
}

const fn host_line(value: ConnectionLine) -> ChartSeriesConnectionLine {
    match value {
        ConnectionLine::Hidden => ChartSeriesConnectionLine::Hidden,
        ConnectionLine::Straight => ChartSeriesConnectionLine::Straight,
        ConnectionLine::Curved => ChartSeriesConnectionLine::Curved,
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

fn host_symbol(value: Option<(SymbolShape, Option<f32>)>) -> TestResult<Option<ChartSeriesSymbol>> {
    Ok(match value {
        None => None,
        Some((shape, None)) => Some(ChartSeriesSymbol::automatic(host_shape(shape))),
        Some((shape, Some(points))) => Some(ChartSeriesSymbol::sized(
            host_shape(shape),
            ChartSeriesSymbolSize::new(points)?,
        )),
    })
}

/// The host route: `None` writes every series, `Some(i)` only series `i`
/// through the one-series route.
fn host_write(
    source: &[u8],
    __sheet: (),
    chart: u64,
    values: &Values,
    one: Option<usize>,
) -> TestResult<Vec<u8>> {
    let mut e = pages::Draft::from_bytes(source)?;
    macro_rules! route {
        ($all:ident, $single:ident, $index:expr, $items:expr) => {{
            let items = $items;
            match one {
                None => e.$all(chart, &items)?,
                Some(i) => e.$single(chart, $index(i), items[i].clone())?,
            }
        }};
    }
    let series = Index::from_zero_based;
    let wedge = ChartPieWedgeIndex::from_zero_based;
    match values {
        Values::Lines(_, v) => route!(
            set_body_chart_series_connection_lines,
            set_body_chart_series_connection_line,
            series,
            v.iter().copied().map(host_line).collect::<Vec<_>>()
        ),
        Values::Wedges(v) => route!(
            set_body_chart_pie_wedge_explosions,
            set_body_chart_pie_wedge_explosion,
            wedge,
            v.iter()
                .map(|percent| ChartPieWedgeExplosion::from_percent(*percent))
                .collect::<Result<Vec<_>, _>>()?
        ),
        Values::Labels(v) => route!(
            set_body_chart_pie_label_distances,
            set_body_chart_pie_label_distance,
            wedge,
            v.iter()
                .map(|percent| ChartPieLabelDistance::from_percent(*percent))
                .collect::<Result<Vec<_>, _>>()?
        ),
        Values::Symbols(_, v) => route!(
            set_body_chart_series_symbols,
            set_body_chart_series_symbol,
            series,
            v.iter()
                .copied()
                .map(host_symbol)
                .collect::<TestResult<Vec<_>>>()?
        ),
        Values::Trends(v) => route!(
            set_body_chart_series_trendlines,
            set_body_chart_series_trendline,
            series,
            v.iter()
                .map(|trend| trend.host())
                .collect::<TestResult<Vec<_>>>()?
        ),
        Values::Errors(v) => route!(
            set_body_chart_series_error_bars,
            set_body_chart_series_error_bar,
            series,
            v.clone()
        ),
        Values::PieLabels(v) => route!(
            set_body_chart_pie_label_visibilities,
            set_body_chart_pie_label_visibility,
            wedge,
            v.clone()
        ),
        Values::ValueLabels(v) => route!(
            set_body_chart_series_value_label_visibilities,
            set_body_chart_series_value_label_visibility,
            series,
            v.clone()
        ),
        Values::Affixes(v) => route!(
            set_body_chart_series_value_label_affixes,
            set_body_chart_series_value_label_affix,
            series,
            v.clone()
        ),
        Values::Formats(v) => route!(
            set_body_chart_series_value_label_number_formats,
            set_body_chart_series_value_label_number_format,
            series,
            v.clone()
        ),
    }
    Ok(e.to_bytes()?)
}

/// The host's read of the property, as comparable text.
fn host_read(bytes: &[u8], __sheet: (), chart: u64, values: &Values) -> TestResult<String> {
    let e = pages::Draft::from_bytes(bytes)?;
    Ok(match values {
        Values::Lines(..) => format!("{:?}", e.body_chart_series_connection_lines(chart)?),
        Values::Wedges(_) => format!("{:?}", e.body_chart_pie_wedge_explosions(chart)?),
        Values::Labels(_) => format!("{:?}", e.body_chart_pie_label_distances(chart)?),
        Values::Symbols(..) => format!("{:?}", e.body_chart_series_symbols(chart)?),
        Values::Trends(_) => format!("{:?}", e.body_chart_series_trendlines(chart)?),
        Values::Errors(_) => format!("{:?}", e.body_chart_series_error_bars(chart)?),
        Values::PieLabels(_) => {
            format!("{:?}", e.body_chart_pie_label_visibilities(chart)?)
        },
        Values::ValueLabels(_) => {
            format!("{:?}", e.body_chart_series_value_label_visibilities(chart)?)
        },
        Values::Affixes(_) => format!("{:?}", e.body_chart_series_value_label_affixes(chart)?),
        Values::Formats(_) => format!(
            "{:?}",
            e.body_chart_series_value_label_number_formats(chart)?
        ),
    })
}

/// The leaf position of the host-added chart: the sheet's last chart.
/// The host-added chart is the body's only discoverable chart.
fn last_chart(_bytes: &[u8]) -> TestResult<usize> {
    Ok(0)
}

fn leaf_write(source: &[u8], values: &Values, one: Option<usize>) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let chart = last_chart(source)?;
    let property = match one {
        None => values.leaf()?,
        Some(i) => {
            // The one-series route is a read, one replacement and a write.
            let wanted = values.leaf()?;
            let mut current = package.body_chart_series(chart, values.kind())?;
            replace(&mut current, &wanted, i)?;
            current
        },
    };
    let mut edit = package.edit_body_chart_series(chart)?;
    edit = edit.set(property)?;
    let commit = edit.commit()?;
    let restored = commit
        .package()
        .apply_body_chart_series(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

/// Series `i` of `current` takes the value of `wanted`.
fn replace(current: &mut Property, wanted: &Property, i: usize) -> TestResult {
    macro_rules! take {
        ($($variant:ident),*) => {
            match (current, wanted) {
                $((Property::$variant(current), Property::$variant(wanted)) => {
                    *current.get_mut(i).ok_or("series")? = wanted.get(i).ok_or("series")?.clone();
                },)*
                _ => return Err("property mismatch".into()),
            }
        };
    }
    take!(
        ConnectionLines,
        WedgeExplosions,
        LabelDistances,
        Symbols,
        Trendlines,
        ErrorBars,
        PieLabels,
        ValueLabels,
        ValueLabelAffixes,
        ValueLabelNumberFormats
    );
    Ok(())
}

fn chart_input(values: &Values, before: &[Values]) -> TestResult<Input> {
    let mut input = input(Format::Pages, values.chart())?;
    let sheet = ();
    for earlier in before {
        input.fixture.bytes = host_write(&input.fixture.bytes, sheet, input.chart, earlier, None)?;
    }
    Ok(input)
}

fn run(
    name: &'static str,
    tier: Tier,
    before: &[Values],
    values: &Values,
    one: Option<usize>,
) -> TestResult {
    let input = chart_input(values, before)?;
    let sheet = ();
    let chart = input.chart;
    let mut case = Case::new(name, Format::Pages, tier, Evidence::Queued, &input.fixture)
        .host(|source| host_write(source, sheet, chart, values, one))
        .leaf(|source| leaf_write(source, values, one))
        .read("host series read", move |bytes| {
            host_read(bytes, sheet, chart, values)
        });
    if tier == Tier::B {
        case = case.leaf_structural_metadata();
    }
    case.run()
        .map_err(|error| format!("{values:?} {one:?}: {error}"))?;
    Ok(())
}

fn lines() -> (Values, Values) {
    (
        Values::Lines(
            Kind::Line2d,
            vec![ConnectionLine::Curved, ConnectionLine::Straight],
        ),
        Values::Lines(
            Kind::Line2d,
            vec![ConnectionLine::Straight, ConnectionLine::Straight],
        ),
    )
}

fn errors() -> TestResult<Values> {
    Ok(Values::Errors(vec![
        ErrorBarSeries::FixedValue {
            direction: ErrorBarDirection::PositiveAndNegative,
            value: FixedValue::new(2.5)?,
        },
        ErrorBarSeries::Percentage {
            direction: ErrorBarDirection::PositiveOnly,
            percentage: Percentage::new(15)?,
        },
    ]))
}

fn custom_errors() -> TestResult<Values> {
    Ok(Values::Errors(vec![
        ErrorBarSeries::None,
        ErrorBarSeries::CustomValues {
            direction: ErrorBarDirection::PositiveAndNegative,
            values: CustomValues::new([1.0, 2.0, 3.0], [0.5, 0.5, 0.5])?,
        },
    ]))
}

#[test]
fn pages_body_chart_series_connection_lines_parity() -> TestResult {
    let _ = Evidence::Queued;
    let name = "pages_body_chart_series_connection_lines_parity";
    let (curved, straight) = lines();
    run(name, Tier::B, &[], &curved, None)?;
    run(
        name,
        Tier::B,
        std::slice::from_ref(&curved),
        &straight,
        None,
    )?;
    run(name, Tier::B, &[], &curved, Some(0))?;
    Ok(())
}

#[test]
fn pages_body_chart_pie_wedge_explosions_parity() -> TestResult {
    let _ = Evidence::Queued;
    let name = "pages_body_chart_pie_wedge_explosions_parity";
    let exploded = Values::Wedges(vec![10.0, 0.0, 25.0]);
    run(name, Tier::B, &[], &exploded, None)?;
    run(
        name,
        Tier::B,
        std::slice::from_ref(&exploded),
        &Values::Wedges(vec![0.0, 0.0, 0.0]),
        None,
    )?;
    run(name, Tier::B, &[], &exploded, Some(2))?;
    Ok(())
}

#[test]
fn pages_body_chart_pie_label_distances_parity() -> TestResult {
    let _ = Evidence::Queued;
    let name = "pages_body_chart_pie_label_distances_parity";
    let far = Values::Labels(vec![40.0, 55.0, 70.0]);
    run(name, Tier::B, &[], &far, None)?;
    run(name, Tier::B, &[], &far, Some(1))?;
    Ok(())
}

#[test]
fn pages_body_chart_series_trendlines_parity() -> TestResult {
    let _ = Evidence::Queued;
    let name = "pages_body_chart_series_trendlines_parity";
    let trends = Values::Trends(vec![Trend::NamedLinear]);
    run(name, Tier::B, &[], &trends, None)?;
    run(
        name,
        Tier::B,
        std::slice::from_ref(&trends),
        &Values::Trends(vec![Trend::Polynomial]),
        None,
    )?;
    run(
        name,
        Tier::B,
        std::slice::from_ref(&trends),
        &Values::Trends(vec![Trend::MovingAverage]),
        Some(0),
    )?;
    run(
        name,
        Tier::B,
        &[trends],
        &Values::Trends(vec![Trend::None]),
        None,
    )?;
    Ok(())
}

#[test]
fn pages_body_chart_series_error_bars_parity() -> TestResult {
    let _ = Evidence::Queued;
    let name = "pages_body_chart_series_error_bars_parity";
    run(name, Tier::B, &[], &errors()?, None)?;
    run(name, Tier::B, &[errors()?], &custom_errors()?, None)?;
    run(name, Tier::B, &[], &errors()?, Some(1))?;
    Ok(())
}

#[test]
fn pages_body_chart_pie_labels_parity() -> TestResult {
    let _ = Evidence::Queued;
    let name = "pages_body_chart_pie_labels_parity";
    let labels = Values::PieLabels(vec![
        LabelVisibility::ALL,
        LabelVisibility::HIDDEN,
        LabelVisibility::DATA_POINT_NAMES_ONLY,
    ]);
    run(name, Tier::B, &[], &labels, None)?;
    run(name, Tier::B, &[], &labels, Some(0))?;
    Ok(())
}

#[test]
fn pages_body_chart_series_value_labels_parity() -> TestResult {
    let _ = Evidence::Queued;
    let name = "pages_body_chart_series_value_labels_parity";
    let shown = Values::ValueLabels(vec![Visibility::Visible, Visibility::Hidden]);
    run(name, Tier::B, &[], &shown, None)?;
    run(name, Tier::B, &[], &shown, Some(0))?;
    Ok(())
}

fn affixes() -> TestResult<Values> {
    Ok(Values::Affixes(vec![
        LabelAffixes::new("$", " M")?,
        LabelAffixes::new("", "%")?,
    ]))
}

fn formats() -> TestResult<Values> {
    Ok(Values::Formats(vec![
        NumberFormat::new(DecimalPlaces::fixed(2)?, NegativeStyle::Parentheses, false),
        NumberFormat::SERIES_VALUE_LABEL_NATIVE_DEFAULT,
    ]))
}

#[test]
fn pages_body_chart_series_value_label_affixes_parity() -> TestResult {
    let _ = Evidence::Queued;
    let name = "pages_body_chart_series_value_label_affixes_parity";
    run(name, Tier::B, &[], &affixes()?, None)?;
    run(
        name,
        Tier::B,
        &[affixes()?],
        &Values::Affixes(vec![
            LabelAffixes::new("", "")?,
            LabelAffixes::new("~", "")?,
        ]),
        None,
    )?;
    run(name, Tier::B, &[], &affixes()?, Some(1))?;
    Ok(())
}

#[test]
fn pages_body_chart_series_value_label_number_formats_parity() -> TestResult {
    let _ = Evidence::Queued;
    let name = "pages_body_chart_series_value_label_number_formats_parity";
    run(name, Tier::B, &[], &formats()?, None)?;
    run(name, Tier::B, &[], &formats()?, Some(0))?;
    Ok(())
}

#[test]
fn pages_body_chart_series_read_parity() -> TestResult {
    let (curved, _) = lines();
    for values in [
        curved,
        Values::Wedges(vec![10.0, 0.0, 25.0]),
        Values::Labels(vec![40.0, 55.0, 70.0]),
        Values::Trends(vec![Trend::Polynomial]),
        errors()?,
        Values::ValueLabels(vec![Visibility::Visible, Visibility::Hidden]),
        Values::Symbols(
            Kind::Line2d,
            vec![Some((SymbolShape::Diamond, Some(9.0))), None],
        ),
        Values::PieLabels(vec![
            LabelVisibility::ALL,
            LabelVisibility::HIDDEN,
            LabelVisibility::VALUES_ONLY,
        ]),
        affixes()?,
        formats()?,
    ] {
        let input = chart_input(&values, std::slice::from_ref(&values))?;
        let read = Package::from_bytes(&input.fixture.bytes)?
            .body_chart_series(last_chart(&input.fixture.bytes)?, values.kind())?;
        assert_eq!(read, values.leaf()?, "{values:?}");
    }
    let _ = (Tier::A, Evidence::Queued);
    Ok(())
}
