//! Paired host-vs-leaf cost probe for the wave W4-C chart property families
//! (record 0414): each route is open + one operation (+ commit + serialize
//! for writes) on a chart the host added to the native chart fixture, run
//! ABBA (host, leaf, leaf, host). Run under callgrind and read the
//! inclusive cost of `w4c_host_route` and `w4c_leaf_route`:
//!
//! ```sh
//! W4C_ROUTE=keynote_style_set taskset -c 3 valgrind --tool=callgrind \
//!   target/release/deps/parity-<hash> --ignored --exact \
//!   chart_property_cost::chart_property_route_cost_abba --test-threads=1
//! callgrind_annotate --inclusive=yes callgrind.out.<pid> | grep w4c_
//! ```

use std::hint::black_box;

use litchi_iwa::charts::{ChartPieStartAngle, ChartSeriesConnectionLine};
use litchi_iwa::shapes::ShapeFill;
use litchi_iwa::{keynote, numbers, pages};
use litchi_iwa_common::chart::axis::Axis;
use litchi_iwa_common::chart::axis::style::Visibility;
use litchi_iwa_common::chart::kind::Kind;
use litchi_iwa_common::chart::number_format::LabelAffixes;
use litchi_iwa_common::chart::pie::StartAngle;
use litchi_iwa_common::chart::series::ConnectionLine;
use litchi_iwa_common::color::{RgbColorSpace, Rgba};
use litchi_iwa_common::shape::fill::Fill;

use crate::chart_style_cases::{Input, Target, input};
use crate::harness::{Format, TestResult};
use crate::leaf_bytes;

/// One measured operation.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Operation {
    StyleSet,
    StyleRead,
    AxisSet,
    SeriesSet,
    SeriesRead,
    SeriesFillSet,
    AffixesSet,
}

const ROUTES: &[(&str, Format, Operation)] = &[
    ("keynote_style_set", Format::Keynote, Operation::StyleSet),
    ("keynote_style_read", Format::Keynote, Operation::StyleRead),
    ("keynote_axis_set", Format::Keynote, Operation::AxisSet),
    ("keynote_series_set", Format::Keynote, Operation::SeriesSet),
    (
        "keynote_series_read",
        Format::Keynote,
        Operation::SeriesRead,
    ),
    (
        "keynote_series_fill_set",
        Format::Keynote,
        Operation::SeriesFillSet,
    ),
    ("pages_style_set", Format::Pages, Operation::StyleSet),
    ("pages_style_read", Format::Pages, Operation::StyleRead),
    ("pages_axis_set", Format::Pages, Operation::AxisSet),
    ("pages_series_set", Format::Pages, Operation::SeriesSet),
    ("pages_series_read", Format::Pages, Operation::SeriesRead),
    (
        "pages_series_fill_set",
        Format::Pages,
        Operation::SeriesFillSet,
    ),
    (
        "numbers_affixes_set",
        Format::Numbers,
        Operation::AffixesSet,
    ),
];

const fn chart_kind(operation: Operation) -> Kind {
    match operation {
        Operation::StyleSet | Operation::StyleRead => Kind::Pie2d,
        Operation::SeriesSet | Operation::SeriesRead => Kind::Line2d,
        Operation::AxisSet | Operation::SeriesFillSet | Operation::AffixesSet => Kind::Column2d,
    }
}

fn fill() -> TestResult<Rgba> {
    Ok(Rgba::new(0.1, 0.4, 0.9, 1.0, RgbColorSpace::Srgb)?)
}

fn affixes() -> TestResult<Vec<LabelAffixes>> {
    Ok(vec![
        LabelAffixes::new("$", " M")?,
        LabelAffixes::new("", "%")?,
    ])
}

fn host_operation(input: &Input, source: &[u8], operation: Operation) -> TestResult<Vec<u8>> {
    let chart = input.chart;
    let angle = || ChartPieStartAngle::from_degrees(123.5);
    let lines = [
        ChartSeriesConnectionLine::Curved,
        ChartSeriesConnectionLine::Straight,
    ];
    let fills = || -> TestResult<Vec<ShapeFill>> {
        Ok(vec![ShapeFill::Solid(fill()?), ShapeFill::Solid(fill()?)])
    };
    Ok(match input.target {
        Target::Keynote { slide } => {
            let mut e = keynote::Draft::from_bytes(source)?;
            match operation {
                Operation::StyleSet => e.set_slide_chart_pie_start_angle(slide, chart, angle()?)?,
                Operation::StyleRead => {
                    return Ok(
                        format!("{:?}", e.slide_chart_pie_start_angle(slide, chart)?).into_bytes(),
                    );
                },
                Operation::AxisSet => e.set_slide_chart_axis_major_gridlines_visible(
                    slide,
                    chart,
                    Axis::Value,
                    Visibility::Visible,
                )?,
                Operation::SeriesSet => {
                    e.set_slide_chart_series_connection_lines(slide, chart, &lines)?;
                },
                Operation::SeriesRead => {
                    return Ok(format!(
                        "{:?}",
                        e.slide_chart_series_connection_lines(slide, chart)?
                    )
                    .into_bytes());
                },
                Operation::SeriesFillSet => {
                    e.set_slide_chart_series_fills(slide, chart, &fills()?)?;
                },
                Operation::AffixesSet => return Err("not a Keynote route".into()),
            }
            e.to_bytes()?
        },
        Target::Pages => {
            let mut e = pages::Draft::from_bytes(source)?;
            match operation {
                Operation::StyleSet => e.set_body_chart_pie_start_angle(chart, angle()?)?,
                Operation::StyleRead => {
                    return Ok(format!("{:?}", e.body_chart_pie_start_angle(chart)?).into_bytes());
                },
                Operation::AxisSet => e.set_body_chart_axis_major_gridlines_visible(
                    chart,
                    Axis::Value,
                    Visibility::Visible,
                )?,
                Operation::SeriesSet => e.set_body_chart_series_connection_lines(chart, &lines)?,
                Operation::SeriesRead => {
                    return Ok(
                        format!("{:?}", e.body_chart_series_connection_lines(chart)?).into_bytes(),
                    );
                },
                Operation::SeriesFillSet => e.set_body_chart_series_fills(chart, &fills()?)?,
                Operation::AffixesSet => return Err("not a Pages route".into()),
            }
            e.to_bytes()?
        },
        Target::Numbers { sheet } => {
            let mut e = numbers::Draft::from_bytes(source)?;
            match operation {
                Operation::AffixesSet => {
                    e.set_sheet_chart_series_value_label_affixes(sheet, chart, &affixes()?)?;
                },
                _ => return Err("not a Numbers route".into()),
            }
            e.to_bytes()?
        },
    })
}

fn keynote_operation(source: &[u8], operation: Operation) -> TestResult<Vec<u8>> {
    use litchi_keynote::chart::{axis, series, style};
    let package = litchi_keynote::Package::from_bytes(source)?;
    let (slide, chart) = (0, 1);
    Ok(match operation {
        Operation::StyleSet => {
            let mut edit = package.edit_slide_chart_style(slide, chart)?;
            edit.set(style::Property::PieStartAngle(StartAngle::from_degrees(
                123.5,
            )?))?;
            leaf_bytes!(edit.commit()?.package())
        },
        Operation::StyleRead => format!(
            "{:?}",
            package.slide_chart_style(slide, chart, style::PropertyKind::PieStartAngle)?
        )
        .into_bytes(),
        Operation::AxisSet => {
            let mut edit = package.edit_slide_chart_axis(slide, chart)?;
            edit.set(
                Axis::Value,
                axis::Property::MajorGridlines(Visibility::Visible),
            )?;
            leaf_bytes!(edit.commit()?.package())
        },
        Operation::SeriesSet => {
            let mut edit = package.edit_slide_chart_series(slide, chart)?;
            edit.set(series::Property::ConnectionLines(vec![
                ConnectionLine::Curved,
                ConnectionLine::Straight,
            ]))?;
            leaf_bytes!(edit.commit()?.package())
        },
        Operation::SeriesRead => format!(
            "{:?}",
            package.slide_chart_series(slide, chart, series::PropertyKind::ConnectionLines)?
        )
        .into_bytes(),
        Operation::SeriesFillSet => {
            let mut edit = package.edit_slide_chart_series(slide, chart)?;
            edit.set(series::Property::Fills(vec![
                Fill::Solid(fill()?),
                Fill::Solid(fill()?),
            ]))?;
            leaf_bytes!(edit.commit()?.package())
        },
        Operation::AffixesSet => return Err("not a Keynote route".into()),
    })
}

fn pages_operation(source: &[u8], operation: Operation) -> TestResult<Vec<u8>> {
    use litchi_pages::chart::{axis, series, style};
    let package = litchi_pages::Package::from_bytes(source)?;
    let chart = 0;
    Ok(match operation {
        Operation::StyleSet => leaf_bytes!(
            package
                .edit_body_chart_style(chart)?
                .set(style::Property::PieStartAngle(StartAngle::from_degrees(
                    123.5
                )?))?
                .commit()?
                .package()
        ),
        Operation::StyleRead => format!(
            "{:?}",
            package.body_chart_style(chart, style::PropertyKind::PieStartAngle)?
        )
        .into_bytes(),
        Operation::AxisSet => leaf_bytes!(
            package
                .edit_body_chart_axis(chart)?
                .set(
                    Axis::Value,
                    axis::Property::MajorGridlines(Visibility::Visible)
                )?
                .commit()?
                .package()
        ),
        Operation::SeriesSet => leaf_bytes!(
            package
                .edit_body_chart_series(chart)?
                .set(series::Property::ConnectionLines(vec![
                    ConnectionLine::Curved,
                    ConnectionLine::Straight,
                ]))?
                .commit()?
                .package()
        ),
        Operation::SeriesRead => format!(
            "{:?}",
            package.body_chart_series(chart, series::PropertyKind::ConnectionLines)?
        )
        .into_bytes(),
        Operation::SeriesFillSet => leaf_bytes!(
            package
                .edit_body_chart_series(chart)?
                .set(series::Property::Fills(vec![
                    Fill::Solid(fill()?),
                    Fill::Solid(fill()?),
                ]))?
                .commit()?
                .package()
        ),
        Operation::AffixesSet => return Err("not a Pages route".into()),
    })
}

fn numbers_operation(source: &[u8], operation: Operation) -> TestResult<Vec<u8>> {
    use litchi_numbers::sheet::chart::series;
    let package = litchi_numbers::Package::from_bytes(source)?;
    match operation {
        Operation::AffixesSet => {
            let mut edit = package.edit_sheet_chart_series(0, 1)?;
            edit.set(series::Property::ValueLabelAffixes(affixes()?))?;
            Ok(leaf_bytes!(edit.commit()?.package()))
        },
        _ => Err("not a Numbers route".into()),
    }
}

#[inline(never)]
fn w4c_host_route(input: &Input, source: &[u8], operation: Operation) -> TestResult<Vec<u8>> {
    host_operation(input, source, operation)
}

#[inline(never)]
fn w4c_leaf_route(format: Format, source: &[u8], operation: Operation) -> TestResult<Vec<u8>> {
    match format {
        Format::Keynote => keynote_operation(source, operation),
        Format::Pages => pages_operation(source, operation),
        Format::Numbers => numbers_operation(source, operation),
    }
}

/// Every route agrees with itself outside the measurement (the parity
/// cases prove host/leaf equality); `W4C_ROUTE` selects one route.
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn chart_property_route_cost_abba() -> TestResult {
    let selected = std::env::var("W4C_ROUTE").ok();
    for &(name, format, operation) in ROUTES {
        if selected.as_deref().is_some_and(|route| route != name) {
            continue;
        }
        let input = input(format, chart_kind(operation))?;
        let source = input.fixture.bytes.clone();
        black_box(w4c_host_route(&input, &source, operation)?);
        black_box(w4c_leaf_route(format, &source, operation)?);
        black_box(w4c_leaf_route(format, &source, operation)?);
        black_box(w4c_host_route(&input, &source, operation)?);
    }
    Ok(())
}
