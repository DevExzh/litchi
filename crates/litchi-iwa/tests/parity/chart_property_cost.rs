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

use litchi_iwa::charts::{
    Chart3dLightingStyle, ChartAxisGridline, ChartAxisGridlineStroke, ChartFont, ChartLegendFill,
    ChartLegendFont, ChartPieStartAngle, ChartSeriesConnectionLine, ChartSeriesValueLabelLocation,
    ChartShadow,
};
use litchi_iwa::shapes::ShapeFill;
use litchi_iwa::text::TextFont;
use litchi_iwa::{keynote, numbers, pages};
use litchi_iwa_common::chart::axis::Axis;
use litchi_iwa_common::chart::axis::style::Visibility;
use litchi_iwa_common::chart::category_labels::{Frequency, Interval, Layout};
use litchi_iwa_common::chart::kind::Kind;
use litchi_iwa_common::chart::number_format::{
    DecimalPlaces, LabelAffixes, NegativeStyle, NumberFormat,
};
use litchi_iwa_common::chart::pie::LeaderLineVisibility;
use litchi_iwa_common::chart::pie::StartAngle;
use litchi_iwa_common::chart::series::ConnectionLine;
use litchi_iwa_common::color::{RgbColorSpace, Rgba};
use litchi_iwa_common::shape::fill::Fill;
use litchi_iwa_common::shape::shadow::{
    Angle as ShadowAngle, Appearance, BlurRadius, Drop, Offset, Opacity,
};
use litchi_iwa_common::shape::stroke::{Pattern, Stroke, Width};
use litchi_iwa_style::chart_style::axis::{GridlineStroke, Property as AxisProperty};
use litchi_iwa_style::chart_style::chart::{Property as ChartProperty, Shadow as LeafShadow};
use litchi_iwa_style::legend_style::{LegendFill, LegendFont, Property as LegendProperty};
use litchi_iwa_style::series::style::ValueLabelLocation;

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
    AxisFormatSet,
    GridlineStrokeSet,
    BorderSet,
    LegendFillSet,
    LegendFontSet,
    LeaderLinesSet,
    LabelLocationsSet,
    BackgroundSet,
    LightingSet,
    ShadowSet,
    CategoryLabelsSet,
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
    (
        "keynote_axis_format_set",
        Format::Keynote,
        Operation::AxisFormatSet,
    ),
    (
        "keynote_gridline_stroke_set",
        Format::Keynote,
        Operation::GridlineStrokeSet,
    ),
    ("keynote_border_set", Format::Keynote, Operation::BorderSet),
    (
        "pages_axis_format_set",
        Format::Pages,
        Operation::AxisFormatSet,
    ),
    (
        "pages_gridline_stroke_set",
        Format::Pages,
        Operation::GridlineStrokeSet,
    ),
    ("pages_border_set", Format::Pages, Operation::BorderSet),
    (
        "numbers_axis_format_set",
        Format::Numbers,
        Operation::AxisFormatSet,
    ),
    (
        "numbers_gridline_stroke_set",
        Format::Numbers,
        Operation::GridlineStrokeSet,
    ),
    ("numbers_border_set", Format::Numbers, Operation::BorderSet),
    (
        "keynote_legend_fill_set",
        Format::Keynote,
        Operation::LegendFillSet,
    ),
    (
        "pages_legend_fill_set",
        Format::Pages,
        Operation::LegendFillSet,
    ),
    (
        "numbers_legend_fill_set",
        Format::Numbers,
        Operation::LegendFillSet,
    ),
    (
        "keynote_background_set",
        Format::Keynote,
        Operation::BackgroundSet,
    ),
    (
        "keynote_lighting_set",
        Format::Keynote,
        Operation::LightingSet,
    ),
    (
        "pages_background_set",
        Format::Pages,
        Operation::BackgroundSet,
    ),
    ("pages_lighting_set", Format::Pages, Operation::LightingSet),
    (
        "numbers_background_set",
        Format::Numbers,
        Operation::BackgroundSet,
    ),
    (
        "numbers_lighting_set",
        Format::Numbers,
        Operation::LightingSet,
    ),
    ("keynote_shadow_set", Format::Keynote, Operation::ShadowSet),
    (
        "keynote_category_labels_set",
        Format::Keynote,
        Operation::CategoryLabelsSet,
    ),
    ("pages_shadow_set", Format::Pages, Operation::ShadowSet),
    (
        "pages_category_labels_set",
        Format::Pages,
        Operation::CategoryLabelsSet,
    ),
    ("numbers_shadow_set", Format::Numbers, Operation::ShadowSet),
    (
        "numbers_category_labels_set",
        Format::Numbers,
        Operation::CategoryLabelsSet,
    ),
    (
        "keynote_legend_font_set",
        Format::Keynote,
        Operation::LegendFontSet,
    ),
    (
        "pages_legend_font_set",
        Format::Pages,
        Operation::LegendFontSet,
    ),
    (
        "numbers_legend_font_set",
        Format::Numbers,
        Operation::LegendFontSet,
    ),
    (
        "keynote_leader_lines_set",
        Format::Keynote,
        Operation::LeaderLinesSet,
    ),
    (
        "keynote_label_locations_set",
        Format::Keynote,
        Operation::LabelLocationsSet,
    ),
    (
        "pages_leader_lines_set",
        Format::Pages,
        Operation::LeaderLinesSet,
    ),
    (
        "pages_label_locations_set",
        Format::Pages,
        Operation::LabelLocationsSet,
    ),
    (
        "numbers_leader_lines_set",
        Format::Numbers,
        Operation::LeaderLinesSet,
    ),
    (
        "numbers_label_locations_set",
        Format::Numbers,
        Operation::LabelLocationsSet,
    ),
];

const fn chart_kind(operation: Operation) -> Kind {
    match operation {
        Operation::StyleSet | Operation::StyleRead | Operation::LeaderLinesSet => Kind::Pie2d,
        Operation::SeriesSet | Operation::SeriesRead => Kind::Line2d,
        Operation::AxisSet
        | Operation::SeriesFillSet
        | Operation::AffixesSet
        | Operation::AxisFormatSet
        | Operation::GridlineStrokeSet
        | Operation::BorderSet
        | Operation::LegendFillSet
        | Operation::LegendFontSet
        | Operation::LabelLocationsSet
        | Operation::BackgroundSet => Kind::Column2d,
        Operation::LightingSet => Kind::Column3d,
        Operation::ShadowSet | Operation::CategoryLabelsSet => Kind::Column2d,
    }
}

fn fill() -> TestResult<Rgba> {
    Ok(Rgba::new(0.1, 0.4, 0.9, 1.0, RgbColorSpace::Srgb)?)
}

fn drop_shadow() -> TestResult<Drop> {
    Ok(Drop::new(
        Appearance::new(
            fill()?,
            BlurRadius::from_points(4)?,
            Offset::from_points(3.0)?,
            Opacity::new(0.5)?,
        ),
        ShadowAngle::from_degrees(120.0)?,
    ))
}

fn layout() -> TestResult<Layout> {
    Ok(Layout::new(Frequency::Every(Interval::new(3)?), false))
}

fn shadow_leaf() -> TestResult<ChartProperty> {
    Ok(ChartProperty::Shadow(LeafShadow::Grouped(drop_shadow()?)))
}

fn label_format() -> TestResult<NumberFormat> {
    Ok(NumberFormat::new(
        DecimalPlaces::fixed(2)?,
        NegativeStyle::Parentheses,
        true,
    ))
}

fn stroke() -> TestResult<Stroke> {
    Ok(Stroke::new(fill()?, Width::new(2.0)?, Pattern::Solid))
}

fn legend_fill_host() -> TestResult<ChartLegendFill> {
    Ok(ChartLegendFill::Fill(ShapeFill::Solid(fill()?)))
}

fn legend_font_host() -> TestResult<ChartLegendFont> {
    Ok(ChartLegendFont::Font(
        ChartFont::new(TextFont::named("Georgia")?).with_bold(true),
    ))
}

fn leaders() -> Vec<LeaderLineVisibility> {
    vec![
        LeaderLineVisibility::Hidden,
        LeaderLineVisibility::Visible,
        LeaderLineVisibility::Hidden,
    ]
}

fn locations() -> Vec<ValueLabelLocation> {
    vec![ValueLabelLocation::Outside, ValueLabelLocation::Middle]
}

fn host_locations() -> Vec<ChartSeriesValueLabelLocation> {
    vec![
        ChartSeriesValueLabelLocation::Outside,
        ChartSeriesValueLabelLocation::Middle,
    ]
}

fn legend_font_leaf() -> TestResult<LegendProperty> {
    Ok(LegendProperty::Font(LegendFont::Font {
        font: TextFont::named("Georgia")?,
        bold: true,
        italic: false,
    }))
}

fn legend_fill_leaf() -> TestResult<LegendProperty> {
    Ok(LegendProperty::Fill(LegendFill::Fill(Fill::Solid(fill()?))))
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
                Operation::AxisFormatSet => {
                    e.set_slide_chart_axis_number_format(
                        slide,
                        chart,
                        Axis::Value,
                        label_format()?,
                    )?;
                },
                Operation::GridlineStrokeSet => e.set_slide_chart_axis_gridline_stroke(
                    slide,
                    chart,
                    Axis::Value,
                    ChartAxisGridline::Major,
                    ChartAxisGridlineStroke::Stroke(stroke()?),
                )?,
                Operation::BorderSet => {
                    e.set_slide_chart_border_stroke(slide, chart, Some(stroke()?))?;
                },
                Operation::BackgroundSet => {
                    e.set_slide_chart_background_fill(slide, chart, &ShapeFill::Solid(fill()?))?;
                },
                Operation::LightingSet => {
                    e.set_slide_chart_3d_lighting_style(slide, chart, Chart3dLightingStyle::Glossy)?
                },
                Operation::ShadowSet => {
                    e.set_slide_chart_shadow(slide, chart, ChartShadow::Grouped(drop_shadow()?))?;
                },
                Operation::CategoryLabelsSet => {
                    e.set_slide_chart_category_label_layout(slide, chart, layout()?)?;
                },
                Operation::LegendFillSet => {
                    e.set_slide_chart_legend_fill(slide, chart, &legend_fill_host()?)?;
                },
                Operation::LegendFontSet => {
                    e.set_slide_chart_legend_font(slide, chart, &legend_font_host()?)?;
                },
                Operation::LeaderLinesSet => {
                    e.set_slide_chart_pie_leader_line_visibilities(slide, chart, &leaders())?;
                },
                Operation::LabelLocationsSet => {
                    e.set_slide_chart_series_value_label_locations(
                        slide,
                        chart,
                        &host_locations(),
                    )?;
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
                Operation::AxisFormatSet => {
                    e.set_body_chart_axis_number_format(chart, Axis::Value, label_format()?)?;
                },
                Operation::GridlineStrokeSet => e.set_body_chart_axis_gridline_stroke(
                    chart,
                    Axis::Value,
                    ChartAxisGridline::Major,
                    ChartAxisGridlineStroke::Stroke(stroke()?),
                )?,
                Operation::BorderSet => e.set_body_chart_border_stroke(chart, Some(stroke()?))?,
                Operation::BackgroundSet => {
                    e.set_body_chart_background_fill(chart, &ShapeFill::Solid(fill()?))?;
                },
                Operation::LightingSet => {
                    e.set_body_chart_3d_lighting_style(chart, Chart3dLightingStyle::Glossy)?;
                },
                Operation::ShadowSet => {
                    e.set_body_chart_shadow(chart, ChartShadow::Grouped(drop_shadow()?))?;
                },
                Operation::CategoryLabelsSet => {
                    e.set_body_chart_category_label_layout(chart, layout()?)?;
                },
                Operation::LegendFillSet => {
                    e.set_body_chart_legend_fill(chart, &legend_fill_host()?)?
                },
                Operation::LegendFontSet => {
                    e.set_body_chart_legend_font(chart, &legend_font_host()?)?;
                },
                Operation::LeaderLinesSet => {
                    e.set_body_chart_pie_leader_line_visibilities(chart, &leaders())?;
                },
                Operation::LabelLocationsSet => {
                    e.set_body_chart_series_value_label_locations(chart, &host_locations())?;
                },
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
                Operation::AxisFormatSet => {
                    e.set_sheet_chart_axis_number_format(
                        sheet,
                        chart,
                        Axis::Value,
                        label_format()?,
                    )?;
                },
                Operation::GridlineStrokeSet => e.set_sheet_chart_axis_gridline_stroke(
                    sheet,
                    chart,
                    Axis::Value,
                    ChartAxisGridline::Major,
                    ChartAxisGridlineStroke::Stroke(stroke()?),
                )?,
                Operation::BorderSet => {
                    e.set_sheet_chart_border_stroke(sheet, chart, Some(stroke()?))?;
                },
                Operation::BackgroundSet => {
                    e.set_sheet_chart_background_fill(sheet, chart, &ShapeFill::Solid(fill()?))?;
                },
                Operation::LightingSet => {
                    e.set_sheet_chart_3d_lighting_style(sheet, chart, Chart3dLightingStyle::Glossy)?
                },
                Operation::ShadowSet => {
                    e.set_sheet_chart_shadow(sheet, chart, ChartShadow::Grouped(drop_shadow()?))?;
                },
                Operation::CategoryLabelsSet => {
                    e.set_sheet_chart_category_label_layout(sheet, chart, layout()?)?;
                },
                Operation::LegendFillSet => {
                    e.set_sheet_chart_legend_fill(sheet, chart, &legend_fill_host()?)?;
                },
                Operation::LegendFontSet => {
                    e.set_sheet_chart_legend_font(sheet, chart, &legend_font_host()?)?;
                },
                Operation::LeaderLinesSet => {
                    e.set_sheet_chart_pie_leader_line_visibilities(sheet, chart, &leaders())?;
                },
                Operation::LabelLocationsSet => {
                    e.set_sheet_chart_series_value_label_locations(
                        sheet,
                        chart,
                        &host_locations(),
                    )?;
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
        Operation::AxisFormatSet | Operation::GridlineStrokeSet => {
            let mut edit = package.edit_slide_chart_axis(slide, chart)?;
            edit.set(Axis::Value, axis_value(operation)?)?;
            leaf_bytes!(edit.commit()?.package())
        },
        Operation::BorderSet => {
            let mut edit = package.edit_slide_chart_style(slide, chart)?;
            edit.set(style::Property::BorderStroke(Some(stroke()?)))?;
            leaf_bytes!(edit.commit()?.package())
        },
        Operation::BackgroundSet => {
            let mut edit = package.edit_slide_chart_style(slide, chart)?;
            edit.set(style::Property::BackgroundFill(Fill::Solid(fill()?)))?;
            leaf_bytes!(edit.commit()?.package())
        },
        Operation::LightingSet => {
            let mut edit = package.edit_slide_chart_style(slide, chart)?;
            edit.set(style::Property::Lighting3d(style::Lighting3d::Glossy))?;
            leaf_bytes!(edit.commit()?.package())
        },
        Operation::ShadowSet => {
            let mut edit = package.edit_slide_chart_style(slide, chart)?;
            edit.set(shadow_leaf()?)?;
            leaf_bytes!(edit.commit()?.package())
        },
        Operation::CategoryLabelsSet => {
            let mut edit = package.edit_slide_chart_axis(slide, chart)?;
            edit.set(Axis::Category, AxisProperty::CategoryLabels(layout()?))?;
            leaf_bytes!(edit.commit()?.package())
        },
        Operation::LegendFillSet => {
            let mut edit = package.edit_slide_chart_legend_style(slide, chart)?;
            edit.set(legend_fill_leaf()?)?;
            leaf_bytes!(edit.commit()?.package())
        },
        Operation::LegendFontSet => {
            let mut edit = package.edit_slide_chart_legend_style(slide, chart)?;
            edit.set(legend_font_leaf()?)?;
            leaf_bytes!(edit.commit()?.package())
        },
        Operation::LeaderLinesSet => {
            let mut edit = package.edit_slide_chart_series(slide, chart)?;
            edit.set(series::Property::LeaderLines(leaders()))?;
            leaf_bytes!(edit.commit()?.package())
        },
        Operation::LabelLocationsSet => {
            let mut edit = package.edit_slide_chart_series(slide, chart)?;
            edit.set(series::Property::ValueLabelLocations(locations()))?;
            leaf_bytes!(edit.commit()?.package())
        },
        Operation::AffixesSet => return Err("not a Keynote route".into()),
    })
}

fn axis_value(operation: Operation) -> TestResult<AxisProperty> {
    Ok(match operation {
        Operation::AxisFormatSet => AxisProperty::NumberFormat(label_format()?),
        _ => AxisProperty::MajorGridlineStroke(GridlineStroke::Stroke(stroke()?)),
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
        Operation::AxisFormatSet | Operation::GridlineStrokeSet => leaf_bytes!(
            package
                .edit_body_chart_axis(chart)?
                .set(Axis::Value, axis_value(operation)?)?
                .commit()?
                .package()
        ),
        Operation::BorderSet => leaf_bytes!(
            package
                .edit_body_chart_style(chart)?
                .set(style::Property::BorderStroke(Some(stroke()?)))?
                .commit()?
                .package()
        ),
        Operation::BackgroundSet => leaf_bytes!(
            package
                .edit_body_chart_style(chart)?
                .set(style::Property::BackgroundFill(Fill::Solid(fill()?)))?
                .commit()?
                .package()
        ),
        Operation::LightingSet => leaf_bytes!(
            package
                .edit_body_chart_style(chart)?
                .set(style::Property::Lighting3d(style::Lighting3d::Glossy))?
                .commit()?
                .package()
        ),
        Operation::ShadowSet => leaf_bytes!(
            package
                .edit_body_chart_style(chart)?
                .set(shadow_leaf()?)?
                .commit()?
                .package()
        ),
        Operation::CategoryLabelsSet => leaf_bytes!(
            package
                .edit_body_chart_axis(chart)?
                .set(Axis::Category, AxisProperty::CategoryLabels(layout()?))?
                .commit()?
                .package()
        ),
        Operation::LegendFillSet => leaf_bytes!(
            package
                .edit_body_chart_legend_style(chart)?
                .set(legend_fill_leaf()?)?
                .commit()?
                .package()
        ),
        Operation::LegendFontSet => leaf_bytes!(
            package
                .edit_body_chart_legend_style(chart)?
                .set(legend_font_leaf()?)?
                .commit()?
                .package()
        ),
        Operation::LeaderLinesSet => leaf_bytes!(
            package
                .edit_body_chart_series(chart)?
                .set(series::Property::LeaderLines(leaders()))?
                .commit()?
                .package()
        ),
        Operation::LabelLocationsSet => leaf_bytes!(
            package
                .edit_body_chart_series(chart)?
                .set(series::Property::ValueLabelLocations(locations()))?
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
        Operation::AxisFormatSet | Operation::GridlineStrokeSet => {
            let mut edit = package.edit_sheet_chart_axis(0, 1)?;
            edit.set(Axis::Value, axis_value(operation)?)?;
            Ok(leaf_bytes!(edit.commit()?.package()))
        },
        Operation::BorderSet => {
            let mut edit = package.edit_sheet_chart_style(0, 1)?;
            edit.set(litchi_numbers::sheet::chart::style::Property::BorderStroke(
                Some(stroke()?),
            ))?;
            Ok(leaf_bytes!(edit.commit()?.package()))
        },
        Operation::BackgroundSet => {
            let mut edit = package.edit_sheet_chart_style(0, 1)?;
            edit.set(
                litchi_numbers::sheet::chart::style::Property::BackgroundFill(Fill::Solid(fill()?)),
            )?;
            Ok(leaf_bytes!(edit.commit()?.package()))
        },
        Operation::LightingSet => {
            use litchi_numbers::sheet::chart::style::{Lighting3d, Property};
            let mut edit = package.edit_sheet_chart_style(0, 1)?;
            edit.set(Property::Lighting3d(Lighting3d::Glossy))?;
            Ok(leaf_bytes!(edit.commit()?.package()))
        },
        Operation::ShadowSet => {
            let mut edit = package.edit_sheet_chart_style(0, 1)?;
            edit.set(shadow_leaf()?)?;
            Ok(leaf_bytes!(edit.commit()?.package()))
        },
        Operation::CategoryLabelsSet => {
            let mut edit = package.edit_sheet_chart_axis(0, 1)?;
            edit.set(Axis::Category, AxisProperty::CategoryLabels(layout()?))?;
            Ok(leaf_bytes!(edit.commit()?.package()))
        },
        Operation::LegendFillSet => {
            let mut edit = package.edit_sheet_chart_legend_style(0, 1)?;
            edit.set(legend_fill_leaf()?)?;
            Ok(leaf_bytes!(edit.commit()?.package()))
        },
        Operation::LegendFontSet => {
            let mut edit = package.edit_sheet_chart_legend_style(0, 1)?;
            edit.set(legend_font_leaf()?)?;
            Ok(leaf_bytes!(edit.commit()?.package()))
        },
        Operation::LeaderLinesSet => {
            let mut edit = package.edit_sheet_chart_series(0, 1)?;
            edit.set(series::Property::LeaderLines(leaders()))?;
            Ok(leaf_bytes!(edit.commit()?.package()))
        },
        Operation::LabelLocationsSet => {
            let mut edit = package.edit_sheet_chart_series(0, 1)?;
            edit.set(series::Property::ValueLabelLocations(locations()))?;
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
