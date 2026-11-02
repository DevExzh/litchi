//! Numbers sheet-chart style parity (wave W2-A, record 0378): the host
//! `set_sheet_chart_*`/`sheet_chart_*` chart-level property routes against
//! `litchi_numbers::Package::edit_sheet_chart_style` (the W1-E chart style
//! engine wrapped by the Numbers family).

use litchi_iwa::charts::{
    Chart3dDepth, Chart3dRotation, Chart3dSeriesGap, ChartCornerRadius, ChartDonutInnerRadius,
    ChartPieStartAngle, ChartRadarStartAngle, ChartRoundedCorners,
};
use litchi_iwa::numbers;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};
use litchi_iwa_common::chart::ChartData;
use litchi_iwa_common::chart::kind::Kind;
use litchi_iwa_common::color::{RgbColorSpace, Rgba};
use litchi_iwa_common::shape::stroke::{Pattern, Stroke, Width};
use litchi_numbers::Package;
use litchi_numbers::sheet::chart::style::{
    Property, PropertyKind, RadarGridShape, chart3d, corner, donut, gaps, legend, pie, radar,
};

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::leaf_bytes;

const CHARTS: &str = "iwork/numbers/chart-data-native.numbers";
/// The host-added chart follows the fixture's native chart.
const CHART: usize = 1;

fn data_for(kind: Kind) -> TestResult<ChartData> {
    if kind.supports_pie_start_angle() {
        return Ok(ChartData::new(
            vec!["North".to_owned(), "South".to_owned(), "West".to_owned()],
            vec!["Revenue".to_owned()],
            vec![vec![Some(12.0)], vec![Some(18.0)], vec![Some(24.0)]],
        )?);
    }
    Ok(ChartData::new(
        vec!["Region 1".to_owned(), "Region 2".to_owned()],
        vec!["April".to_owned(), "May".to_owned(), "June".to_owned()],
        vec![
            vec![Some(17.0), Some(26.0), Some(53.0)],
            vec![Some(55.0), Some(43.0), Some(70.0)],
        ],
    )?)
}

/// `chart-data-native.numbers` with one host-added chart of `kind`, and
/// the host identifiers of the sheet and the chart.
fn input(kind: Kind) -> TestResult<(Fixture, u64, u64)> {
    let mut editor = numbers::Draft::from_bytes(&fixture(CHARTS)?.bytes)?;
    let sheet = editor.sheets()?.first().ok_or("no sheet")?.id();
    let chart = editor.add_sheet_chart(
        sheet,
        kind,
        data_for(kind)?,
        DrawablePoint { x: 120.0, y: 140.0 },
        DrawableSize {
            width: 360.0,
            height: 240.0,
        },
    )?;
    Ok((
        Fixture {
            path: "chart-data-native.numbers + add_sheet_chart",
            bytes: editor.to_bytes()?,
        },
        sheet,
        chart.drawable_object_id,
    ))
}

fn host_write(source: &[u8], sheet: u64, chart: u64, value: Property) -> TestResult<Vec<u8>> {
    let mut e = numbers::Draft::from_bytes(source)?;
    match value {
        Property::PieStartAngle(v) => e.set_sheet_chart_pie_start_angle(
            sheet,
            chart,
            ChartPieStartAngle::from_degrees(v.degrees())?,
        )?,
        Property::DonutInnerRadius(v) => e.set_sheet_chart_donut_inner_radius(
            sheet,
            chart,
            ChartDonutInnerRadius::from_percent(v.percent())?,
        )?,
        Property::RadarStartAngle(v) => e.set_sheet_chart_radar_start_angle(
            sheet,
            chart,
            ChartRadarStartAngle::from_degrees(v.degrees())?,
        )?,
        Property::SeriesGap3d(v) => {
            e.set_sheet_chart_3d_series_gap(sheet, chart, Chart3dSeriesGap::new(v.percent())?)?;
        },
        Property::Rotation3d(v) => e.set_sheet_chart_3d_rotation(
            sheet,
            chart,
            Chart3dRotation::from_degrees(v.x_degrees(), v.y_degrees())?,
        )?,
        Property::Depth3d(v) => {
            e.set_sheet_chart_3d_depth(sheet, chart, Chart3dDepth::from_percent(v.percent())?)?;
        },
        Property::RoundedCorners(v) => e.set_sheet_chart_rounded_corners(
            sheet,
            chart,
            ChartRoundedCorners::new(
                ChartCornerRadius::new(v.radius().percent())?,
                v.outside_corners_only(),
            ),
        )?,
        Property::LegendFrame(frame) => {
            use litchi_iwa::charts::{ChartLegendFrame, ChartLegendRect};
            let frame = match frame {
                legend::Frame::Automatic => ChartLegendFrame::Automatic,
                legend::Frame::Frame(rect) => {
                    let [x, y, width, height] = rect.native();
                    ChartLegendFrame::Frame(ChartLegendRect::from_points(x, y, width, height)?)
                },
            };
            e.set_sheet_chart_legend_frame(sheet, chart, frame)?;
        },
        Property::LegendVisible(v) => e.set_sheet_chart_legend_visible(sheet, chart, v)?,
        Property::BorderVisible(v) => e.set_sheet_chart_border_visible(sheet, chart, v)?,
        Property::IncludesHiddenData(v) => {
            e.set_sheet_chart_includes_hidden_data(sheet, chart, v)?;
        },
        Property::BarShape3d(v) => e.set_sheet_chart_3d_bar_shape(sheet, chart, v)?,
        Property::RadarGridShape(v) => e.set_sheet_chart_radar_grid_shape(
            sheet,
            chart,
            match v {
                RadarGridShape::Curved => litchi_iwa::charts::ChartRadarGridShape::Curved,
                _ => litchi_iwa::charts::ChartRadarGridShape::Straight,
            },
        )?,
        Property::GapSpacing(v) => e.set_sheet_chart_gap_spacing(sheet, chart, v)?,
        Property::BorderStroke(v) => e.set_sheet_chart_border_stroke(sheet, chart, v)?,
        _ => return Err("unmapped property".into()),
    }
    Ok(e.to_bytes()?)
}

fn host_read(bytes: &[u8], sheet: u64, chart: u64, kind: PropertyKind) -> TestResult<String> {
    let e = numbers::Draft::from_bytes(bytes)?;
    Ok(match kind {
        PropertyKind::PieStartAngle => {
            format!(
                "{:?}",
                e.sheet_chart_pie_start_angle(sheet, chart)?.degrees()
            )
        },
        PropertyKind::DonutInnerRadius => {
            format!(
                "{:?}",
                e.sheet_chart_donut_inner_radius(sheet, chart)?.percent()
            )
        },
        PropertyKind::RadarStartAngle => {
            format!(
                "{:?}",
                e.sheet_chart_radar_start_angle(sheet, chart)?.degrees()
            )
        },
        PropertyKind::SeriesGap3d => {
            format!("{:?}", e.sheet_chart_3d_series_gap(sheet, chart)?.percent())
        },
        PropertyKind::Rotation3d => {
            let rotation = e.sheet_chart_3d_rotation(sheet, chart)?;
            format!("{:?}", (rotation.x_degrees(), rotation.y_degrees()))
        },
        PropertyKind::Depth3d => format!("{:?}", e.sheet_chart_3d_depth(sheet, chart)?.percent()),
        PropertyKind::RoundedCorners => {
            let corners = e.sheet_chart_rounded_corners(sheet, chart)?;
            format!(
                "{:?}",
                (corners.radius().percent(), corners.outside_corners_only())
            )
        },
        PropertyKind::LegendFrame => {
            use litchi_iwa::charts::ChartLegendFrame;
            match e.sheet_chart_legend_frame(sheet, chart)? {
                ChartLegendFrame::Automatic => format!("{:?}", legend::Frame::Automatic),
                ChartLegendFrame::Frame(rect) => format!(
                    "{:?}",
                    legend::Frame::Frame(legend::Rect::from_points(
                        rect.x().points(),
                        rect.y().points(),
                        rect.width().points(),
                        rect.height().points(),
                    )?)
                ),
            }
        },
        PropertyKind::LegendVisible => format!("{:?}", e.sheet_chart_legend_visible(sheet, chart)?),
        PropertyKind::BorderVisible => format!("{:?}", e.sheet_chart_border_visible(sheet, chart)?),
        PropertyKind::IncludesHiddenData => {
            format!("{:?}", e.sheet_chart_includes_hidden_data(sheet, chart)?)
        },
        PropertyKind::BarShape3d => format!("{:?}", e.sheet_chart_3d_bar_shape(sheet, chart)?),
        PropertyKind::RadarGridShape => format!(
            "{:?}",
            e.sheet_chart_radar_grid_shape(sheet, chart)?
                == litchi_iwa::charts::ChartRadarGridShape::Curved
        ),
        PropertyKind::GapSpacing => format!("{:?}", e.sheet_chart_gap_spacing(sheet, chart)?),
        PropertyKind::BorderStroke => format!("{:?}", e.sheet_chart_border_stroke(sheet, chart)?),
        _ => return Err("unmapped kind".into()),
    })
}

/// The leaf read in the host read's rendering.
fn leaf_read(bytes: &[u8], kind: PropertyKind) -> TestResult<String> {
    let value = Package::from_bytes(bytes)?.sheet_chart_style(0, CHART, kind)?;
    Ok(match value {
        Property::PieStartAngle(v) => format!("{:?}", v.degrees()),
        Property::DonutInnerRadius(v) => format!("{:?}", v.percent()),
        Property::RadarStartAngle(v) => format!("{:?}", v.degrees()),
        Property::SeriesGap3d(v) => format!("{:?}", v.percent()),
        Property::Rotation3d(v) => format!("{:?}", (v.x_degrees(), v.y_degrees())),
        Property::Depth3d(v) => format!("{:?}", v.percent()),
        Property::RoundedCorners(v) => {
            format!("{:?}", (v.radius().percent(), v.outside_corners_only()))
        },
        Property::LegendFrame(frame) => format!("{frame:?}"),
        Property::LegendVisible(v)
        | Property::BorderVisible(v)
        | Property::IncludesHiddenData(v) => {
            format!("{v:?}")
        },
        Property::BarShape3d(v) => format!("{v:?}"),
        Property::RadarGridShape(v) => format!("{:?}", v == RadarGridShape::Curved),
        Property::GapSpacing(v) => format!("{v:?}"),
        Property::BorderStroke(v) => format!("{v:?}"),
        _ => return Err("unmapped property".into()),
    })
}

/// Set `value` on a host-added chart of `kind` (or reset it after a host
/// set) through both routes.
fn style_case(name: &'static str, kind: Kind, value: Property, reset: bool) -> TestResult {
    let (mut input, sheet, chart) = input(kind)?;
    let property = value.kind();
    if reset {
        input.bytes = host_write(&input.bytes, sheet, chart, value)?;
    }
    let target = if reset {
        property.native_default()
    } else {
        value
    };
    Case::new(name, Format::Numbers, Tier::A, Evidence::Queued, &input)
        .host(|source| host_write(source, sheet, chart, target))
        .leaf(|source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_sheet_chart_style(0, CHART)?;
            if reset {
                edit.reset(property)?;
            } else {
                edit.set(value)?;
            }
            let commit = edit.commit()?;
            let restored = commit
                .package()
                .apply_sheet_chart_style(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        })
        .read("host chart property", move |bytes| {
            host_read(bytes, sheet, chart, property)
        })
        .read("leaf sheet_chart_style", move |bytes| {
            leaf_read(bytes, property)
        })
        .run()?;
    Ok(())
}

fn cases() -> TestResult<Vec<(Kind, Property)>> {
    Ok(vec![
        (
            Kind::Pie2d,
            Property::PieStartAngle(pie::StartAngle::from_degrees(123.5)?),
        ),
        (
            Kind::Donut2d,
            Property::DonutInnerRadius(donut::InnerRadius::from_percent(40.0)?),
        ),
        (
            Kind::Radar2d,
            Property::RadarStartAngle(radar::StartAngle::from_degrees(45.0)?),
        ),
        (
            Kind::Line3d,
            Property::SeriesGap3d(chart3d::SeriesGap::new(150)?),
        ),
        (
            Kind::Column3d,
            Property::Rotation3d(chart3d::Rotation::from_degrees(12.0, -30.0)?),
        ),
        (
            Kind::Bar3d,
            Property::Depth3d(chart3d::Depth::from_percent(62.5)?),
        ),
        (
            Kind::Column2d,
            Property::RoundedCorners(corner::Rounding::new(corner::Radius::new(30.0)?, true)),
        ),
        (
            Kind::Column2d,
            Property::LegendFrame(legend::Frame::Frame(legend::Rect::from_points(
                24.0, 30.0, 140.0, 48.0,
            )?)),
        ),
    ])
}

#[test]
fn numbers_sheet_chart_style_set_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    for (kind, value) in cases()? {
        style_case("numbers_sheet_chart_style_set_parity", kind, value, false)
            .map_err(|error| format!("{kind:?} {value:?}: {error}"))?;
    }
    Ok(())
}

#[test]
fn numbers_sheet_chart_style_reset_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    for (kind, value) in cases()? {
        style_case("numbers_sheet_chart_style_reset_parity", kind, value, true)
            .map_err(|error| format!("{kind:?} {value:?}: {error}"))?;
    }
    Ok(())
}

#[test]
fn numbers_sheet_chart_style_read_parity() -> TestResult {
    for (kind, value) in cases()? {
        let (input, sheet, chart) = input(kind)?;
        for bytes in [
            input.bytes.clone(),
            host_write(&input.bytes, sheet, chart, value)?,
        ] {
            assert_eq!(
                host_read(&bytes, sheet, chart, value.kind())?,
                leaf_read(&bytes, value.kind())?,
                "{kind:?} {value:?}"
            );
        }
    }
    let _ = (Tier::A, Evidence::Queued);
    Ok(())
}

/// The chart-level option switches and values (legend, border, hidden data,
/// 3D bar shape, radar grid, gaps).
fn option_cases() -> TestResult<Vec<(Kind, Property)>> {
    Ok(vec![
        (Kind::Column2d, Property::LegendVisible(false)),
        (Kind::Column2d, Property::BorderVisible(true)),
        (Kind::Column2d, Property::IncludesHiddenData(false)),
        (
            Kind::Column3d,
            Property::BarShape3d(chart3d::BarShape::Cylinder),
        ),
        (
            Kind::Radar2d,
            Property::RadarGridShape(RadarGridShape::Curved),
        ),
        (
            Kind::Column2d,
            Property::BorderStroke(Some(Stroke::new(
                Rgba::new(0.8, 0.1, 0.1, 1.0, RgbColorSpace::Srgb)?,
                Width::new(3.0)?,
                Pattern::RoundedDash,
            ))),
        ),
        (Kind::Column2d, Property::BorderStroke(None)),
        (
            Kind::Column2d,
            Property::GapSpacing(gaps::Spacing::new(
                gaps::Percentage::new(25.0)?,
                gaps::Percentage::new(80.0)?,
            )),
        ),
    ])
}

#[test]
fn numbers_sheet_chart_options_set_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    for (kind, value) in option_cases()? {
        style_case("numbers_sheet_chart_options_set_parity", kind, value, false)
            .map_err(|error| format!("{kind:?} {value:?}: {error}"))?;
        // A host-added chart shows its legend, the native default hides it:
        // the reset of a hidden legend is a no-op.
        if value.kind() != PropertyKind::LegendVisible {
            style_case("numbers_sheet_chart_options_set_parity", kind, value, true)
                .map_err(|error| format!("{kind:?} {value:?} reset: {error}"))?;
        }
    }
    Ok(())
}

#[test]
fn numbers_sheet_chart_options_read_parity() -> TestResult {
    for (kind, value) in option_cases()? {
        let (input, sheet, chart) = input(kind)?;
        for bytes in [
            input.bytes.clone(),
            host_write(&input.bytes, sheet, chart, value)?,
        ] {
            assert_eq!(
                host_read(&bytes, sheet, chart, value.kind())?,
                leaf_read(&bytes, value.kind())?,
                "{kind:?} {value:?}"
            );
        }
    }
    let _ = (Tier::A, Evidence::Queued);
    Ok(())
}
