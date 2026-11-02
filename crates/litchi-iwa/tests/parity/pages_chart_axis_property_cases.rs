//! Pages body chart axis-style parity (wave W4-C, record 0414): the host
//! `set_body_chart_axis_{line,major_gridlines,minor_gridlines,
//! minor_tick_marks}_visible`, `set_body_chart_axis_tick_mark_location`
//! and `set_body_chart_value_axis_minimum_label_visible` routes against
//! `litchi_pages::chart::axis::Edit::set`.

use litchi_iwa::charts::{ChartAxisGridline, ChartAxisGridlineStroke};
use litchi_iwa::pages;
use litchi_iwa_common::chart::axis::label_angle::LabelAngle;
use litchi_iwa_common::chart::axis::label_position_3d::LabelPosition3d;
use litchi_iwa_common::chart::axis::style::Visibility;
use litchi_iwa_common::chart::axis::{Axis, TickMarkLocation};
use litchi_iwa_common::chart::category_labels::{Frequency, Interval, Layout};
use litchi_iwa_common::chart::kind::Kind;
use litchi_iwa_common::chart::number_format::{
    DecimalPlaces, LabelAffixes, NegativeStyle, NumberFormat,
};
use litchi_iwa_common::color::{RgbColorSpace, Rgba};
use litchi_iwa_common::shape::stroke::{Pattern, Stroke, Width};
use litchi_pages::Package;
use litchi_pages::chart::axis::{GridlineStroke, Property, PropertyKind};

use crate::chart_style_cases::input;
use crate::harness::{Case, Evidence, Format, LeafRun, TestResult, Tier};
use crate::leaf_bytes;

const fn host_gridline(value: GridlineStroke) -> ChartAxisGridlineStroke {
    match value {
        GridlineStroke::Stroke(stroke) => ChartAxisGridlineStroke::Stroke(stroke),
        GridlineStroke::None => ChartAxisGridlineStroke::NoStroke,
        _ => ChartAxisGridlineStroke::Inherited,
    }
}

fn host_write(
    source: &[u8],
    _sheet: (),
    chart: u64,
    axis: Axis,
    value: &Property,
) -> TestResult<Vec<u8>> {
    let mut e = pages::Draft::from_bytes(source)?;
    match value.clone() {
        Property::MajorGridlineStroke(v) => e.set_body_chart_axis_gridline_stroke(
            chart,
            axis,
            ChartAxisGridline::Major,
            host_gridline(v),
        )?,
        Property::MinorGridlineStroke(v) => e.set_body_chart_axis_gridline_stroke(
            chart,
            axis,
            ChartAxisGridline::Minor,
            host_gridline(v),
        )?,
        Property::LabelAngle(v) => e.set_body_chart_axis_label_angle(chart, axis, v)?,
        Property::LabelAffixes(v) => e.set_body_chart_axis_label_affixes(chart, axis, v)?,
        Property::NumberFormat(v) => e.set_body_chart_axis_number_format(chart, axis, v)?,
        Property::LabelPosition3d(v) => {
            e.set_body_chart_3d_value_axis_label_position(chart, v)?;
        },
        Property::CategoryLabels(v) => e.set_body_chart_category_label_layout(chart, v)?,
        Property::Line(v) => e.set_body_chart_axis_line_visible(chart, axis, v)?,
        Property::MajorGridlines(v) => {
            e.set_body_chart_axis_major_gridlines_visible(chart, axis, v)?;
        },
        Property::MinorGridlines(v) => {
            e.set_body_chart_axis_minor_gridlines_visible(chart, axis, v)?;
        },
        Property::MinorTickMarks(v) => {
            e.set_body_chart_axis_minor_tick_marks_visible(chart, axis, v)?;
        },
        Property::TickMarkLocation(v) => {
            e.set_body_chart_axis_tick_mark_location(chart, axis, v)?;
        },
        Property::MinimumLabel(v) => {
            e.set_body_chart_value_axis_minimum_label_visible(chart, v)?;
        },
        Property::Labels(v) => {
            e.set_body_chart_axis_labels_visible(chart, axis, v.is_visible())?;
        },
        Property::SeriesNames(v) => {
            e.set_body_chart_category_axis_series_names_visible(chart, v.is_visible())?;
        },
        _ => return Err("unmapped property".into()),
    }
    Ok(e.to_bytes()?)
}

fn host_read(
    bytes: &[u8],
    _sheet: (),
    chart: u64,
    axis: Axis,
    kind: PropertyKind,
) -> TestResult<String> {
    let e = pages::Draft::from_bytes(bytes)?;
    let gridline = |stroke: ChartAxisGridlineStroke| match stroke {
        ChartAxisGridlineStroke::Stroke(stroke) => format!("{:?}", GridlineStroke::Stroke(stroke)),
        ChartAxisGridlineStroke::NoStroke => format!("{:?}", GridlineStroke::None),
        _ => format!("{:?}", GridlineStroke::Inherited),
    };
    Ok(match kind {
        PropertyKind::MajorGridlineStroke => {
            gridline(e.body_chart_axis_gridline_stroke(chart, axis, ChartAxisGridline::Major)?)
        },
        PropertyKind::MinorGridlineStroke => {
            gridline(e.body_chart_axis_gridline_stroke(chart, axis, ChartAxisGridline::Minor)?)
        },
        PropertyKind::LabelAngle => {
            format!("{:?}", e.body_chart_axis_label_angle(chart, axis)?)
        },
        PropertyKind::LabelAffixes => {
            format!("{:?}", e.body_chart_axis_label_affixes(chart, axis)?)
        },
        PropertyKind::NumberFormat => {
            format!("{:?}", e.body_chart_axis_number_format(chart, axis)?)
        },
        PropertyKind::LabelPosition3d => {
            format!("{:?}", e.body_chart_3d_value_axis_label_position(chart)?)
        },
        PropertyKind::CategoryLabels => {
            format!("{:?}", e.body_chart_category_label_layout(chart)?)
        },
        PropertyKind::Line => format!("{:?}", e.body_chart_axis_line_visible(chart, axis)?),
        PropertyKind::MajorGridlines => format!(
            "{:?}",
            e.body_chart_axis_major_gridlines_visible(chart, axis)?
        ),
        PropertyKind::MinorGridlines => format!(
            "{:?}",
            e.body_chart_axis_minor_gridlines_visible(chart, axis)?
        ),
        PropertyKind::MinorTickMarks => format!(
            "{:?}",
            e.body_chart_axis_minor_tick_marks_visible(chart, axis)?
        ),
        PropertyKind::TickMarkLocation => {
            format!("{:?}", e.body_chart_axis_tick_mark_location(chart, axis)?)
        },
        PropertyKind::MinimumLabel => format!(
            "{:?}",
            e.body_chart_value_axis_minimum_label_visible(chart)?
        ),
        PropertyKind::Labels => format!(
            "{:?}",
            Visibility::from(e.body_chart_axis_labels_visible(chart, axis)?)
        ),
        PropertyKind::SeriesNames => format!(
            "{:?}",
            Visibility::from(e.body_chart_category_axis_series_names_visible(chart)?)
        ),
        _ => return Err("unmapped kind".into()),
    })
}

fn leaf_read(bytes: &[u8], axis: Axis, kind: PropertyKind) -> TestResult<String> {
    let value = Package::from_bytes(bytes)?.body_chart_axis(last_chart(bytes)?, axis, kind)?;
    Ok(match value {
        Property::Line(v)
        | Property::MajorGridlines(v)
        | Property::MinorGridlines(v)
        | Property::MinorTickMarks(v)
        | Property::MinimumLabel(v)
        | Property::Labels(v)
        | Property::SeriesNames(v) => format!("{v:?}"),
        Property::TickMarkLocation(v) => format!("{v:?}"),
        Property::MajorGridlineStroke(v) | Property::MinorGridlineStroke(v) => format!("{v:?}"),
        Property::LabelAngle(v) => format!("{v:?}"),
        Property::LabelAffixes(v) => format!("{v:?}"),
        Property::NumberFormat(v) => format!("{v:?}"),
        Property::LabelPosition3d(v) => format!("{v:?}"),
        Property::CategoryLabels(v) => format!("{v:?}"),
        _ => return Err("unmapped property".into()),
    })
}

/// The host-added chart is the body's only discoverable chart.
fn last_chart(_bytes: &[u8]) -> TestResult<usize> {
    Ok(0)
}

fn cases() -> Vec<(Axis, Property)> {
    vec![
        (Axis::Value, Property::Line(Visibility::Visible)),
        (Axis::Category, Property::Line(Visibility::Visible)),
        (Axis::Value, Property::MajorGridlines(Visibility::Visible)),
        (
            Axis::Category,
            Property::MajorGridlines(Visibility::Visible),
        ),
        (Axis::Value, Property::MinorGridlines(Visibility::Visible)),
        (
            Axis::Category,
            Property::MinorGridlines(Visibility::Visible),
        ),
        (Axis::Value, Property::MinorTickMarks(Visibility::Hidden)),
        (Axis::Category, Property::MinorTickMarks(Visibility::Hidden)),
        (
            Axis::Value,
            Property::TickMarkLocation(TickMarkLocation::Outside),
        ),
        (
            Axis::Category,
            Property::TickMarkLocation(TickMarkLocation::Inside),
        ),
        (Axis::Value, Property::MinimumLabel(Visibility::Hidden)),
        (Axis::Value, Property::Labels(Visibility::Hidden)),
        (Axis::Category, Property::Labels(Visibility::Hidden)),
        (Axis::Category, Property::SeriesNames(Visibility::Visible)),
        (
            Axis::Value,
            Property::MajorGridlineStroke(GridlineStroke::Stroke(stroke())),
        ),
        (
            Axis::Category,
            Property::MinorGridlineStroke(GridlineStroke::None),
        ),
        (
            Axis::Category,
            Property::LabelAngle(LabelAngle::new(315.0).expect("angle")),
        ),
        (
            Axis::Value,
            Property::LabelAffixes(LabelAffixes::new("$", " M").expect("affixes")),
        ),
        (
            Axis::Value,
            Property::NumberFormat(NumberFormat::new(
                DecimalPlaces::fixed(2).expect("places"),
                NegativeStyle::Parentheses,
                true,
            )),
        ),
        (
            Axis::Value,
            Property::LabelPosition3d(LabelPosition3d::Leading),
        ),
        (
            Axis::Category,
            Property::CategoryLabels(Layout::new(
                Frequency::Every(Interval::new(3).expect("interval")),
                false,
            )),
        ),
        (
            Axis::Category,
            Property::CategoryLabels(Layout::new(Frequency::None, true)),
        ),
        (
            Axis::Category,
            Property::CategoryLabels(Layout::new(Frequency::All, true)),
        ),
        (
            Axis::Category,
            Property::CategoryLabels(Layout::new(Frequency::Unsupported(-4), false)),
        ),
        (
            Axis::Category,
            Property::CategoryLabels(Layout::new(Frequency::None, false)),
        ),
    ]
}

fn stroke() -> Stroke {
    Stroke::new(
        Rgba::new(0.2, 0.4, 0.6, 1.0, RgbColorSpace::Srgb).expect("color"),
        Width::new(2.0).expect("width"),
        Pattern::Solid,
    )
}

/// The chart kind a case needs: the 3D label position needs a 3D chart.
const fn chart_kind(value: &Property) -> Kind {
    if matches!(value, Property::LabelPosition3d(_)) {
        Kind::Column3d
    } else {
        Kind::Column2d
    }
}

/// `value`, or its opposite when the host-added chart already stores it.
fn changed(bytes: &[u8], axis: Axis, value: Property) -> TestResult<Property> {
    let current =
        Package::from_bytes(bytes)?.body_chart_axis(last_chart(bytes)?, axis, value.kind())?;
    if current != value {
        return Ok(value);
    }
    let flip = |v: Visibility| {
        if v.is_visible() {
            Visibility::Hidden
        } else {
            Visibility::Visible
        }
    };
    Ok(match value {
        Property::Line(v) => Property::Line(flip(v)),
        Property::MajorGridlines(v) => Property::MajorGridlines(flip(v)),
        Property::MinorGridlines(v) => Property::MinorGridlines(flip(v)),
        Property::MinorTickMarks(v) => Property::MinorTickMarks(flip(v)),
        Property::MinimumLabel(v) => Property::MinimumLabel(flip(v)),
        Property::Labels(v) => Property::Labels(flip(v)),
        Property::SeriesNames(v) => Property::SeriesNames(flip(v)),
        Property::TickMarkLocation(_) => Property::TickMarkLocation(TickMarkLocation::Centered),
        Property::MajorGridlineStroke(_) => Property::MajorGridlineStroke(GridlineStroke::None),
        Property::MinorGridlineStroke(_) => {
            Property::MinorGridlineStroke(GridlineStroke::Stroke(stroke()))
        },
        Property::LabelAngle(_) => Property::LabelAngle(LabelAngle::new(30.0)?),
        Property::LabelAffixes(_) => Property::LabelAffixes(LabelAffixes::new("~", "")?),
        Property::NumberFormat(_) => Property::NumberFormat(NumberFormat::AXIS_NATIVE_DEFAULT),
        Property::LabelPosition3d(_) => Property::LabelPosition3d(LabelPosition3d::Trailing),
        Property::CategoryLabels(_) => Property::CategoryLabels(Layout::new(Frequency::All, false)),
        _ => return Err("unmapped property".into()),
    })
}

#[test]
fn pages_body_chart_axis_set_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    for (axis, value) in cases() {
        let input = input(Format::Pages, chart_kind(&value))?;
        let sheet = ();
        let chart = input.chart;
        let kind = value.kind();
        let value = changed(&input.fixture.bytes, axis, value)?;
        let staged = value.clone();
        Case::new(
            "pages_body_chart_axis_set_parity",
            Format::Pages,
            Tier::A,
            Evidence::Queued,
            &input.fixture,
        )
        .host(|source| host_write(source, sheet, chart, axis, &value))
        .leaf(|source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_body_chart_axis(last_chart(source)?)?;
            edit = edit.set(axis, staged.clone())?;
            let commit = edit.commit()?;
            let restored = commit
                .package()
                .apply_body_chart_axis(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        })
        .read("host axis property", move |bytes| {
            host_read(bytes, sheet, chart, axis, kind)
        })
        .read("leaf axis property", move |bytes| {
            leaf_read(bytes, axis, kind)
        })
        .run()
        .map_err(|error| format!("{axis:?} {value:?}: {error}"))?;
    }
    Ok(())
}

#[test]
fn pages_body_chart_axis_read_parity() -> TestResult {
    for (axis, value) in cases() {
        let input = input(Format::Pages, chart_kind(&value))?;
        let sheet = ();
        for bytes in [
            input.fixture.bytes.clone(),
            host_write(&input.fixture.bytes, sheet, input.chart, axis, &value)?,
        ] {
            assert_eq!(
                host_read(&bytes, sheet, input.chart, axis, value.kind())?,
                leaf_read(&bytes, axis, value.kind())?,
                "{axis:?} {value:?}"
            );
        }
    }
    let _ = (Tier::A, Evidence::Queued);
    Ok(())
}
