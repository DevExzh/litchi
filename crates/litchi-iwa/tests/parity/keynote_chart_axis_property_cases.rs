//! Keynote slide chart axis-style parity (wave W4-C, record 0414): the host
//! `set_slide_chart_axis_{line,major_gridlines,minor_gridlines,
//! minor_tick_marks}_visible`, `set_slide_chart_axis_tick_mark_location`
//! and `set_slide_chart_value_axis_minimum_label_visible` routes against
//! `litchi_keynote::chart::axis::Edit::set`.

use litchi_iwa::keynote;
use litchi_iwa_common::chart::axis::style::Visibility;
use litchi_iwa_common::chart::axis::{Axis, TickMarkLocation};
use litchi_iwa_common::chart::kind::Kind;
use litchi_keynote::Package;
use litchi_keynote::chart::axis::{Property, PropertyKind};

use crate::chart_style_cases::{Target, input};
use crate::harness::{Case, Evidence, Format, LeafRun, TestResult, Tier};
use crate::leaf_bytes;

fn host_write(
    source: &[u8],
    sheet: usize,
    chart: u64,
    axis: Axis,
    value: Property,
) -> TestResult<Vec<u8>> {
    let mut e = keynote::Draft::from_bytes(source)?;
    match value {
        Property::Line(v) => e.set_slide_chart_axis_line_visible(sheet, chart, axis, v)?,
        Property::MajorGridlines(v) => {
            e.set_slide_chart_axis_major_gridlines_visible(sheet, chart, axis, v)?;
        },
        Property::MinorGridlines(v) => {
            e.set_slide_chart_axis_minor_gridlines_visible(sheet, chart, axis, v)?;
        },
        Property::MinorTickMarks(v) => {
            e.set_slide_chart_axis_minor_tick_marks_visible(sheet, chart, axis, v)?;
        },
        Property::TickMarkLocation(v) => {
            e.set_slide_chart_axis_tick_mark_location(sheet, chart, axis, v)?;
        },
        Property::MinimumLabel(v) => {
            e.set_slide_chart_value_axis_minimum_label_visible(sheet, chart, v)?;
        },
        Property::Labels(v) => {
            e.set_slide_chart_axis_labels_visible(sheet, chart, axis, v.is_visible())?;
        },
        Property::SeriesNames(v) => {
            e.set_slide_chart_category_axis_series_names_visible(sheet, chart, v.is_visible())?;
        },
        _ => return Err("unmapped property".into()),
    }
    Ok(e.to_bytes()?)
}

fn host_read(
    bytes: &[u8],
    sheet: usize,
    chart: u64,
    axis: Axis,
    kind: PropertyKind,
) -> TestResult<String> {
    let e = keynote::Draft::from_bytes(bytes)?;
    Ok(match kind {
        PropertyKind::Line => format!("{:?}", e.slide_chart_axis_line_visible(sheet, chart, axis)?),
        PropertyKind::MajorGridlines => format!(
            "{:?}",
            e.slide_chart_axis_major_gridlines_visible(sheet, chart, axis)?
        ),
        PropertyKind::MinorGridlines => format!(
            "{:?}",
            e.slide_chart_axis_minor_gridlines_visible(sheet, chart, axis)?
        ),
        PropertyKind::MinorTickMarks => format!(
            "{:?}",
            e.slide_chart_axis_minor_tick_marks_visible(sheet, chart, axis)?
        ),
        PropertyKind::TickMarkLocation => format!(
            "{:?}",
            e.slide_chart_axis_tick_mark_location(sheet, chart, axis)?
        ),
        PropertyKind::MinimumLabel => format!(
            "{:?}",
            e.slide_chart_value_axis_minimum_label_visible(sheet, chart)?
        ),
        PropertyKind::Labels => format!(
            "{:?}",
            Visibility::from(e.slide_chart_axis_labels_visible(sheet, chart, axis)?)
        ),
        PropertyKind::SeriesNames => format!(
            "{:?}",
            Visibility::from(e.slide_chart_category_axis_series_names_visible(sheet, chart)?)
        ),
        _ => return Err("unmapped kind".into()),
    })
}

fn leaf_read(bytes: &[u8], axis: Axis, kind: PropertyKind) -> TestResult<String> {
    let value = Package::from_bytes(bytes)?.slide_chart_axis(0, last_chart(bytes)?, axis, kind)?;
    Ok(match value {
        Property::Line(v)
        | Property::MajorGridlines(v)
        | Property::MinorGridlines(v)
        | Property::MinorTickMarks(v)
        | Property::MinimumLabel(v)
        | Property::Labels(v)
        | Property::SeriesNames(v) => format!("{v:?}"),
        Property::TickMarkLocation(v) => format!("{v:?}"),
        _ => return Err("unmapped property".into()),
    })
}

/// The host-added chart follows the fixture's native chart.
fn last_chart(_bytes: &[u8]) -> TestResult<usize> {
    Ok(1)
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
    ]
}

/// `value`, or its opposite when the host-added chart already stores it.
fn changed(bytes: &[u8], axis: Axis, value: Property) -> TestResult<Property> {
    let current =
        Package::from_bytes(bytes)?.slide_chart_axis(0, last_chart(bytes)?, axis, value.kind())?;
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
        _ => return Err("unmapped property".into()),
    })
}

#[test]
fn keynote_slide_chart_axis_set_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    for (axis, value) in cases() {
        let input = input(Format::Keynote, Kind::Column2d)?;
        let Target::Keynote { slide: sheet } = input.target else {
            return Err("not a Keynote target".into());
        };
        let chart = input.chart;
        let kind = value.kind();
        let value = changed(&input.fixture.bytes, axis, value)?;
        Case::new(
            "keynote_slide_chart_axis_set_parity",
            Format::Keynote,
            Tier::B,
            Evidence::Queued,
            &input.fixture,
        )
        .leaf_invalidates_root_previews()
        .leaf_structural_metadata()
        .host(|source| host_write(source, sheet, chart, axis, value))
        .leaf(|source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_slide_chart_axis(0, last_chart(source)?)?;
            edit.set(axis, value)?;
            let commit = edit.commit()?;
            let restored = commit
                .package()
                .apply_slide_chart_axis(&commit.patch().inverse())?;
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
fn keynote_slide_chart_axis_read_parity() -> TestResult {
    for (axis, value) in cases() {
        let input = input(Format::Keynote, Kind::Column2d)?;
        let Target::Keynote { slide: sheet } = input.target else {
            return Err("not a Keynote target".into());
        };
        for bytes in [
            input.fixture.bytes.clone(),
            host_write(&input.fixture.bytes, sheet, input.chart, axis, value)?,
        ] {
            assert_eq!(
                host_read(&bytes, sheet, input.chart, axis, value.kind())?,
                leaf_read(&bytes, axis, value.kind())?,
                "{axis:?} {value:?}"
            );
        }
    }
    let _ = (Tier::B, Evidence::Queued);
    Ok(())
}
