//! Numbers sheet chart axis-style parity (wave W2-A, record 0378): the host
//! `set_sheet_chart_axis_{line,major_gridlines,minor_gridlines,
//! minor_tick_marks}_visible`, `set_sheet_chart_axis_tick_mark_location`
//! and `set_sheet_chart_value_axis_minimum_label_visible` routes against
//! `litchi_numbers::sheet::chart::axis::Edit::set`.

use litchi_iwa::numbers;
use litchi_iwa_common::chart::axis::style::Visibility;
use litchi_iwa_common::chart::axis::{Axis, TickMarkLocation};
use litchi_iwa_common::chart::kind::Kind;
use litchi_numbers::Package;
use litchi_numbers::sheet::chart::axis::{Property, PropertyKind};
use litchi_numbers::sheet::drawable::Kind as DrawableKind;

use crate::chart_style_cases::{Target, input};
use crate::harness::{Case, Evidence, Format, LeafRun, TestResult, Tier};
use crate::leaf_bytes;

fn host_write(
    source: &[u8],
    sheet: u64,
    chart: u64,
    axis: Axis,
    value: Property,
) -> TestResult<Vec<u8>> {
    let mut e = numbers::Draft::from_bytes(source)?;
    match value {
        Property::Line(v) => e.set_sheet_chart_axis_line_visible(sheet, chart, axis, v)?,
        Property::MajorGridlines(v) => {
            e.set_sheet_chart_axis_major_gridlines_visible(sheet, chart, axis, v)?;
        },
        Property::MinorGridlines(v) => {
            e.set_sheet_chart_axis_minor_gridlines_visible(sheet, chart, axis, v)?;
        },
        Property::MinorTickMarks(v) => {
            e.set_sheet_chart_axis_minor_tick_marks_visible(sheet, chart, axis, v)?;
        },
        Property::TickMarkLocation(v) => {
            e.set_sheet_chart_axis_tick_mark_location(sheet, chart, axis, v)?;
        },
        Property::MinimumLabel(v) => {
            e.set_sheet_chart_value_axis_minimum_label_visible(sheet, chart, v)?;
        },
        _ => return Err("unmapped property".into()),
    }
    Ok(e.to_bytes()?)
}

fn host_read(
    bytes: &[u8],
    sheet: u64,
    chart: u64,
    axis: Axis,
    kind: PropertyKind,
) -> TestResult<String> {
    let e = numbers::Draft::from_bytes(bytes)?;
    Ok(match kind {
        PropertyKind::Line => format!("{:?}", e.sheet_chart_axis_line_visible(sheet, chart, axis)?),
        PropertyKind::MajorGridlines => format!(
            "{:?}",
            e.sheet_chart_axis_major_gridlines_visible(sheet, chart, axis)?
        ),
        PropertyKind::MinorGridlines => format!(
            "{:?}",
            e.sheet_chart_axis_minor_gridlines_visible(sheet, chart, axis)?
        ),
        PropertyKind::MinorTickMarks => format!(
            "{:?}",
            e.sheet_chart_axis_minor_tick_marks_visible(sheet, chart, axis)?
        ),
        PropertyKind::TickMarkLocation => format!(
            "{:?}",
            e.sheet_chart_axis_tick_mark_location(sheet, chart, axis)?
        ),
        PropertyKind::MinimumLabel => format!(
            "{:?}",
            e.sheet_chart_value_axis_minimum_label_visible(sheet, chart)?
        ),
        _ => return Err("unmapped kind".into()),
    })
}

fn leaf_read(bytes: &[u8], axis: Axis, kind: PropertyKind) -> TestResult<String> {
    let value = Package::from_bytes(bytes)?.sheet_chart_axis(0, last_chart(bytes)?, axis, kind)?;
    Ok(match value {
        Property::Line(v)
        | Property::MajorGridlines(v)
        | Property::MinorGridlines(v)
        | Property::MinorTickMarks(v)
        | Property::MinimumLabel(v) => format!("{v:?}"),
        Property::TickMarkLocation(v) => format!("{v:?}"),
        _ => return Err("unmapped property".into()),
    })
}

fn last_chart(bytes: &[u8]) -> TestResult<usize> {
    let count = Package::from_bytes(bytes)?
        .sheet_drawable_list(0)?
        .iter()
        .filter(|item| item.kind() == DrawableKind::Chart)
        .count();
    count.checked_sub(1).ok_or_else(|| "no chart".into())
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
    ]
}

/// `value`, or its opposite when the host-added chart already stores it.
fn changed(bytes: &[u8], axis: Axis, value: Property) -> TestResult<Property> {
    let current =
        Package::from_bytes(bytes)?.sheet_chart_axis(0, last_chart(bytes)?, axis, value.kind())?;
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
        Property::TickMarkLocation(_) => Property::TickMarkLocation(TickMarkLocation::Centered),
        _ => return Err("unmapped property".into()),
    })
}

#[test]
fn numbers_sheet_chart_axis_set_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    for (axis, value) in cases() {
        let input = input(Format::Numbers, Kind::Column2d)?;
        let Target::Numbers { sheet } = input.target else {
            return Err("not a Numbers target".into());
        };
        let chart = input.chart;
        let kind = value.kind();
        let value = changed(&input.fixture.bytes, axis, value)?;
        Case::new(
            "numbers_sheet_chart_axis_set_parity",
            Format::Numbers,
            Tier::A,
            Evidence::Queued,
            &input.fixture,
        )
        .host(|source| host_write(source, sheet, chart, axis, value))
        .leaf(|source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_sheet_chart_axis(0, last_chart(source)?)?;
            edit.set(axis, value)?;
            let commit = edit.commit()?;
            let restored = commit
                .package()
                .apply_sheet_chart_axis(&commit.patch().inverse())?;
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
fn numbers_sheet_chart_axis_read_parity() -> TestResult {
    for (axis, value) in cases() {
        let input = input(Format::Numbers, Kind::Column2d)?;
        let Target::Numbers { sheet } = input.target else {
            return Err("not a Numbers target".into());
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
    let _ = (Tier::A, Evidence::Queued);
    Ok(())
}
