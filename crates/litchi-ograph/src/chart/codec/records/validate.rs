//! Typed chart-model validation for MS-OGRAPH record emission.

use super::super::super::axis;
use super::super::super::format::Format;
use super::super::super::model::{
    Cache, Chart, Family, Owner, Role, Value, XlValue, cache_dimensions, dimensions_cover,
};
use super::links;
use super::wire::{DATA_LAB_EXT, DATA_LAB_EXT_CONTENTS, TEXT, invalid_model, limit};
use crate::{Error, Limits, Result};

pub(super) fn validate(chart: &Chart, limits: Limits, require_topology: bool) -> Result<()> {
    if chart.series.len() > limits.max_series {
        return limit("series count", chart.series.len(), limits.max_series);
    }
    if chart.groups.len() > limits.max_groups {
        return limit("group count", chart.groups.len(), limits.max_groups);
    }
    if chart.axes.len() > limits.max_axes {
        return limit("axis count", chart.axes.len(), limits.max_axes);
    }
    if chart.caches.len() > limits.max_cached_values {
        return limit(
            "cached value count",
            chart.caches.len(),
            limits.max_cached_values,
        );
    }
    if !valid_props(chart.props.flags) {
        return invalid_model(
            "properties",
            "ShtProps reserved bits, blank mode, or plot-area flags are invalid",
        );
    }
    if (require_topology && !(1..=2).contains(&chart.parents.len()))
        || (!require_topology && chart.parents.len() > 2)
    {
        return invalid_model(
            "axis parents",
            "the chart has an invalid number of AxisParent collections",
        );
    }
    for (index, parent) in chart.parents.iter().copied().enumerate() {
        if parent.id().index() != index || !parent.pos().is_plot() {
            return invalid_model(
                "axis parents",
                "axis parents must be primary then optional secondary with plot positions",
            );
        }
    }
    if !chart.dimensions.matches(chart.context.kind()) {
        return invalid_model("Dimensions", "dimensions do not match the chart producer");
    }
    let derived = cache_dimensions(&chart.caches, chart.context.kind())?;
    if !dimensions_cover(chart.dimensions, derived) {
        return invalid_model(
            "Dimensions",
            "declared dimensions do not cover the cached chart cells",
        );
    }

    let mut orders = [false; 10];
    for group in &chart.groups {
        if chart
            .parents
            .get(group.parent.index())
            .is_none_or(|parent| parent.id() != group.parent)
        {
            return invalid_model("group", "chart group refers to a missing axis parent");
        }
        let order = usize::from(group.order.get());
        let Some(seen) = orders.get_mut(order) else {
            return invalid_model("group", "chart-group order exceeds nine");
        };
        if *seen {
            return invalid_model("group", "chart-group order is duplicated");
        }
        *seen = true;
        match group.family {
            Family::Line { flags } | Family::Area { flags } if flags & !7 != 0 => {
                return invalid_model("group", "line or area reserved flags are nonzero");
            },
            Family::Bar { flags, .. } if flags & !0xF != 0 => {
                return invalid_model("group", "bar reserved flags are nonzero");
            },
            Family::Pie {
                rotation,
                hole,
                flags,
            } if rotation > 360 || hole > 90 || flags & !3 != 0 => {
                return invalid_model("group", "pie settings are outside their BIFF ranges");
            },
            Family::Radar { flags, .. } | Family::Surface { flags } if flags & !3 != 0 => {
                return invalid_model("group", "radar or surface reserved flags are nonzero");
            },
            _ => {},
        }
        let mut prior = None;
        for line in &group.lines {
            let current = line.kind as u16;
            if prior.is_some_and(|value| current <= value) {
                return invalid_model(
                    "group",
                    "chart-group line kinds are duplicated or out of order",
                );
            }
            prior = Some(current);
        }
        if group.drop_bars.len() > 2 {
            return invalid_model(
                "group",
                "a chart group has more than two DropBar collections",
            );
        }
        if !group.drop_bars.is_empty() && !matches!(group.family, Family::Line { .. }) {
            return invalid_model("group", "DropBar is only valid on a line chart group");
        }
    }

    for series in &chart.series {
        match &series.owner {
            Owner::Group(group) if usize::from(group.get()) >= chart.groups.len() => {
                return invalid_model("series", "series refers to a missing chart group");
            },
            Owner::Trend { parent, .. } | Owner::ErrorBar { parent, .. } => {
                let zero_based =
                    parent
                        .series()
                        .get()
                        .checked_sub(1)
                        .ok_or(Error::InvalidModel {
                            field: "series",
                            reason: "auxiliary parent is not a one-based series index",
                        })?;
                let zero_based = usize::try_from(zero_based).map_err(|_| Error::InvalidModel {
                    field: "series",
                    reason: "auxiliary parent index exceeds this platform",
                })?;
                if chart
                    .series
                    .get(zero_based)
                    .is_none_or(|parent| !matches!(parent.owner, Owner::Group(_)))
                {
                    return invalid_model(
                        "series",
                        "auxiliary series must refer to an existing regular series",
                    );
                }
            },
            _ => {},
        }
        for (binding, role) in series.ai.ordered().into_iter().zip(Role::ALL) {
            if binding.link().role() != role {
                return invalid_model("AI", "series AI roles are not in canonical order");
            }
            links::validate_link(binding.link(), chart.context, limits)?;
            check_string(binding.text(), "AI text")?;
        }
    }
    check_string(chart.title.as_deref(), "title")?;

    for axis in &chart.axes {
        if chart
            .parents
            .get(axis.parent.index())
            .is_none_or(|parent| parent.id() != axis.parent)
        {
            return invalid_model("axis", "axis refers to a missing axis parent");
        }
        if let Some(scale) = axis.scale {
            let values = [
                scale.min,
                scale.max,
                scale.major,
                scale.minor,
                scale.crossing,
            ];
            if !values.into_iter().all(f64::is_finite)
                || scale.max < scale.min
                || scale.major < 0.0
                || scale.minor < 0.0
            {
                return invalid_model("axis", "scale is not finite, ordered, and nonnegative");
            }
        }
        let mut prior = None;
        for line in &axis.lines {
            let current = line_kind(line.kind);
            if prior.is_some_and(|value| current <= value) {
                return invalid_model("axis", "line roles are duplicated or out of order");
            }
            prior = Some(current);
        }
    }

    for value in &chart.caches {
        if value.kind() != chart.context.kind() {
            return invalid_model("cache", "cached cell does not match producer context");
        }
        match value {
            Cache::Excel {
                value: XlValue::Number(number),
                ..
            }
            | Cache::Graph {
                value: Value::Number(number),
                ..
            } if !number.is_finite() => {
                return invalid_model("cache", "cached number is not finite");
            },
            Cache::Excel {
                value: XlValue::Text(text),
                ..
            } => check_xl_string(text, limits)?,
            Cache::Graph {
                value: Value::Text(text),
                ..
            } => check_string(Some(text), "cached text")?,
            _ => {},
        }
    }
    for format in &chart.formats {
        if let Format::Marker { data } = format
            && data.len() > limits.biff.max_record_bytes
        {
            return limit("record bytes", data.len(), limits.biff.max_record_bytes);
        }
    }
    for label in &chart.labels {
        if !matches!(label.kind, DATA_LAB_EXT | DATA_LAB_EXT_CONTENTS | TEXT) {
            return invalid_model("label", "record kind is not a supported data-label kind");
        }
        if label.data.len() > limits.biff.max_record_bytes {
            return limit(
                "record bytes",
                label.data.len(),
                limits.biff.max_record_bytes,
            );
        }
    }

    let mut total = 0usize;
    for raw in &chart.unknown {
        if raw.data().len() > limits.biff.max_record_bytes {
            return limit(
                "record bytes",
                raw.data().len(),
                limits.biff.max_record_bytes,
            );
        }
        total = total
            .checked_add(raw.data().len())
            .ok_or(Error::SizeOverflow {
                resource: "unknown chart bytes",
            })?;
    }
    if total > limits.max_unknown_bytes {
        return limit("unknown chart bytes", total, limits.max_unknown_bytes);
    }
    Ok(())
}

pub(super) fn check_string(value: Option<&str>, field: &'static str) -> Result<()> {
    let Some(value) = value else {
        return Ok(());
    };
    if value.encode_utf16().count() > usize::from(u8::MAX) {
        return invalid_model(field, "chart string exceeds 255 UTF-16 code units");
    }
    Ok(())
}

pub(super) fn check_xl_string(value: &str, limits: Limits) -> Result<()> {
    let count = value.encode_utf16().count();
    if count > usize::from(u16::MAX) {
        return invalid_model(
            "cached text",
            "Excel chart string exceeds 65,535 UTF-16 code units",
        );
    }
    let width = if value.encode_utf16().any(|unit| unit > u16::from(u8::MAX)) {
        2usize
    } else {
        1usize
    };
    let payload = 9usize
        .checked_add(count.checked_mul(width).ok_or(Error::SizeOverflow {
            resource: "cached chart string",
        })?)
        .ok_or(Error::SizeOverflow {
            resource: "cached chart string",
        })?;
    if payload > limits.biff.max_record_bytes {
        return limit("record bytes", payload, limits.biff.max_record_bytes);
    }
    Ok(())
}

pub(crate) fn valid_props(flags: u32) -> bool {
    let reserved_clear = flags & 0x0000_FFE4 == 0 && flags & 0xFF00_0000 == 0;
    let blank = (flags >> 16) & 0xFF;
    let auto_plot = flags & (1 << 4) != 0;
    let manual_plot = flags & (1 << 3) != 0;
    reserved_clear && blank <= 2 && (!auto_plot || manual_plot)
}

pub(super) fn line_kind(kind: axis::LineKind) -> u8 {
    match kind {
        axis::LineKind::Axis => 0,
        axis::LineKind::MajorGrid => 1,
        axis::LineKind::MinorGrid => 2,
        axis::LineKind::Wall => 3,
    }
}
