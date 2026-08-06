//! BIFF chart-substream model codecs.

#[cfg(test)]
use litchi_biff::Encoder as GraphEncoder;
use litchi_ograph::chart::{format, group};
use litchi_ograph::record::{chart3d, frame, line, marker, pie, series};

use super::model::*;
use super::package::{is_chart_bof, ranges_with};
use super::wire::*;
use crate::{Error, Result};

#[cfg(test)]
use super::package::chart_bof;

#[derive(Clone, Copy)]
enum PendingLine {
    Axis { owner: usize, kind: AxisLineKind },
    Group { owner: usize, kind: line::Kind },
}

struct PendingDrop {
    owner: usize,
    depth: usize,
    gap: group::Gap,
    line: Option<format::Line>,
    area: Option<format::Area>,
}

pub(super) fn parse_chart(input: &[u8], limits: Limits) -> Result<Chart> {
    let records = ranges_with(input, limits.max_records_per_chart)?;
    if records
        .first()
        .is_none_or(|v| v.kind != BOF || !is_chart_bof(&input[v.body_start..v.body_end]))
        || records.last().is_none_or(|v| v.kind != EOF)
    {
        return invalid(BOF, "chart substream must be bounded by chart BOF and EOF");
    }
    let mut chart = Chart {
        groups: Vec::new(),
        ..Default::default()
    };
    let mut depth = 0usize;
    let mut current_series = None;
    let mut series_depth = None;
    let mut current_axis = None;
    let mut axis_depth = None;
    let mut group_depth = None;
    let mut cache_kind = None;
    let mut pending_line = None;
    let mut pending_drop: Option<PendingDrop> = None;
    let mut pending_begin = false;
    for value in records.iter().skip(1).take(records.len() - 2) {
        let data = &input[value.body_start..value.body_end];
        if pending_line.is_some() && value.kind != LINE_FORMAT {
            return invalid(
                value.kind,
                "line owner is not followed immediately by LineFormat",
            );
        }
        if let Some(drop) = &pending_drop
            && depth == drop.depth
        {
            if drop.line.is_none() && value.kind != LINE_FORMAT {
                return invalid(value.kind, "DropBar Begin is not followed by LineFormat");
            }
            if drop.line.is_some() && drop.area.is_none() && value.kind != AREA_FORMAT {
                return invalid(
                    value.kind,
                    "DropBar LineFormat is not followed by AreaFormat",
                );
            }
        }
        if pending_begin && value.kind != BEGIN {
            return invalid(value.kind, "DropBar is not followed immediately by Begin");
        }
        match value.kind {
            BEGIN => {
                marker::Begin::from_payload(data).map_err(|error| graph_error(BEGIN, error))?;
                pending_begin = false;
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| Error::InvalidData("chart nesting overflow".into()))?;
                if depth > 128 {
                    return invalid(BEGIN, "chart nesting exceeds limit");
                }
            },
            END => {
                marker::End::from_payload(data).map_err(|error| graph_error(END, error))?;
                if depth == 0 {
                    return invalid(END, "unbalanced End record");
                }
                if series_depth == Some(depth) {
                    current_series = None;
                    series_depth = None;
                }
                if pending_drop
                    .as_ref()
                    .is_some_and(|drop| drop.depth == depth)
                {
                    let drop = pending_drop.take().ok_or_else(|| {
                        invalid_error(DROP_BAR, "DropBar collection state disappeared")
                    })?;
                    let line = drop.line.ok_or_else(|| {
                        invalid_error(DROP_BAR, "DropBar collection has no LineFormat")
                    })?;
                    let area = drop.area.ok_or_else(|| {
                        invalid_error(DROP_BAR, "DropBar collection has no AreaFormat")
                    })?;
                    chart
                        .groups
                        .get_mut(drop.owner)
                        .ok_or_else(|| invalid_error(DROP_BAR, "DropBar owner is missing"))?
                        .drop_bars
                        .push(group::DropBar {
                            gap: drop.gap,
                            line,
                            area,
                        });
                }
                if axis_depth == Some(depth) {
                    current_axis = None;
                    axis_depth = None;
                }
                if group_depth == Some(depth) {
                    group_depth = None;
                }
                depth -= 1;
            },
            CHART => {
                exact(data, 16, CHART)?;
                chart.x = i32_at(data, 0)?;
                chart.y = i32_at(data, 4)?;
                chart.width = i32_at(data, 8)?;
                chart.height = i32_at(data, 12)?;
            },
            SHT_PROPS => {
                exact(data, 4, SHT_PROPS)?;
                chart.sheet_properties = u32_at(data, 0)?;
                validate_sheet_properties(chart.sheet_properties)?;
            },
            SERIES => {
                exact(data, 12, SERIES)?;
                if chart.series.len() >= limits.max_series {
                    return invalid(SERIES, "series count exceeds limit");
                }
                let kind = match u16_at(data, 0)? {
                    1 => DataKind::Numeric,
                    3 => DataKind::Text,
                    _ => return invalid(SERIES, "invalid category data type"),
                };
                if u16_at(data, 2)? != 1 || u16_at(data, 8)? != 1 {
                    return invalid(SERIES, "series numeric data type fields are invalid");
                }
                chart.series.push(Series {
                    category_kind: kind,
                    category_count: bounded_count(data, 4)?,
                    value_count: bounded_count(data, 6)?,
                    bubble_count: bounded_count(data, 10)?,
                    chart_group: 0,
                    name: None,
                    links: Vec::new(),
                });
                current_series = Some(chart.series.len() - 1);
                series_depth = Some(depth + 1);
            },
            BRAI => {
                if let Some(series) = current_series {
                    let link = parse_link(data, limits)?;
                    chart
                        .series
                        .get_mut(series)
                        .ok_or_else(|| invalid_error(BRAI, "Series index is invalid"))?
                        .links
                        .push(link);
                } else {
                    chart.unknown_records.push(Raw {
                        record_type: value.kind,
                        data: data.to_vec(),
                    });
                }
            },
            SER_TO_CRT => {
                exact(data, 2, SER_TO_CRT)?;
                let index = current_series.ok_or_else(|| {
                    invalid_error(SER_TO_CRT, "SerToCrt appears outside a Series")
                })?;
                chart
                    .series
                    .get_mut(index)
                    .ok_or_else(|| invalid_error(SER_TO_CRT, "Series index is invalid"))?
                    .chart_group = u16_at(data, 0)?;
            },
            SERIES_TEXT => {
                let mut text = Some(parse_short_text(data)?);
                if let Some(index) = current_series {
                    let series = chart
                        .series
                        .get_mut(index)
                        .ok_or_else(|| invalid_error(SERIES_TEXT, "Series index is invalid"))?;
                    if series.name.is_none() {
                        series.name = text.take();
                    }
                }
                if chart.title.is_none() {
                    chart.title = text;
                }
            },
            CHART_FORMAT => {
                if group_depth.is_some() {
                    return invalid(CHART_FORMAT, "ChartFormat collections overlap");
                }
                exact(data, 20, CHART_FORMAT)?;
                if data[..16].iter().any(|v| *v != 0) || u16_at(data, 16)? & !1 != 0 {
                    return invalid(CHART_FORMAT, "ChartFormat reserved fields are nonzero");
                }
                let order = u16_at(data, 18)?;
                chart.groups.push(Group {
                    order,
                    vary_colors: u16_at(data, 16)? & 1 != 0,
                    kind: GroupKind::Line { flags: 0 },
                    lines: Vec::new(),
                    drop_bars: Vec::new(),
                });
                group_depth = depth.checked_add(1);
            },
            BAR | LINE | PIE | AREA | SCATTER | RADAR | RADAR_AREA | SURFACE => {
                if group_depth != Some(depth) {
                    return invalid(value.kind, "chart family appears outside ChartFormat");
                }
                let kind = parse_group(value.kind, data)?;
                chart
                    .groups
                    .last_mut()
                    .ok_or_else(|| invalid_error(value.kind, "chart family has no group owner"))?
                    .kind = kind;
            },
            CRT_LINE => {
                if group_depth != Some(depth) {
                    return invalid(CRT_LINE, "CrtLine appears outside ChartFormat");
                }
                let value =
                    line::Line::from_payload(data).map_err(|error| graph_error(CRT_LINE, error))?;
                let owner =
                    chart.groups.len().checked_sub(1).ok_or_else(|| {
                        invalid_error(CRT_LINE, "CrtLine has no ChartFormat owner")
                    })?;
                pending_line = Some(PendingLine::Group {
                    owner,
                    kind: value.kind(),
                });
            },
            DROP_BAR => {
                if group_depth != Some(depth) {
                    return invalid(DROP_BAR, "DropBar appears outside ChartFormat");
                }
                if pending_drop.is_some() {
                    return invalid(DROP_BAR, "DropBar collections overlap");
                }
                exact(data, 2, DROP_BAR)?;
                let gap = group::Gap::new(u16_at(data, 0)?)
                    .ok_or_else(|| invalid_error(DROP_BAR, "DropBar gap exceeds 500"))?;
                let owner =
                    chart.groups.len().checked_sub(1).ok_or_else(|| {
                        invalid_error(DROP_BAR, "DropBar has no ChartFormat owner")
                    })?;
                if chart
                    .groups
                    .get(owner)
                    .is_none_or(|group| group.drop_bars.len() >= 2)
                {
                    return invalid(
                        DROP_BAR,
                        "chart group has more than two DropBar collections",
                    );
                }
                pending_drop = Some(PendingDrop {
                    owner,
                    depth: depth
                        .checked_add(1)
                        .ok_or_else(|| Error::InvalidData("DropBar nesting overflow".into()))?,
                    gap,
                    line: None,
                    area: None,
                });
                pending_begin = true;
            },
            AXIS => {
                if current_axis.is_some() {
                    return invalid(AXIS, "Axis collections overlap");
                }
                exact(data, 18, AXIS)?;
                if data[2..].iter().any(|v| *v != 0) {
                    return invalid(AXIS, "Axis reserved fields are nonzero");
                }
                let kind = match u16_at(data, 0)? {
                    0 => AxisKind::CategoryOrHorizontal,
                    1 => AxisKind::ValueOrVertical,
                    2 => AxisKind::Series,
                    _ => return invalid(AXIS, "invalid axis kind"),
                };
                chart.axes.push(Axis {
                    kind,
                    scale: None,
                    tick: None,
                    lines: Vec::new(),
                });
                current_axis = Some(chart.axes.len() - 1);
                axis_depth = depth.checked_add(1);
            },
            AXES_USED => {
                exact(data, 2, AXES_USED)?;
                if !matches!(u16_at(data, 0)?, 1 | 2) {
                    return invalid(AXES_USED, "AxesUsed must specify one or two axis groups");
                }
            },
            AXIS_PARENT => {
                exact(data, 18, AXIS_PARENT)?;
                if u16_at(data, 0)? > 1 {
                    return invalid(AXIS_PARENT, "AxisParent index must be primary or secondary");
                }
            },
            VALUE_RANGE => {
                exact(data, 42, VALUE_RANGE)?;
                let axis = current_axis
                    .ok_or_else(|| invalid_error(VALUE_RANGE, "ValueRange appears before Axis"))?;
                chart
                    .axes
                    .get_mut(axis)
                    .ok_or_else(|| invalid_error(VALUE_RANGE, "Axis index is invalid"))?
                    .scale = Some(Scale {
                    minimum: f64_at(data, 0)?,
                    maximum: f64_at(data, 8)?,
                    major: f64_at(data, 16)?,
                    minor: f64_at(data, 24)?,
                    crossing: f64_at(data, 32)?,
                    flags: u16_at(data, 40)?,
                });
            },
            TICK => {
                if data.len() < 26 {
                    return invalid(TICK, "Tick record is truncated");
                }
                let axis =
                    current_axis.ok_or_else(|| invalid_error(TICK, "Tick appears before Axis"))?;
                chart
                    .axes
                    .get_mut(axis)
                    .ok_or_else(|| invalid_error(TICK, "Axis index is invalid"))?
                    .tick = Some(Tick {
                    major: data[0],
                    minor: data[1],
                    label_position: data[2],
                    background: data[3],
                    color: array_at(data, 4)?,
                    flags: u16_at(data, 24)?,
                });
            },
            AXIS_LINE => {
                if axis_depth != Some(depth) {
                    return invalid(AXIS_LINE, "AxisLine appears outside Axis");
                }
                exact(data, 2, AXIS_LINE)?;
                let kind = match u16_at(data, 0)? {
                    0 => AxisLineKind::Axis,
                    1 => AxisLineKind::MajorGridlines,
                    2 => AxisLineKind::MinorGridlines,
                    3 => AxisLineKind::WallsOrFloor,
                    _ => return invalid(AXIS_LINE, "invalid AxisLine kind"),
                };
                let axis = current_axis
                    .ok_or_else(|| invalid_error(AXIS_LINE, "AxisLine appears before Axis"))?;
                if chart.axes.get(axis).is_none() {
                    return invalid(AXIS_LINE, "Axis index is invalid");
                }
                pending_line = Some(PendingLine::Axis { owner: axis, kind });
            },
            LINE_FORMAT => {
                let value = parse_line_format(data)?;
                let shared = shared_line(&value);
                match pending_line.take() {
                    Some(PendingLine::Axis { owner, kind }) => chart
                        .axes
                        .get_mut(owner)
                        .ok_or_else(|| invalid_error(LINE_FORMAT, "pending Axis owner is missing"))?
                        .lines
                        .push(AxisLine {
                            kind,
                            format: value,
                        }),
                    Some(PendingLine::Group { owner, kind }) => chart
                        .groups
                        .get_mut(owner)
                        .ok_or_else(|| {
                            invalid_error(LINE_FORMAT, "pending ChartFormat owner is missing")
                        })?
                        .lines
                        .push(group::Line {
                            kind,
                            format: shared,
                        }),
                    None => {
                        if let Some(drop) = pending_drop.as_mut().filter(|drop| drop.depth == depth)
                        {
                            if drop.line.replace(shared).is_some() {
                                return invalid(
                                    LINE_FORMAT,
                                    "DropBar has multiple LineFormat records",
                                );
                            }
                        } else {
                            chart.formatting.push(Format::Line(value));
                        }
                    },
                }
            },
            AREA_FORMAT => {
                let value = parse_area_format(data)?;
                if let Some(drop) = pending_drop.as_mut().filter(|drop| drop.depth == depth) {
                    if drop.area.replace(shared_area(&value)).is_some() {
                        return invalid(AREA_FORMAT, "DropBar has multiple AreaFormat records");
                    }
                } else {
                    chart.formatting.push(Format::Area(value));
                }
            },
            MARKER_FORMAT => chart.formatting.push(Format::Marker {
                data: data.to_vec(),
            }),
            DATA_FORMAT => {
                exact(data, 8, DATA_FORMAT)?;
                chart.formatting.push(Format::Data {
                    point: u16_at(data, 0)?,
                    series: u16_at(data, 2)?,
                    flags: u16_at(data, 6)?,
                });
            },
            PIE_FORMAT => {
                let format = pie::Format::from_payload(data)
                    .map_err(|error| graph_error(PIE_FORMAT, error))?;
                chart.formatting.push(Format::Pie(format));
            },
            LEGEND => {
                exact(data, 20, LEGEND)?;
                chart.legend = Some(Legend {
                    x: i32_at(data, 0)?,
                    y: i32_at(data, 4)?,
                    width: i32_at(data, 8)?,
                    height: i32_at(data, 12)?,
                    position: data[16],
                    spacing: data[17],
                    flags: u16_at(data, 18)?,
                });
            },
            PLOT_AREA => {
                marker::PlotArea::from_payload(data)
                    .map_err(|error| graph_error(PLOT_AREA, error))?;
                chart.plot_area_present = true;
            },
            DATA_LAB_EXT | DATA_LAB_EXT_CONTENTS | TEXT => chart.data_labels.push(Label {
                record_type: value.kind,
                data: data.to_vec(),
            }),
            FRAME | CRT_LINK | SERIES_FORMAT | SERIES_PARENT | BAR_SHAPE => {
                match value.kind {
                    FRAME => {
                        frame::Frame::from_payload(data)
                            .map_err(|error| graph_error(FRAME, error))?;
                    },
                    CRT_LINK => {
                        line::Link::from_payload(data)
                            .map_err(|error| graph_error(CRT_LINK, error))?;
                    },
                    SERIES_FORMAT => {
                        series::Format::from_payload(data)
                            .map_err(|error| graph_error(SERIES_FORMAT, error))?;
                    },
                    SERIES_PARENT => {
                        series::Parent::from_payload(data)
                            .map_err(|error| graph_error(SERIES_PARENT, error))?;
                    },
                    BAR_SHAPE => {
                        chart3d::BarShape::from_payload(data)
                            .map_err(|error| graph_error(BAR_SHAPE, error))?;
                    },
                    _ => {},
                }
                chart.unknown_records.push(Raw {
                    record_type: value.kind,
                    data: data.to_vec(),
                });
            },
            CONTINUE | SERIES_LIST | CAT_SER_RANGE | DEFAULT_TEXT | FONT_X | OBJECT_LINK
            | PLOT_GROWTH => chart.unknown_records.push(Raw {
                record_type: value.kind,
                data: data.to_vec(),
            }),
            SI_INDEX => {
                exact(data, 2, SI_INDEX)?;
                cache_kind = Some(CacheKind::from_wire(u16_at(data, 0)?).ok_or_else(|| {
                    invalid_error(SI_INDEX, "SIIndex selects an undefined cache section")
                })?);
            },
            BLANK => {
                exact(data, 6, BLANK)?;
                let kind = cache_kind
                    .ok_or_else(|| invalid_error(SI_INDEX, "cache data precedes SIIndex"))?;
                chart.cached_values.push(parse_cache(kind, BLANK, data)?);
            },
            NUMBER => {
                exact(data, 14, NUMBER)?;
                let kind = cache_kind
                    .ok_or_else(|| invalid_error(SI_INDEX, "cache data precedes SIIndex"))?;
                chart.cached_values.push(parse_cache(kind, NUMBER, data)?);
            },
            LABEL => {
                let kind = cache_kind
                    .ok_or_else(|| invalid_error(SI_INDEX, "cache data precedes SIIndex"))?;
                chart.cached_values.push(parse_cache(kind, LABEL, data)?);
            },
            BOF | EOF => return invalid(value.kind, "nested BOF/EOF is invalid in a chart"),
            _ => chart.unknown_records.push(Raw {
                record_type: value.kind,
                data: data.to_vec(),
            }),
        }
    }
    if depth != 0 {
        return invalid(BEGIN, "chart Begin/End collections are unbalanced");
    }
    if pending_line.is_some() || pending_drop.is_some() || pending_begin {
        return invalid(
            CHART,
            "chart collection ended with incomplete formatting ownership",
        );
    }
    chart.validate(limits)?;
    Ok(chart)
}

fn parse_cache(kind: CacheKind, record_type: u16, data: &[u8]) -> Result<Cache> {
    let (point, series, format, value) = match record_type {
        BLANK => {
            exact(data, 6, BLANK)?;
            (
                u16_at(data, 0)?,
                u8::try_from(u16_at(data, 2)?)
                    .map_err(|_| invalid_error(BLANK, "cached column exceeds BIFF8 grid"))?,
                u16_at(data, 4)?,
                Value::Blank,
            )
        },
        NUMBER => {
            exact(data, 14, NUMBER)?;
            (
                u16_at(data, 0)?,
                u8::try_from(u16_at(data, 2)?)
                    .map_err(|_| invalid_error(NUMBER, "cached column exceeds BIFF8 grid"))?,
                u16_at(data, 4)?,
                Value::Number(f64_at(data, 6)?),
            )
        },
        LABEL => {
            if data.len() < 9 {
                return invalid(LABEL, "cached Label is truncated");
            }
            (
                u16_at(data, 0)?,
                u8::try_from(u16_at(data, 2)?)
                    .map_err(|_| invalid_error(LABEL, "cached column exceeds BIFF8 grid"))?,
                u16_at(data, 4)?,
                Value::Text(parse_xl_unicode_string(&data[6..])?),
            )
        },
        _ => return invalid(record_type, "record is not a chart cache cell"),
    };
    Cache::new(kind, series, point, format, value)
}

fn write_cache(cache: &Cache) -> Result<(u16, Vec<u8>)> {
    let mut data = Vec::new();
    data.extend(cache.point.to_le_bytes());
    data.extend(u16::from(cache.series).to_le_bytes());
    data.extend(cache.format.to_le_bytes());
    let record_type = match &cache.value {
        Value::Number(number) => {
            if !number.is_finite() {
                return invalid(NUMBER, "cached chart number must be finite");
            }
            data.extend(number.to_le_bytes());
            NUMBER
        },
        Value::Text(text) => {
            data.extend(xl_unicode_string(text)?);
            LABEL
        },
        Value::Blank => BLANK,
    };
    Ok((record_type, data))
}

fn write_link(link: &DataLink, limits: Limits) -> Result<Vec<u8>> {
    validate_link(link, limits)?;
    let mut data = vec![link.role as u8, link.source as u8];
    data.extend(u16::from(link.unlinked_number_format).to_le_bytes());
    data.extend(link.number_format.to_le_bytes());
    data.extend(
        u16::try_from(link.formula_tokens.len())
            .map_err(|_| Error::InvalidData("chart formula exceeds u16".into()))?
            .to_le_bytes(),
    );
    data.extend(&link.formula_tokens);
    Ok(data)
}

/// Patches only typed series links and cache cells in an existing chart stream.
///
/// Every other record is copied from the source allocation, including unknown
/// records and their original order. The semantic models must have identical
/// chart topology; this function is deliberately not a general chart writer.
pub(crate) fn patch_chart_data(
    input: &[u8],
    old: &Chart,
    new: &Chart,
    limits: Limits,
) -> Result<Vec<u8>> {
    old.validate_semantics(limits)?;
    new.validate_semantics(limits)?;
    ensure_data_edit_shape(old, new)?;
    let source_chart = parse_chart(input, limits)?;
    if !source_chart.eq(old) {
        return Err(Error::UnsafeEdit(
            "chart source differs from its typed snapshot".into(),
        ));
    }

    let mut link_changes = old
        .series
        .iter()
        .map(|series| vec![false; series.links.len()])
        .collect::<Vec<_>>();
    for (series_index, (old_series, new_series)) in old.series.iter().zip(&new.series).enumerate() {
        for (link_index, (old_link, new_link)) in
            old_series.links.iter().zip(&new_series.links).enumerate()
        {
            if old_link.formula_tokens == new_link.formula_tokens
                && old_link.references == new_link.references
            {
                continue;
            }
            if new_link.source != Source::Cells {
                return Err(Error::UnsafeEdit(
                    "only worksheet-cell chart formulas can be patched".into(),
                ));
            }
            Formula::parse(new_link.formula_tokens.clone(), limits)?;
            link_changes[series_index][link_index] = true;
        }
    }

    let mut cache_changes = vec![false; old.cached_values.len()];
    for (index, (old_cache, new_cache)) in
        old.cached_values.iter().zip(&new.cached_values).enumerate()
    {
        if old_cache.format != new_cache.format || old_cache.value != new_cache.value {
            cache_changes[index] = true;
        }
    }
    if !link_changes.iter().flatten().any(|value| *value)
        && !cache_changes.iter().any(|value| *value)
    {
        return Ok(input.to_vec());
    }

    let records = ranges_with(input, limits.max_records_per_chart)?;
    if records.first().is_none_or(|value| {
        value.kind != BOF || !is_chart_bof(&input[value.body_start..value.body_end])
    }) || records.last().is_none_or(|value| value.kind != EOF)
    {
        return invalid(BOF, "chart data patch requires a complete chart substream");
    }
    let mut output = Vec::with_capacity(input.len());
    let mut depth = 0usize;
    let mut current_series = None;
    let mut series_depth = None;
    let mut seen_series = 0usize;
    let mut seen_links = vec![0usize; old.series.len()];
    let mut cache_kind = None;
    let mut seen_caches = 0usize;

    for value in records {
        let encoded = &input[value.start..value.end];
        let data = &input[value.body_start..value.body_end];
        match value.kind {
            BEGIN => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| Error::InvalidData("chart nesting overflow".into()))?;
                output.extend_from_slice(encoded);
            },
            END => {
                if depth == 0 {
                    return invalid(END, "unbalanced End record in chart data patch");
                }
                if series_depth == Some(depth) {
                    current_series = None;
                    series_depth = None;
                }
                depth -= 1;
                output.extend_from_slice(encoded);
            },
            SERIES => {
                if seen_series >= old.series.len() {
                    return Err(Error::UnsafeEdit(
                        "chart source contains an unmodeled series record".into(),
                    ));
                }
                current_series = Some(seen_series);
                series_depth = Some(
                    depth
                        .checked_add(1)
                        .ok_or_else(|| Error::InvalidData("chart nesting overflow".into()))?,
                );
                seen_series += 1;
                output.extend_from_slice(encoded);
            },
            BRAI => {
                let Some(series_index) = current_series else {
                    output.extend_from_slice(encoded);
                    continue;
                };
                let link_index = *seen_links.get(series_index).ok_or_else(|| {
                    Error::InvalidData("chart series link state is missing".into())
                })?;
                *seen_links.get_mut(series_index).ok_or_else(|| {
                    Error::InvalidData("chart series link state disappeared".into())
                })? += 1;
                let old_link = old
                    .series
                    .get(series_index)
                    .and_then(|series| series.links.get(link_index))
                    .ok_or_else(|| Error::UnsafeEdit("chart source link is not modeled".into()))?;
                let parsed = parse_link(data, limits)?;
                if parsed != *old_link {
                    return Err(Error::UnsafeEdit(
                        "chart source link differs from its typed snapshot".into(),
                    ));
                }
                let changed = link_changes
                    .get(series_index)
                    .and_then(|links| links.get(link_index))
                    .copied()
                    .ok_or_else(|| Error::UnsafeEdit("chart source link is not modeled".into()))?;
                if changed {
                    let link = &new.series[series_index].links[link_index];
                    output.extend(record(BRAI, &write_link(link, limits)?)?);
                    link_changes[series_index][link_index] = false;
                } else {
                    output.extend_from_slice(encoded);
                }
            },
            SI_INDEX => {
                exact(data, 2, SI_INDEX)?;
                cache_kind = Some(CacheKind::from_wire(u16_at(data, 0)?).ok_or_else(|| {
                    invalid_error(SI_INDEX, "SIIndex selects an undefined cache section")
                })?);
                output.extend_from_slice(encoded);
            },
            BLANK | NUMBER | LABEL => {
                let kind = cache_kind
                    .ok_or_else(|| invalid_error(SI_INDEX, "cache data precedes SIIndex"))?;
                let parsed = parse_cache(kind, value.kind, data)?;
                let cache_index = seen_caches;
                let old_cache = old
                    .cached_values
                    .get(cache_index)
                    .ok_or_else(|| Error::UnsafeEdit("chart source cache is not modeled".into()))?;
                if parsed != *old_cache {
                    return Err(Error::UnsafeEdit(
                        "chart source cache differs from its typed snapshot".into(),
                    ));
                }
                seen_caches = seen_caches
                    .checked_add(1)
                    .ok_or_else(|| Error::InvalidData("chart cache count overflow".into()))?;
                let changed = cache_changes
                    .get(cache_index)
                    .copied()
                    .ok_or_else(|| Error::UnsafeEdit("chart source cache is not modeled".into()))?;
                if changed {
                    let (record_type, data) = write_cache(&new.cached_values[cache_index])?;
                    output.extend(record(record_type, &data)?);
                    cache_changes[cache_index] = false;
                } else {
                    output.extend_from_slice(encoded);
                }
            },
            _ => output.extend_from_slice(encoded),
        }
    }
    if depth != 0
        || current_series.is_some()
        || series_depth.is_some()
        || seen_series != old.series.len()
        || seen_caches != old.cached_values.len()
        || seen_links
            .iter()
            .enumerate()
            .any(|(index, count)| *count != old.series[index].links.len())
        || link_changes.iter().flatten().any(|value| *value)
        || cache_changes.iter().any(|value| *value)
    {
        return Err(Error::UnsafeEdit(
            "chart data edit did not find every typed target record".into(),
        ));
    }
    Ok(output)
}

fn ensure_data_edit_shape(old: &Chart, new: &Chart) -> Result<()> {
    if old.x != new.x
        || old.y != new.y
        || old.width != new.width
        || old.height != new.height
        || old.sheet_properties != new.sheet_properties
        || old.plot_area_present != new.plot_area_present
        || old.title != new.title
        || old.groups != new.groups
        || old.axes != new.axes
        || old.legend != new.legend
        || old.formatting != new.formatting
        || old.data_labels != new.data_labels
        || old.unknown_records != new.unknown_records
        || old.series.len() != new.series.len()
        || old.cached_values.len() != new.cached_values.len()
    {
        return Err(Error::UnsafeEdit(
            "chart data edit would change an unsupported graph structure".into(),
        ));
    }
    for (old_series, new_series) in old.series.iter().zip(&new.series) {
        if old_series.category_kind != new_series.category_kind
            || old_series.category_count != new_series.category_count
            || old_series.value_count != new_series.value_count
            || old_series.bubble_count != new_series.bubble_count
            || old_series.chart_group != new_series.chart_group
            || old_series.name != new_series.name
            || old_series.links.len() != new_series.links.len()
        {
            return Err(Error::UnsafeEdit(
                "chart data edit would change series topology".into(),
            ));
        }
        for (old_link, new_link) in old_series.links.iter().zip(&new_series.links) {
            if old_link.role != new_link.role
                || old_link.source != new_link.source
                || old_link.unlinked_number_format != new_link.unlinked_number_format
                || old_link.number_format != new_link.number_format
            {
                return Err(Error::UnsafeEdit(
                    "chart data edit would change a link's semantic metadata".into(),
                ));
            }
        }
    }
    for (old_cache, new_cache) in old.cached_values.iter().zip(&new.cached_values) {
        if old_cache.kind != new_cache.kind
            || old_cache.point != new_cache.point
            || old_cache.series != new_cache.series
        {
            return Err(Error::UnsafeEdit(
                "chart data edit would change cache placement".into(),
            ));
        }
    }
    Ok(())
}

pub(super) fn validate_link(link: &DataLink, limits: Limits) -> Result<()> {
    if link.formula_tokens.len() > limits.max_formula_bytes {
        return invalid(BRAI, "BRAI formula length exceeds the configured limit");
    }
    if link.source == Source::Automatic && !link.formula_tokens.is_empty() {
        return invalid(BRAI, "automatic BRAI must have an empty formula");
    }
    if parse_chart_references(&link.formula_tokens)? != link.references {
        return invalid(
            BRAI,
            "BRAI references do not match its inert formula tokens",
        );
    }
    for value in &link.references {
        if value.first_row > value.last_row
            || value.first_column > value.last_column
            || value.last_column > 255
        {
            return invalid(
                BRAI,
                "chart formula reference is outside workbook or BIFF8 grid bounds",
            );
        }
    }
    Ok(())
}

pub(crate) fn parse_link(data: &[u8], limits: Limits) -> Result<DataLink> {
    if data.len() < 8 {
        return invalid(BRAI, "BRAI is truncated");
    }
    let flags = u16_at(data, 2)?;
    if flags & !1 != 0 {
        return invalid(BRAI, "BRAI reserved flags are nonzero");
    }
    let len = usize::from(u16_at(data, 6)?);
    if data.len() != 8 + len {
        return invalid(BRAI, "BRAI formula length mismatch");
    }
    let formula_tokens = data[8..].to_vec();
    let references = parse_chart_references(&formula_tokens)?;
    let value = DataLink {
        role: match data[0] {
            0 => Role::Name,
            1 => Role::Values,
            2 => Role::Categories,
            3 => Role::Bubbles,
            _ => return invalid(BRAI, "BRAI role is invalid"),
        },
        source: match data[1] {
            0 => Source::Automatic,
            1 => Source::Literal,
            2 => Source::Cells,
            _ => return invalid(BRAI, "BRAI source kind is invalid"),
        },
        unlinked_number_format: flags & 1 != 0,
        number_format: u16_at(data, 4)?,
        formula_tokens,
        references,
    };
    validate_link(&value, limits)?;
    Ok(value)
}

pub(crate) fn parse_chart_references(tokens: &[u8]) -> Result<Vec<CellRef>> {
    if tokens.is_empty() {
        return Ok(Vec::new());
    }
    let opcode = tokens[0] & 0x1f;
    match (opcode, tokens.len()) {
        (0x1a, 7) => {
            let col = u16_at(tokens, 5)? & 0x3fff;
            Ok(vec![CellRef {
                extern_sheet_index: u16_at(tokens, 1)?,
                first_row: u16_at(tokens, 3)?,
                last_row: u16_at(tokens, 3)?,
                first_column: col,
                last_column: col,
            }])
        },
        (0x1b, 11) => Ok(vec![CellRef {
            extern_sheet_index: u16_at(tokens, 1)?,
            first_row: u16_at(tokens, 3)?,
            last_row: u16_at(tokens, 5)?,
            first_column: u16_at(tokens, 7)? & 0x3fff,
            last_column: u16_at(tokens, 9)? & 0x3fff,
        }]),
        _ => Ok(Vec::new()),
    }
}

pub(crate) fn parse_group(kind: u16, data: &[u8]) -> Result<GroupKind> {
    Ok(match kind {
        BAR => {
            exact(data, 6, BAR)?;
            GroupKind::Bar {
                overlap: i16_at(data, 0)?,
                gap: u16_at(data, 2)?,
                flags: u16_at(data, 4)?,
            }
        },
        LINE => {
            exact(data, 2, LINE)?;
            GroupKind::Line {
                flags: u16_at(data, 0)?,
            }
        },
        AREA => {
            exact(data, 2, AREA)?;
            GroupKind::Area {
                flags: u16_at(data, 0)?,
            }
        },
        PIE => {
            exact(data, 6, PIE)?;
            GroupKind::Pie {
                rotation: u16_at(data, 0)?,
                hole_size: u16_at(data, 2)?,
                flags: u16_at(data, 4)?,
            }
        },
        SCATTER => {
            exact(data, 6, SCATTER)?;
            GroupKind::Scatter {
                bubble_size_percent: u16_at(data, 0)?,
                bubble_size_type: u16_at(data, 2)?,
                flags: u16_at(data, 4)?,
            }
        },
        RADAR | RADAR_AREA => {
            exact(data, 2, kind)?;
            GroupKind::Radar {
                filled: kind == RADAR_AREA,
                flags: u16_at(data, 0)?,
            }
        },
        SURFACE => {
            exact(data, 2, SURFACE)?;
            GroupKind::Surface {
                flags: u16_at(data, 0)?,
            }
        },
        _ => return invalid(kind, "unsupported chart group record"),
    })
}

#[cfg(test)]
pub(super) fn serialize_chart(chart: &Chart, limits: Limits) -> Result<Vec<u8>> {
    chart.validate(limits)?;
    if !chart.unknown_records.is_empty() {
        return Err(Error::UnsafeEdit(
            "opaque chart records have no proven canonical placement".to_string(),
        ));
    }
    let mut out = chart_encoder(limits)?;
    push_record(&mut out, BOF, &chart_bof())?;
    let mut geometry = Vec::new();
    for value in [chart.x, chart.y, chart.width, chart.height] {
        geometry.extend(value.to_le_bytes());
    }
    push_record(&mut out, CHART, &geometry)?;
    push_record(&mut out, BEGIN, &[])?;
    push_record(&mut out, SHT_PROPS, &chart.sheet_properties.to_le_bytes())?;
    for series in &chart.series {
        let mut body = Vec::new();
        body.extend(
            (if series.category_kind == DataKind::Numeric {
                1u16
            } else {
                3
            })
            .to_le_bytes(),
        );
        body.extend(1u16.to_le_bytes());
        body.extend(series.category_count.to_le_bytes());
        body.extend(series.value_count.to_le_bytes());
        body.extend(1u16.to_le_bytes());
        body.extend(series.bubble_count.to_le_bytes());
        push_record(&mut out, SERIES, &body)?;
        push_record(&mut out, BEGIN, &[])?;
        for link in &series.links {
            push_record(&mut out, BRAI, &write_link(link, limits)?)?;
        }
        if let Some(name) = &series.name {
            push_record(&mut out, SERIES_TEXT, &short_text(name)?)?;
        }
        push_record(&mut out, SER_TO_CRT, &series.chart_group.to_le_bytes())?;
        push_record(&mut out, END, &[])?;
    }
    push_record(
        &mut out,
        AXES_USED,
        &(if chart.groups.len() > 1 { 2u16 } else { 1 }).to_le_bytes(),
    )?;
    push_record(&mut out, AXIS_PARENT, &[0; 18])?;
    push_record(&mut out, BEGIN, &[])?;
    for axis in &chart.axes {
        let mut body = vec![0; 18];
        body[..2].copy_from_slice(
            &(match axis.kind {
                AxisKind::CategoryOrHorizontal => 0u16,
                AxisKind::ValueOrVertical => 1,
                AxisKind::Series => 2,
            })
            .to_le_bytes(),
        );
        push_record(&mut out, AXIS, &body)?;
        push_record(&mut out, BEGIN, &[])?;
        if let Some(scale) = &axis.scale {
            let mut data = Vec::new();
            for value in [
                scale.minimum,
                scale.maximum,
                scale.major,
                scale.minor,
                scale.crossing,
            ] {
                data.extend(value.to_le_bytes());
            }
            data.extend(scale.flags.to_le_bytes());
            push_record(&mut out, VALUE_RANGE, &data)?;
        }
        if let Some(tick) = &axis.tick {
            let mut data = vec![0; 26];
            data[0] = tick.major;
            data[1] = tick.minor;
            data[2] = tick.label_position;
            data[3] = tick.background;
            data[4..8].copy_from_slice(&tick.color);
            data[24..26].copy_from_slice(&tick.flags.to_le_bytes());
            push_record(&mut out, TICK, &data)?;
        }
        for line in &axis.lines {
            let id = match line.kind {
                AxisLineKind::Axis => 0u16,
                AxisLineKind::MajorGridlines => 1,
                AxisLineKind::MinorGridlines => 2,
                AxisLineKind::WallsOrFloor => 3,
            };
            push_record(&mut out, AXIS_LINE, &id.to_le_bytes())?;
            push_record(&mut out, LINE_FORMAT, &write_line(&line.format))?;
        }
        push_record(&mut out, END, &[])?;
    }
    for group in &chart.groups {
        let mut data = vec![0; 20];
        data[16..18].copy_from_slice(&u16::from(group.vary_colors).to_le_bytes());
        data[18..20].copy_from_slice(&group.order.to_le_bytes());
        push_record(&mut out, CHART_FORMAT, &data)?;
        push_record(&mut out, BEGIN, &[])?;
        write_group(&mut out, group)?;
        push_record(&mut out, END, &[])?;
    }
    if chart.plot_area_present {
        push_record(&mut out, PLOT_AREA, &[])?;
    }
    if let Some(legend) = &chart.legend {
        let mut data = Vec::new();
        for v in [legend.x, legend.y, legend.width, legend.height] {
            data.extend(v.to_le_bytes());
        }
        data.push(legend.position);
        data.push(legend.spacing);
        data.extend(legend.flags.to_le_bytes());
        push_record(&mut out, LEGEND, &data)?;
    }
    if let Some(title) = &chart.title {
        push_record(&mut out, SERIES_TEXT, &short_text(title)?)?;
    }
    for format in &chart.formatting {
        match format {
            Format::Line(value) => push_record(&mut out, LINE_FORMAT, &write_line(value))?,
            Format::Area(value) => push_record(&mut out, AREA_FORMAT, &write_area(value))?,
            Format::Marker { data } => push_record(&mut out, MARKER_FORMAT, data)?,
            Format::Data {
                point,
                series,
                flags,
            } => {
                let mut data = Vec::new();
                data.extend(point.to_le_bytes());
                data.extend(series.to_le_bytes());
                data.extend(0u16.to_le_bytes());
                data.extend(flags.to_le_bytes());
                push_record(&mut out, DATA_FORMAT, &data)?;
            },
            Format::Pie(value) => push_record(&mut out, PIE_FORMAT, &value.payload())?,
        }
    }
    for label in &chart.data_labels {
        push_record(&mut out, label.record_type, &label.data)?;
    }
    let mut active_cache = None;
    for value in &chart.cached_values {
        if active_cache != Some(value.kind) {
            push_record(&mut out, SI_INDEX, &value.kind.wire().to_le_bytes())?;
            active_cache = Some(value.kind);
        }
        let (record_type, data) = write_cache(value)?;
        push_record(&mut out, record_type, &data)?;
    }
    for value in &chart.unknown_records {
        if !known_record(value.record_type) {
            push_record(&mut out, value.record_type, &value.data)?;
        }
    }
    push_record(&mut out, END, &[])?;
    push_record(&mut out, END, &[])?;
    push_record(&mut out, EOF, &[])?;
    Ok(out.finish())
}

#[cfg(test)]
pub(super) fn write_group(out: &mut GraphEncoder, group: &Group) -> Result<()> {
    match &group.kind {
        GroupKind::Line { flags } => push_record(out, LINE, &flags.to_le_bytes())?,
        GroupKind::Area { flags } => push_record(out, AREA, &flags.to_le_bytes())?,
        GroupKind::Bar {
            overlap,
            gap,
            flags,
        } => {
            let mut d = overlap.to_le_bytes().to_vec();
            d.extend(gap.to_le_bytes());
            d.extend(flags.to_le_bytes());
            push_record(out, BAR, &d)?;
        },
        GroupKind::Pie {
            rotation,
            hole_size,
            flags,
        } => {
            let mut d = rotation.to_le_bytes().to_vec();
            d.extend(hole_size.to_le_bytes());
            d.extend(flags.to_le_bytes());
            push_record(out, PIE, &d)?;
        },
        GroupKind::Scatter {
            bubble_size_percent,
            bubble_size_type,
            flags,
        } => {
            let mut d = bubble_size_percent.to_le_bytes().to_vec();
            d.extend(bubble_size_type.to_le_bytes());
            d.extend(flags.to_le_bytes());
            push_record(out, SCATTER, &d)?;
        },
        GroupKind::Radar { filled, flags } => push_record(
            out,
            if *filled { RADAR_AREA } else { RADAR },
            &flags.to_le_bytes(),
        )?,
        GroupKind::Surface { flags } => push_record(out, SURFACE, &flags.to_le_bytes())?,
    }
    for value in &group.lines {
        push_record(out, CRT_LINE, &line::Line::new(value.kind).payload())?;
        push_record(out, LINE_FORMAT, &shared_line_bytes(&value.format))?;
    }
    for value in &group.drop_bars {
        push_record(out, DROP_BAR, &value.gap.get().to_le_bytes())?;
        push_record(out, BEGIN, &[])?;
        push_record(out, LINE_FORMAT, &shared_line_bytes(&value.line))?;
        push_record(out, AREA_FORMAT, &shared_area_bytes(&value.area))?;
        push_record(out, END, &[])?;
    }
    Ok(())
}
