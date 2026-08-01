use std::char;

use super::axis::{self, Axis};
use super::format::{self, Format};
use super::group;
use super::model::{
    Cache, Cell, CellRef, Chart, Context, Count, DataKind, Family, Group, GroupId, Label, Legend,
    Link, Origin, Props, Raw, Rect, Role, RowCol, Series, Source, Value,
};
use super::{
    BOF, BOF_BYTES, EOF, EXCEL_DOC_TYPE, EXCEL_VERSION, GRAPH_DOC_TYPE, GRAPH_VERSION, Kind, Ref,
};
use crate::limits::as_u64;
use crate::raw::{Encoder, Kind as RecordKind, RecordRef};
use crate::{Error, Limits, Result};

const CONTINUE: RecordKind = RecordKind::new(0x003C);
const GRAPH_BLANK: RecordKind = RecordKind::new(0x0001);
const GRAPH_NUMBER: RecordKind = RecordKind::new(0x0003);
const EXCEL_BLANK: RecordKind = RecordKind::new(0x0201);
const EXCEL_NUMBER: RecordKind = RecordKind::new(0x0203);
const CELL_LABEL: RecordKind = RecordKind::new(0x0204);
const DATA_LAB_EXT: RecordKind = RecordKind::new(0x086A);
const DATA_LAB_EXT_CONTENTS: RecordKind = RecordKind::new(0x086B);
const CHART_REC: RecordKind = RecordKind::new(0x1002);
const SERIES: RecordKind = RecordKind::new(0x1003);
const DATA_FORMAT: RecordKind = RecordKind::new(0x1006);
const LINE_FORMAT: RecordKind = RecordKind::new(0x1007);
const MARKER_FORMAT: RecordKind = RecordKind::new(0x1009);
const AREA_FORMAT: RecordKind = RecordKind::new(0x100A);
const PIE_FORMAT: RecordKind = RecordKind::new(0x100B);
const SERIES_TEXT: RecordKind = RecordKind::new(0x100D);
const CHART_FORMAT: RecordKind = RecordKind::new(0x1014);
const LEGEND: RecordKind = RecordKind::new(0x1015);
const SERIES_LIST: RecordKind = RecordKind::new(0x1016);
const BAR: RecordKind = RecordKind::new(0x1017);
const LINE: RecordKind = RecordKind::new(0x1018);
const PIE: RecordKind = RecordKind::new(0x1019);
const AREA: RecordKind = RecordKind::new(0x101A);
const SCATTER: RecordKind = RecordKind::new(0x101B);
const CRT_LINE: RecordKind = RecordKind::new(0x101C);
const AXIS: RecordKind = RecordKind::new(0x101D);
const TICK: RecordKind = RecordKind::new(0x101E);
const VALUE_RANGE: RecordKind = RecordKind::new(0x101F);
const CAT_SER_RANGE: RecordKind = RecordKind::new(0x1020);
const AXIS_LINE: RecordKind = RecordKind::new(0x1021);
const DEFAULT_TEXT: RecordKind = RecordKind::new(0x1024);
const TEXT: RecordKind = RecordKind::new(0x1025);
const FONT_X: RecordKind = RecordKind::new(0x1026);
const OBJECT_LINK: RecordKind = RecordKind::new(0x1027);
const FRAME: RecordKind = RecordKind::new(0x1032);
const BEGIN: RecordKind = RecordKind::new(0x1033);
const END: RecordKind = RecordKind::new(0x1034);
const PLOT_AREA: RecordKind = RecordKind::new(0x1035);
const DROP_BAR: RecordKind = RecordKind::new(0x103D);
const RADAR: RecordKind = RecordKind::new(0x103E);
const SURFACE: RecordKind = RecordKind::new(0x103F);
const RADAR_AREA: RecordKind = RecordKind::new(0x1040);
const AXIS_PARENT: RecordKind = RecordKind::new(0x1041);
const SHT_PROPS: RecordKind = RecordKind::new(0x1044);
const SER_TO_CRT: RecordKind = RecordKind::new(0x1045);
const AXES_USED: RecordKind = RecordKind::new(0x1046);
const BRAI: RecordKind = RecordKind::new(0x1051);
const PLOT_GROWTH: RecordKind = RecordKind::new(0x1064);
const SI_INDEX: RecordKind = RecordKind::new(0x1065);

#[derive(Clone, Copy)]
enum PendingLine {
    Axis {
        owner: usize,
        kind: axis::LineKind,
    },
    Group {
        owner: usize,
        kind: crate::record::line::Kind,
    },
}

struct PendingDrop {
    owner: usize,
    depth: usize,
    gap: group::Gap,
    line: Option<format::Line>,
    area: Option<format::Area>,
}

pub(super) fn parse(input: Ref<'_>, context: Context, limits: Limits) -> Result<Chart> {
    if input.kind() != context.kind() {
        return invalid_model("context", "producer kind does not match the chart BOF");
    }

    let mut chart = Chart {
        context,
        rect: Rect::default(),
        props: Props {
            flags: 2,
            plot_area: false,
        },
        title: None,
        series: Vec::new(),
        groups: Vec::new(),
        axes: Vec::new(),
        legend: None,
        caches: Vec::new(),
        formats: Vec::new(),
        labels: Vec::new(),
        unknown: Vec::new(),
        origin: Origin::Fresh,
        dirty: false,
        limits,
        authoring_proven: false,
    };
    let mut depth = 0usize;
    let mut current_series = None;
    let mut series_depth = None;
    let mut current_axis = None;
    let mut axis_depth = None;
    let mut group_depth = None;
    let mut pending_begin = false;
    let mut cache_index = 0u16;
    let mut pending_axis_line = None;
    let mut pending_drop: Option<PendingDrop> = None;
    let mut unknown_bytes = 0usize;
    let mut first = true;

    for item in input.records() {
        let record = item?;
        if record.payload().len() > limits.max_record_bytes {
            return limit(
                "record bytes",
                record.payload().len(),
                limits.max_record_bytes,
            );
        }
        if first {
            first = false;
            if record.kind() != BOF {
                return invalid(record, "chart substream does not start with BOF");
            }
            continue;
        }
        if pending_axis_line.is_some() && record.kind() != LINE_FORMAT {
            return invalid(
                record,
                "line owner is not followed immediately by LineFormat",
            );
        }
        if let Some(drop) = &pending_drop
            && depth == drop.depth
        {
            if drop.line.is_none() && record.kind() != LINE_FORMAT {
                return invalid(record, "DropBar Begin is not followed by LineFormat");
            }
            if drop.line.is_some() && drop.area.is_none() && record.kind() != AREA_FORMAT {
                return invalid(record, "DropBar LineFormat is not followed by AreaFormat");
            }
        }
        if record.kind() == EOF {
            break;
        }
        if pending_begin && record.kind() != BEGIN {
            return invalid(
                record,
                "collection-owning record is not followed immediately by Begin",
            );
        }
        let data = record.payload();
        match record.kind() {
            BEGIN => {
                exact(record, 0)?;
                pending_begin = false;
                depth = depth.checked_add(1).ok_or(Error::SizeOverflow {
                    resource: "chart nesting",
                })?;
                if depth > limits.max_nesting {
                    return limit("chart nesting", depth, limits.max_nesting);
                }
            },
            END => {
                exact(record, 0)?;
                if depth == 0 {
                    return invalid(record, "End record has no matching Begin");
                }
                if pending_drop
                    .as_ref()
                    .is_some_and(|drop| drop.depth == depth)
                {
                    let drop = pending_drop.take().ok_or(Error::InvalidChart {
                        offset: record.offset(),
                        reason: "DropBar collection state disappeared",
                    })?;
                    let line = drop.line.ok_or(Error::InvalidChart {
                        offset: record.offset(),
                        reason: "DropBar collection has no LineFormat",
                    })?;
                    let area = drop.area.ok_or(Error::InvalidChart {
                        offset: record.offset(),
                        reason: "DropBar collection has no AreaFormat",
                    })?;
                    let group = chart
                        .groups
                        .get_mut(drop.owner)
                        .ok_or(Error::InvalidChart {
                            offset: record.offset(),
                            reason: "DropBar owner is missing",
                        })?;
                    push(
                        &mut group.drop_bars,
                        group::DropBar {
                            gap: drop.gap,
                            line,
                            area,
                        },
                        "chart drop bars",
                    )?;
                }
                if series_depth == Some(depth) {
                    current_series = None;
                    series_depth = None;
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
            CHART_REC => {
                exact(record, 16)?;
                chart.rect = Rect {
                    x: i32_at(data, 0, record)?,
                    y: i32_at(data, 4, record)?,
                    width: i32_at(data, 8, record)?,
                    height: i32_at(data, 12, record)?,
                };
                pending_begin = true;
            },
            SHT_PROPS => {
                exact(record, 4)?;
                let flags = u32_at(data, 0, record)?;
                if !valid_props(flags) {
                    return invalid(record, "ShtProps reserved bits or blank mode are invalid");
                }
                chart.props.flags = flags;
            },
            SERIES => {
                if current_series.is_some() {
                    return invalid(record, "Series collections overlap");
                }
                exact(record, 12)?;
                check_add(chart.series.len(), limits.max_series, "series count")?;
                let category_kind = match u16_at(data, 0, record)? {
                    1 => DataKind::Numeric,
                    3 => DataKind::Text,
                    _ => return invalid(record, "invalid category data type"),
                };
                if u16_at(data, 2, record)? != 1 || u16_at(data, 8, record)? != 1 {
                    return invalid(record, "series numeric data type fields are invalid");
                }
                let series = Series {
                    category_kind,
                    category_count: count_at(data, 4, record)?,
                    value_count: count_at(data, 6, record)?,
                    bubble_count: count_at(data, 10, record)?,
                    group: GroupId::ZERO,
                    name: None,
                    links: Vec::new(),
                };
                push(&mut chart.series, series, "chart series")?;
                current_series = chart.series.len().checked_sub(1);
                series_depth = depth.checked_add(1);
                pending_begin = true;
            },
            BRAI => {
                if series_depth != Some(depth) {
                    return invalid(record, "BRAI appears outside a Series collection");
                }
                let link = parse_link(record, context, limits)?;
                let index = current_series.ok_or(Error::InvalidChart {
                    offset: record.offset(),
                    reason: "BRAI appears outside a Series collection",
                })?;
                let series = chart.series.get_mut(index).ok_or(Error::InvalidChart {
                    offset: record.offset(),
                    reason: "BRAI refers to a missing Series",
                })?;
                push(&mut series.links, link, "series links")?;
            },
            SER_TO_CRT => {
                exact(record, 2)?;
                if series_depth != Some(depth) {
                    return invalid(record, "SerToCrt appears outside a Series collection");
                }
                let index = current_series.ok_or(Error::InvalidChart {
                    offset: record.offset(),
                    reason: "SerToCrt appears outside a Series collection",
                })?;
                let raw = u16_at(data, 0, record)?;
                let raw = u8::try_from(raw).map_err(|_| Error::InvalidChart {
                    offset: record.offset(),
                    reason: "series chart-group index exceeds nine",
                })?;
                let group = GroupId::new(raw).ok_or(Error::InvalidChart {
                    offset: record.offset(),
                    reason: "series chart-group index exceeds nine",
                })?;
                chart
                    .series
                    .get_mut(index)
                    .ok_or(Error::InvalidChart {
                        offset: record.offset(),
                        reason: "SerToCrt refers to a missing Series",
                    })?
                    .group = group;
            },
            SERIES_TEXT => {
                let text = parse_short_text(record)?;
                if let Some(series) = current_series
                    .filter(|_| series_depth == Some(depth))
                    .and_then(|index| chart.series.get_mut(index))
                    && series.name.is_none()
                {
                    series.name = Some(text);
                } else if chart.title.is_none() {
                    chart.title = Some(text);
                }
            },
            CHART_FORMAT => {
                if group_depth.is_some() {
                    return invalid(record, "ChartFormat collections overlap");
                }
                exact(record, 20)?;
                let reserved = data.get(..16).ok_or(Error::InvalidChart {
                    offset: record.offset(),
                    reason: "ChartFormat reserved fields are truncated",
                })?;
                let vary = u16_at(data, 16, record)?;
                if reserved.iter().any(|value| *value != 0) || vary & !1 != 0 {
                    return invalid(record, "ChartFormat reserved fields are nonzero");
                }
                check_add(chart.groups.len(), limits.max_groups, "group count")?;
                let raw = u16_at(data, 18, record)?;
                let raw = u8::try_from(raw).map_err(|_| Error::InvalidChart {
                    offset: record.offset(),
                    reason: "chart-group order exceeds nine",
                })?;
                let order = GroupId::new(raw).ok_or(Error::InvalidChart {
                    offset: record.offset(),
                    reason: "chart-group order exceeds nine",
                })?;
                push(
                    &mut chart.groups,
                    Group {
                        order,
                        vary_colors: vary & 1 != 0,
                        family: Family::Line { flags: 0 },
                        lines: Vec::new(),
                        drop_bars: Vec::new(),
                    },
                    "chart groups",
                )?;
                group_depth = depth.checked_add(1);
                pending_begin = true;
            },
            BAR | LINE | PIE | AREA | SCATTER | RADAR | RADAR_AREA | SURFACE => {
                if group_depth != Some(depth) {
                    return invalid(record, "chart-family record appears outside ChartFormat");
                }
                let family = parse_family(record)?;
                chart
                    .groups
                    .last_mut()
                    .ok_or(Error::InvalidChart {
                        offset: record.offset(),
                        reason: "chart-family record has no ChartFormat owner",
                    })?
                    .family = family;
            },
            CRT_LINE => {
                if group_depth != Some(depth) {
                    return invalid(record, "CrtLine appears outside ChartFormat");
                }
                let value = crate::record::line::Line::from_payload(data).map_err(|_| {
                    Error::InvalidChart {
                        offset: record.offset(),
                        reason: "CrtLine kind or payload is invalid",
                    }
                })?;
                let index = chart
                    .groups
                    .len()
                    .checked_sub(1)
                    .ok_or(Error::InvalidChart {
                        offset: record.offset(),
                        reason: "CrtLine has no ChartFormat owner",
                    })?;
                pending_axis_line = Some(PendingLine::Group {
                    owner: index,
                    kind: value.kind(),
                });
            },
            DROP_BAR => {
                if group_depth != Some(depth) {
                    return invalid(record, "DropBar appears outside ChartFormat");
                }
                if pending_drop.is_some() {
                    return invalid(record, "DropBar collections overlap");
                }
                exact(record, 2)?;
                let gap = group::Gap::new(u16_at(data, 0, record)?).ok_or(Error::InvalidChart {
                    offset: record.offset(),
                    reason: "DropBar gap exceeds 500",
                })?;
                let owner = chart
                    .groups
                    .len()
                    .checked_sub(1)
                    .ok_or(Error::InvalidChart {
                        offset: record.offset(),
                        reason: "DropBar has no ChartFormat owner",
                    })?;
                if chart
                    .groups
                    .get(owner)
                    .is_none_or(|group| group.drop_bars.len() >= 2)
                {
                    return invalid(
                        record,
                        "a chart group has more than two DropBar collections",
                    );
                }
                pending_drop = Some(PendingDrop {
                    owner,
                    depth: depth.checked_add(1).ok_or(Error::SizeOverflow {
                        resource: "DropBar nesting",
                    })?,
                    gap,
                    line: None,
                    area: None,
                });
                pending_begin = true;
            },
            AXIS => {
                if current_axis.is_some() {
                    return invalid(record, "Axis collections overlap");
                }
                exact(record, 18)?;
                let reserved = data.get(2..).ok_or(Error::InvalidChart {
                    offset: record.offset(),
                    reason: "Axis reserved fields are truncated",
                })?;
                if reserved.iter().any(|value| *value != 0) {
                    return invalid(record, "Axis reserved fields are nonzero");
                }
                let kind = match u16_at(data, 0, record)? {
                    0 => axis::Kind::Category,
                    1 => axis::Kind::Value,
                    2 => axis::Kind::Series,
                    _ => return invalid(record, "invalid axis kind"),
                };
                check_add(chart.axes.len(), limits.max_axes, "axis count")?;
                push(&mut chart.axes, Axis::new(kind), "chart axes")?;
                current_axis = chart.axes.len().checked_sub(1);
                axis_depth = depth.checked_add(1);
                pending_begin = true;
            },
            AXES_USED => {
                exact(record, 2)?;
                if !matches!(u16_at(data, 0, record)?, 1 | 2) {
                    return invalid(record, "AxesUsed must specify one or two axis groups");
                }
            },
            AXIS_PARENT => {
                exact(record, 18)?;
                if u16_at(data, 0, record)? > 1 {
                    return invalid(record, "AxisParent index must be primary or secondary");
                }
                pending_begin = true;
            },
            VALUE_RANGE => {
                exact(record, 42)?;
                if axis_depth != Some(depth) {
                    return invalid(record, "ValueRange appears outside an Axis collection");
                }
                let axis = current_axis
                    .and_then(|index| chart.axes.get_mut(index))
                    .ok_or(Error::InvalidChart {
                        offset: record.offset(),
                        reason: "ValueRange appears before Axis",
                    })?;
                axis.scale = Some(axis::Scale {
                    min: f64_at(data, 0, record)?,
                    max: f64_at(data, 8, record)?,
                    major: f64_at(data, 16, record)?,
                    minor: f64_at(data, 24, record)?,
                    crossing: f64_at(data, 32, record)?,
                    flags: u16_at(data, 40, record)?,
                });
            },
            TICK => {
                if data.len() < 26 {
                    return invalid(record, "Tick record is shorter than 26 bytes");
                }
                if axis_depth != Some(depth) {
                    return invalid(record, "Tick appears outside an Axis collection");
                }
                let axis = current_axis
                    .and_then(|index| chart.axes.get_mut(index))
                    .ok_or(Error::InvalidChart {
                        offset: record.offset(),
                        reason: "Tick appears before Axis",
                    })?;
                axis.tick = Some(axis::Tick {
                    major: byte_at(data, 0, record)?,
                    minor: byte_at(data, 1, record)?,
                    label: byte_at(data, 2, record)?,
                    background: byte_at(data, 3, record)?,
                    color: array4_at(data, 4, record)?,
                    flags: u16_at(data, 24, record)?,
                });
            },
            AXIS_LINE => {
                exact(record, 2)?;
                if axis_depth != Some(depth) {
                    return invalid(record, "AxisLine appears outside an Axis collection");
                }
                let kind = match u16_at(data, 0, record)? {
                    0 => axis::LineKind::Axis,
                    1 => axis::LineKind::MajorGrid,
                    2 => axis::LineKind::MinorGrid,
                    3 => axis::LineKind::Wall,
                    _ => return invalid(record, "invalid AxisLine kind"),
                };
                let index = current_axis.ok_or(Error::InvalidChart {
                    offset: record.offset(),
                    reason: "AxisLine appears before Axis",
                })?;
                if chart.axes.get(index).is_none() {
                    return Err(Error::InvalidChart {
                        offset: record.offset(),
                        reason: "AxisLine refers to a missing Axis",
                    });
                }
                pending_axis_line = Some(PendingLine::Axis { owner: index, kind });
            },
            LINE_FORMAT => {
                let value = parse_line(record)?;
                match pending_axis_line.take() {
                    Some(PendingLine::Axis { owner, kind }) => {
                        let axis = chart.axes.get_mut(owner).ok_or(Error::InvalidChart {
                            offset: record.offset(),
                            reason: "pending axis owner is missing",
                        })?;
                        push(
                            &mut axis.lines,
                            axis::Line {
                                kind,
                                format: value,
                            },
                            "axis lines",
                        )?;
                    },
                    Some(PendingLine::Group { owner, kind }) => {
                        let group = chart.groups.get_mut(owner).ok_or(Error::InvalidChart {
                            offset: record.offset(),
                            reason: "pending chart-group owner is missing",
                        })?;
                        push(
                            &mut group.lines,
                            group::Line {
                                kind,
                                format: value,
                            },
                            "chart-group lines",
                        )?;
                    },
                    None => {
                        if let Some(drop) = pending_drop.as_mut().filter(|drop| drop.depth == depth)
                        {
                            if drop.line.replace(value).is_some() {
                                return invalid(record, "DropBar has more than one LineFormat");
                            }
                        } else {
                            push(&mut chart.formats, Format::Line(value), "chart formats")?;
                        }
                    },
                }
            },
            AREA_FORMAT => {
                let value = parse_area(record)?;
                if let Some(drop) = pending_drop.as_mut().filter(|drop| drop.depth == depth) {
                    if drop.area.replace(value).is_some() {
                        return invalid(record, "DropBar has more than one AreaFormat");
                    }
                } else {
                    push(&mut chart.formats, Format::Area(value), "chart formats")?;
                }
            },
            MARKER_FORMAT => {
                let data = copy(data, "marker payload", limits.max_record_bytes)?;
                push(&mut chart.formats, Format::Marker { data }, "chart formats")?;
            },
            DATA_FORMAT => {
                exact(record, 8)?;
                push(
                    &mut chart.formats,
                    Format::Data {
                        point: u16_at(data, 0, record)?,
                        series: u16_at(data, 2, record)?,
                        flags: u16_at(data, 6, record)?,
                    },
                    "chart formats",
                )?;
            },
            PIE_FORMAT => {
                exact(record, 2)?;
                push(
                    &mut chart.formats,
                    Format::Pie {
                        explosion: u16_at(data, 0, record)?,
                    },
                    "chart formats",
                )?;
            },
            LEGEND => {
                exact(record, 20)?;
                chart.legend = Some(Legend {
                    x: i32_at(data, 0, record)?,
                    y: i32_at(data, 4, record)?,
                    width: i32_at(data, 8, record)?,
                    height: i32_at(data, 12, record)?,
                    position: byte_at(data, 16, record)?,
                    spacing: byte_at(data, 17, record)?,
                    flags: u16_at(data, 18, record)?,
                });
            },
            PLOT_AREA => {
                exact(record, 0)?;
                chart.props.plot_area = true;
            },
            DATA_LAB_EXT | DATA_LAB_EXT_CONTENTS | TEXT => {
                let data = copy(data, "data-label payload", limits.max_record_bytes)?;
                push(
                    &mut chart.labels,
                    Label {
                        kind: record.kind(),
                        data,
                    },
                    "chart labels",
                )?;
            },
            SI_INDEX if context.kind() == Kind::Excel => {
                exact(record, 2)?;
                cache_index = u16_at(data, 0, record)?;
            },
            kind if kind == number_kind(context.kind()) => {
                let (value_offset, format) = match context.kind() {
                    Kind::Graph => {
                        exact(record, 15)?;
                        if byte_at(data, 4, record)? != 0 {
                            return invalid(record, "Graph Number reserved byte is nonzero");
                        }
                        (7, u16_at(data, 5, record)?)
                    },
                    Kind::Excel => {
                        exact(record, 14)?;
                        (6, u16_at(data, 4, record)?)
                    },
                };
                check_add(
                    chart.caches.len(),
                    limits.max_cached_values,
                    "cached value count",
                )?;
                let cell = cache_cell(data, context, record)?;
                push(
                    &mut chart.caches,
                    Cache {
                        index: cache_index,
                        cell,
                        format,
                        value: Value::Number(f64_at(data, value_offset, record)?),
                    },
                    "chart cache",
                )?;
            },
            CELL_LABEL => {
                let (string_offset, format) = match context.kind() {
                    Kind::Graph => {
                        if data.len() < 9 || byte_at(data, 4, record)? != 0 {
                            return invalid(
                                record,
                                "Graph Label is truncated or reserved byte is nonzero",
                            );
                        }
                        (7, u16_at(data, 5, record)?)
                    },
                    Kind::Excel => {
                        if data.len() < 8 {
                            return invalid(record, "Excel Label is shorter than eight bytes");
                        }
                        (6, u16_at(data, 4, record)?)
                    },
                };
                check_add(
                    chart.caches.len(),
                    limits.max_cached_values,
                    "cached value count",
                )?;
                let cell = cache_cell(data, context, record)?;
                let string = data.get(string_offset..).ok_or(Error::InvalidChart {
                    offset: record.offset(),
                    reason: "cached Label string is truncated",
                })?;
                push(
                    &mut chart.caches,
                    Cache {
                        index: cache_index,
                        cell,
                        format,
                        value: Value::Text(parse_string(string, record)?),
                    },
                    "chart cache",
                )?;
            },
            kind if kind == blank_kind(context.kind()) => {
                let format = match context.kind() {
                    Kind::Graph => {
                        exact(record, 7)?;
                        if byte_at(data, 4, record)? != 0 {
                            return invalid(record, "Graph Blank reserved byte is nonzero");
                        }
                        u16_at(data, 5, record)?
                    },
                    Kind::Excel => {
                        exact(record, 6)?;
                        u16_at(data, 4, record)?
                    },
                };
                check_add(
                    chart.caches.len(),
                    limits.max_cached_values,
                    "cached value count",
                )?;
                let cell = cache_cell(data, context, record)?;
                push(
                    &mut chart.caches,
                    Cache {
                        index: cache_index,
                        cell,
                        format,
                        value: Value::Blank,
                    },
                    "chart cache",
                )?;
            },
            BOF => return invalid(record, "nested BOF in chart substream"),
            CONTINUE | SERIES_LIST | CAT_SER_RANGE | DEFAULT_TEXT | FONT_X | OBJECT_LINK
            | FRAME | PLOT_GROWTH => {
                add_raw(&mut chart, record, &mut unknown_bytes, limits)?;
            },
            _ => add_raw(&mut chart, record, &mut unknown_bytes, limits)?,
        }
    }
    if depth != 0 {
        return Err(Error::InvalidChart {
            offset: input.as_bytes().len(),
            reason: "chart Begin/End collections are unbalanced",
        });
    }
    if pending_begin
        || pending_drop.is_some()
        || current_series.is_some()
        || current_axis.is_some()
        || group_depth.is_some()
    {
        return Err(Error::InvalidChart {
            offset: input.as_bytes().len(),
            reason: "chart collection owner is missing its complete Begin/End collection",
        });
    }
    validate(&chart, limits)?;
    Ok(chart)
}

fn parse_link(record: RecordRef<'_>, context: Context, limits: Limits) -> Result<Link> {
    let data = record.payload();
    match context.kind() {
        Kind::Graph => {
            exact(record, 8)?;
            let flags = u16_at(data, 2, record)?;
            if flags & 0x0002 == 0 || flags & !0x0003 != 0 {
                return invalid(record, "Graph BRAI reserved bits are invalid");
            }
            let row_col = RowCol::new(u16_at(data, 6, record)?).ok_or(Error::InvalidChart {
                offset: record.offset(),
                reason: "Graph BRAI row or column exceeds 3,999",
            })?;
            Ok(Link::Graph {
                role: parse_role(byte_at(data, 0, record)?, record)?,
                source: parse_source(byte_at(data, 1, record)?, record)?,
                unlinked_format: flags & 1 != 0,
                number_format: u16_at(data, 4, record)?,
                row_col,
            })
        },
        Kind::Excel => {
            if data.len() < 8 {
                return invalid(record, "Excel BRAI is shorter than eight bytes");
            }
            let flags = u16_at(data, 2, record)?;
            if flags & !1 != 0 {
                return invalid(record, "Excel BRAI reserved bits are nonzero");
            }
            let formula_len = usize::from(u16_at(data, 6, record)?);
            let expected = 8usize.checked_add(formula_len).ok_or(Error::SizeOverflow {
                resource: "BRAI formula",
            })?;
            if data.len() != expected {
                return invalid(
                    record,
                    "Excel BRAI formula length does not match its payload",
                );
            }
            if formula_len > limits.max_formula_bytes {
                return limit("formula bytes", formula_len, limits.max_formula_bytes);
            }
            let tokens = data.get(8..).ok_or(Error::InvalidChart {
                offset: record.offset(),
                reason: "Excel BRAI formula is truncated",
            })?;
            let formula = copy(tokens, "formula tokens", limits.max_formula_bytes)?;
            let refs = parse_refs(tokens, record)?;
            let link = Link::Excel {
                role: parse_role(byte_at(data, 0, record)?, record)?,
                source: parse_source(byte_at(data, 1, record)?, record)?,
                unlinked_format: flags & 1 != 0,
                number_format: u16_at(data, 4, record)?,
                formula,
                refs,
            };
            if matches!(
                &link,
                Link::Excel {
                    source: Source::Automatic,
                    formula,
                    ..
                } if !formula.is_empty()
            ) {
                return invalid(record, "automatic Excel BRAI has a nonempty formula");
            }
            Ok(link)
        },
    }
}

fn parse_refs(tokens: &[u8], record: RecordRef<'_>) -> Result<Vec<CellRef>> {
    if tokens.is_empty() {
        return Ok(Vec::new());
    }
    let opcode = byte_at(tokens, 0, record)? & 0x1F;
    let value = match (opcode, tokens.len()) {
        (0x1A, 7) => {
            let col = u16_at(tokens, 5, record)? & 0x3FFF;
            let col = u8::try_from(col).map_err(|_| Error::InvalidChart {
                offset: record.offset(),
                reason: "chart formula column exceeds the BIFF8 grid",
            })?;
            Some(CellRef {
                external_sheet: u16_at(tokens, 1, record)?,
                first_row: u16_at(tokens, 3, record)?,
                last_row: u16_at(tokens, 3, record)?,
                first_col: col,
                last_col: col,
            })
        },
        (0x1B, 11) => {
            let first_col = u16_at(tokens, 7, record)? & 0x3FFF;
            let last_col = u16_at(tokens, 9, record)? & 0x3FFF;
            Some(CellRef {
                external_sheet: u16_at(tokens, 1, record)?,
                first_row: u16_at(tokens, 3, record)?,
                last_row: u16_at(tokens, 5, record)?,
                first_col: u8::try_from(first_col).map_err(|_| Error::InvalidChart {
                    offset: record.offset(),
                    reason: "chart formula column exceeds the BIFF8 grid",
                })?,
                last_col: u8::try_from(last_col).map_err(|_| Error::InvalidChart {
                    offset: record.offset(),
                    reason: "chart formula column exceeds the BIFF8 grid",
                })?,
            })
        },
        _ => None,
    };
    let Some(value) = value else {
        return Ok(Vec::new());
    };
    let mut refs = Vec::new();
    refs.try_reserve_exact(1).map_err(|_| Error::Allocation {
        resource: "chart references",
    })?;
    refs.push(value);
    Ok(refs)
}

fn parse_family(record: RecordRef<'_>) -> Result<Family> {
    let data = record.payload();
    match record.kind() {
        BAR => {
            exact(record, 6)?;
            Ok(Family::Bar {
                overlap: group::Overlap::new(i16_at(data, 0, record)?).ok_or(
                    Error::InvalidChart {
                        offset: record.offset(),
                        reason: "bar overlap is outside -100 through 100",
                    },
                )?,
                gap: group::Gap::new(u16_at(data, 2, record)?).ok_or(Error::InvalidChart {
                    offset: record.offset(),
                    reason: "bar gap exceeds 500",
                })?,
                flags: u16_at(data, 4, record)?,
            })
        },
        LINE => {
            exact(record, 2)?;
            Ok(Family::Line {
                flags: u16_at(data, 0, record)?,
            })
        },
        AREA => {
            exact(record, 2)?;
            Ok(Family::Area {
                flags: u16_at(data, 0, record)?,
            })
        },
        PIE => {
            exact(record, 6)?;
            Ok(Family::Pie {
                rotation: u16_at(data, 0, record)?,
                hole: u16_at(data, 2, record)?,
                flags: u16_at(data, 4, record)?,
            })
        },
        SCATTER => {
            exact(record, 6)?;
            Ok(Family::Scatter {
                bubble_percent: group::BubblePercent::new(u16_at(data, 0, record)?).ok_or(
                    Error::InvalidChart {
                        offset: record.offset(),
                        reason: "scatter bubble percentage exceeds 300",
                    },
                )?,
                bubble_kind: match u16_at(data, 2, record)? {
                    1 => group::BubbleKind::Area,
                    2 => group::BubbleKind::Width,
                    _ => {
                        return invalid(record, "scatter bubble-size kind is not 1 or 2");
                    },
                },
                flags: u16_at(data, 4, record)?,
            })
        },
        RADAR | RADAR_AREA => {
            exact(record, 2)?;
            Ok(Family::Radar {
                filled: record.kind() == RADAR_AREA,
                flags: u16_at(data, 0, record)?,
            })
        },
        SURFACE => {
            exact(record, 2)?;
            Ok(Family::Surface {
                flags: u16_at(data, 0, record)?,
            })
        },
        _ => invalid(record, "unsupported chart-family record"),
    }
}

fn parse_line(record: RecordRef<'_>) -> Result<format::Line> {
    exact(record, 12)?;
    let data = record.payload();
    Ok(format::Line {
        color: array4_at(data, 0, record)?,
        pattern: u16_at(data, 4, record)?,
        weight: i16_at(data, 6, record)?,
        flags: u16_at(data, 8, record)?,
        color_index: u16_at(data, 10, record)?,
    })
}

fn parse_area(record: RecordRef<'_>) -> Result<format::Area> {
    exact(record, 16)?;
    let data = record.payload();
    Ok(format::Area {
        foreground: array4_at(data, 0, record)?,
        background: array4_at(data, 4, record)?,
        pattern: u16_at(data, 8, record)?,
        flags: u16_at(data, 10, record)?,
        foreground_index: u16_at(data, 12, record)?,
        background_index: u16_at(data, 14, record)?,
    })
}

fn add_raw(
    chart: &mut Chart,
    record: RecordRef<'_>,
    unknown_bytes: &mut usize,
    limits: Limits,
) -> Result<()> {
    *unknown_bytes =
        unknown_bytes
            .checked_add(record.payload().len())
            .ok_or(Error::SizeOverflow {
                resource: "unknown chart bytes",
            })?;
    if *unknown_bytes > limits.max_unknown_bytes {
        return limit(
            "unknown chart bytes",
            *unknown_bytes,
            limits.max_unknown_bytes,
        );
    }
    let data = copy(
        record.payload(),
        "unknown chart payload",
        limits.max_record_bytes,
    )?;
    push(
        &mut chart.unknown,
        Raw::parsed(record.kind(), data, record.offset()),
        "unknown chart records",
    )
}

fn validate(chart: &Chart, limits: Limits) -> Result<()> {
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

    let mut orders = [false; 10];
    for group in &chart.groups {
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
        if usize::from(series.group.get()) >= chart.groups.len() {
            return invalid_model("series", "series refers to a missing chart group");
        }
        check_string(series.name.as_deref(), "series name")?;
        for link in &series.links {
            validate_link(link, chart.context, limits)?;
        }
    }
    check_string(chart.title.as_deref(), "title")?;

    for axis in &chart.axes {
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
        if !cache_matches(value.cell, chart.context.kind()) {
            return invalid_model("cache", "cell coordinate does not match producer context");
        }
        if chart.context.kind() == Kind::Graph && value.index != 0 {
            return invalid_model("cache", "Graph cache entries do not use SIIndex");
        }
        match &value.value {
            Value::Number(number) if !number.is_finite() => {
                return invalid_model("cache", "cached number is not finite");
            },
            Value::Text(text) => check_string(Some(text), "cached text")?,
            _ => {},
        }
    }
    for format in &chart.formats {
        if let Format::Marker { data } = format
            && data.len() > limits.max_record_bytes
        {
            return limit("record bytes", data.len(), limits.max_record_bytes);
        }
    }
    for label in &chart.labels {
        if !matches!(label.kind, DATA_LAB_EXT | DATA_LAB_EXT_CONTENTS | TEXT) {
            return invalid_model("label", "record kind is not a supported data-label kind");
        }
        if label.data.len() > limits.max_record_bytes {
            return limit("record bytes", label.data.len(), limits.max_record_bytes);
        }
    }

    let mut total = 0usize;
    for raw in &chart.unknown {
        if raw.data().len() > limits.max_record_bytes {
            return limit("record bytes", raw.data().len(), limits.max_record_bytes);
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

fn validate_link(link: &Link, context: Context, limits: Limits) -> Result<()> {
    match (context.kind(), link) {
        (Kind::Graph, Link::Graph { .. }) => {},
        (
            Kind::Excel,
            Link::Excel {
                source,
                formula,
                refs,
                ..
            },
        ) => {
            if formula.len() > limits.max_formula_bytes {
                return limit("formula bytes", formula.len(), limits.max_formula_bytes);
            }
            let maximum = limits.max_record_bytes.saturating_sub(8);
            if formula.len() > maximum {
                return limit("formula bytes", formula.len(), maximum);
            }
            if *source == Source::Automatic && !formula.is_empty() {
                return invalid_model("link", "automatic Excel BRAI has a nonempty formula");
            }
            for value in refs {
                if value.first_row > value.last_row || value.first_col > value.last_col {
                    return invalid_model("link", "cell range is reversed");
                }
                if let Some(count) = context.external_sheet_count()
                    && usize::from(value.external_sheet) >= count
                {
                    return invalid_model("link", "external-sheet index is out of range");
                }
            }
        },
        (Kind::Graph, Link::Excel { .. }) => {
            return invalid_model("link", "Excel BRAI cannot be encoded in a Graph chart");
        },
        (Kind::Excel, Link::Graph { .. }) => {
            return invalid_model("link", "Graph BRAI cannot be encoded in an Excel chart");
        },
    }
    Ok(())
}

pub(super) fn encode(chart: &Chart, limits: Limits) -> Result<Vec<u8>> {
    if !chart.authoring_proven {
        return Err(Error::UnsupportedAuthoring {
            reason: "complete CHARTSHEET, CHARTFOMATS, and SERIESDATA scaffolding is not modeled",
        });
    }
    validate(chart, limits)?;
    let mut out = Encoder::with_limits(limits)?;
    out.push(BOF, &bof(chart.context.kind()))?;

    let mut rect = [0u8; 16];
    rect.get_mut(0..4)
        .ok_or(Error::SizeOverflow {
            resource: "Chart rectangle",
        })?
        .copy_from_slice(&chart.rect.x.to_le_bytes());
    rect.get_mut(4..8)
        .ok_or(Error::SizeOverflow {
            resource: "Chart rectangle",
        })?
        .copy_from_slice(&chart.rect.y.to_le_bytes());
    rect.get_mut(8..12)
        .ok_or(Error::SizeOverflow {
            resource: "Chart rectangle",
        })?
        .copy_from_slice(&chart.rect.width.to_le_bytes());
    rect.get_mut(12..16)
        .ok_or(Error::SizeOverflow {
            resource: "Chart rectangle",
        })?
        .copy_from_slice(&chart.rect.height.to_le_bytes());
    out.push(CHART_REC, &rect)?;
    out.push(BEGIN, &[])?;
    out.push(SHT_PROPS, &chart.props.flags.to_le_bytes())?;

    for series in &chart.series {
        let mut data = [0u8; 12];
        put_u16(
            &mut data,
            0,
            match series.category_kind {
                DataKind::Numeric => 1,
                DataKind::Text => 3,
            },
        )?;
        put_u16(&mut data, 2, 1)?;
        put_u16(&mut data, 4, series.category_count.get())?;
        put_u16(&mut data, 6, series.value_count.get())?;
        put_u16(&mut data, 8, 1)?;
        put_u16(&mut data, 10, series.bubble_count.get())?;
        out.push(SERIES, &data)?;
        out.push(BEGIN, &[])?;
        for link in &series.links {
            let data = encode_link(link, chart.context, limits)?;
            out.push(BRAI, &data)?;
        }
        if let Some(name) = &series.name {
            out.push(SERIES_TEXT, &short_text(name)?)?;
        }
        out.push(SER_TO_CRT, &u16::from(series.group.get()).to_le_bytes())?;
        out.push(END, &[])?;
    }

    let axis_groups = if chart.groups.len() > 1 { 2u16 } else { 1u16 };
    out.push(AXES_USED, &axis_groups.to_le_bytes())?;
    out.push(AXIS_PARENT, &[0; 18])?;
    out.push(BEGIN, &[])?;
    for axis in &chart.axes {
        encode_axis(&mut out, axis)?;
    }
    for group in &chart.groups {
        encode_group(&mut out, group)?;
    }
    if chart.props.plot_area {
        out.push(PLOT_AREA, &[])?;
    }
    if let Some(legend) = chart.legend {
        let mut data = [0u8; 20];
        put_i32(&mut data, 0, legend.x)?;
        put_i32(&mut data, 4, legend.y)?;
        put_i32(&mut data, 8, legend.width)?;
        put_i32(&mut data, 12, legend.height)?;
        put_byte(&mut data, 16, legend.position)?;
        put_byte(&mut data, 17, legend.spacing)?;
        put_u16(&mut data, 18, legend.flags)?;
        out.push(LEGEND, &data)?;
    }
    if let Some(title) = &chart.title {
        out.push(SERIES_TEXT, &short_text(title)?)?;
    }
    for value in &chart.formats {
        encode_format(&mut out, value)?;
    }
    for label in &chart.labels {
        out.push(label.kind, &label.data)?;
    }

    let mut active_cache = None;
    for cache in &chart.caches {
        if chart.context.kind() == Kind::Excel && active_cache != Some(cache.index) {
            out.push(SI_INDEX, &cache.index.to_le_bytes())?;
            active_cache = Some(cache.index);
        }
        match &cache.value {
            Value::Number(number) => match cache.cell {
                Cell::Excel { row, col } if chart.context.kind() == Kind::Excel => {
                    let mut data = [0u8; 14];
                    put_u16(&mut data, 0, row)?;
                    put_u16(&mut data, 2, u16::from(col))?;
                    put_u16(&mut data, 4, cache.format)?;
                    put_f64(&mut data, 6, *number)?;
                    out.push(EXCEL_NUMBER, &data)?;
                },
                Cell::Graph { row, col } if chart.context.kind() == Kind::Graph => {
                    let mut data = [0u8; 15];
                    put_u16(&mut data, 0, row.get())?;
                    put_u16(&mut data, 2, col.get())?;
                    put_byte(&mut data, 4, 0)?;
                    put_u16(&mut data, 5, cache.format)?;
                    put_f64(&mut data, 7, *number)?;
                    out.push(GRAPH_NUMBER, &data)?;
                },
                _ => {
                    return invalid_model(
                        "cache",
                        "cell coordinate does not match producer context",
                    );
                },
            },
            Value::Text(text) => {
                let string = biff_string(text)?;
                let prefix = match chart.context.kind() {
                    Kind::Graph => 7usize,
                    Kind::Excel => 6usize,
                };
                let capacity = prefix
                    .checked_add(string.len())
                    .ok_or(Error::SizeOverflow {
                        resource: "cached chart string",
                    })?;
                let mut data = vec_with_capacity(capacity, "cached chart string")?;
                match cache.cell {
                    Cell::Excel { row, col } if chart.context.kind() == Kind::Excel => {
                        data.extend_from_slice(&row.to_le_bytes());
                        data.extend_from_slice(&u16::from(col).to_le_bytes());
                        data.extend_from_slice(&cache.format.to_le_bytes());
                    },
                    Cell::Graph { row, col } if chart.context.kind() == Kind::Graph => {
                        data.extend_from_slice(&row.get().to_le_bytes());
                        data.extend_from_slice(&col.get().to_le_bytes());
                        data.push(0);
                        data.extend_from_slice(&cache.format.to_le_bytes());
                    },
                    _ => {
                        return invalid_model(
                            "cache",
                            "cell coordinate does not match producer context",
                        );
                    },
                }
                data.extend_from_slice(&string);
                out.push(CELL_LABEL, &data)?;
            },
            Value::Blank => match cache.cell {
                Cell::Excel { row, col } if chart.context.kind() == Kind::Excel => {
                    let mut data = [0u8; 6];
                    put_u16(&mut data, 0, row)?;
                    put_u16(&mut data, 2, u16::from(col))?;
                    put_u16(&mut data, 4, cache.format)?;
                    out.push(EXCEL_BLANK, &data)?;
                },
                Cell::Graph { row, col } if chart.context.kind() == Kind::Graph => {
                    let mut data = [0u8; 7];
                    put_u16(&mut data, 0, row.get())?;
                    put_u16(&mut data, 2, col.get())?;
                    put_byte(&mut data, 4, 0)?;
                    put_u16(&mut data, 5, cache.format)?;
                    out.push(GRAPH_BLANK, &data)?;
                },
                _ => {
                    return invalid_model(
                        "cache",
                        "cell coordinate does not match producer context",
                    );
                },
            },
        }
    }
    out.push(END, &[])?;
    out.push(END, &[])?;
    out.push(EOF, &[])?;
    Ok(out.finish())
}

fn bof(kind: Kind) -> [u8; BOF_BYTES] {
    let mut data = [0u8; BOF_BYTES];
    match kind {
        Kind::Excel => {
            data[0..2].copy_from_slice(&EXCEL_VERSION.to_le_bytes());
            data[2..4].copy_from_slice(&EXCEL_DOC_TYPE.to_le_bytes());
            data[4..6].copy_from_slice(&0x0DBB_u16.to_le_bytes());
            data[6..8].copy_from_slice(&0x07CC_u16.to_le_bytes());
            data[8..12].copy_from_slice(&0u32.to_le_bytes());
            data[12..16].copy_from_slice(&6u32.to_le_bytes());
        },
        Kind::Graph => {
            data[0..2].copy_from_slice(&GRAPH_VERSION.to_le_bytes());
            data[2..4].copy_from_slice(&GRAPH_DOC_TYPE.to_le_bytes());
            data[4..6].copy_from_slice(&0x0DBB_u16.to_le_bytes());
            data[6..8].copy_from_slice(&0x07CD_u16.to_le_bytes());
            data[8..12].copy_from_slice(&(0x0000_0009_u32 | (6 << 14)).to_le_bytes());
            data[12..16].copy_from_slice(&(0x06_u32 | (6 << 8)).to_le_bytes());
        },
    }
    data
}

fn encode_link(link: &Link, context: Context, limits: Limits) -> Result<Vec<u8>> {
    validate_link(link, context, limits)?;
    match link {
        Link::Graph {
            role,
            source,
            unlinked_format,
            number_format,
            row_col,
        } => {
            let mut data = vec_with_capacity(8, "Graph BRAI")?;
            data.push(*role as u8);
            data.push(*source as u8);
            data.extend_from_slice(&(u16::from(*unlinked_format) | 2).to_le_bytes());
            data.extend_from_slice(&number_format.to_le_bytes());
            data.extend_from_slice(&row_col.get().to_le_bytes());
            Ok(data)
        },
        Link::Excel {
            role,
            source,
            unlinked_format,
            number_format,
            formula,
            ..
        } => {
            let capacity = 8usize
                .checked_add(formula.len())
                .ok_or(Error::SizeOverflow {
                    resource: "Excel BRAI",
                })?;
            let mut data = vec_with_capacity(capacity, "Excel BRAI")?;
            data.push(*role as u8);
            data.push(*source as u8);
            data.extend_from_slice(&u16::from(*unlinked_format).to_le_bytes());
            data.extend_from_slice(&number_format.to_le_bytes());
            let length = u16::try_from(formula.len()).map_err(|_| Error::InvalidModel {
                field: "link",
                reason: "formula length exceeds u16",
            })?;
            data.extend_from_slice(&length.to_le_bytes());
            data.extend_from_slice(formula);
            Ok(data)
        },
    }
}

fn encode_axis(out: &mut Encoder, axis: &Axis) -> Result<()> {
    let mut body = [0u8; 18];
    put_u16(
        &mut body,
        0,
        match axis.kind {
            axis::Kind::Category => 0,
            axis::Kind::Value => 1,
            axis::Kind::Series => 2,
        },
    )?;
    out.push(AXIS, &body)?;
    out.push(BEGIN, &[])?;
    if let Some(scale) = axis.scale {
        let mut data = [0u8; 42];
        put_f64(&mut data, 0, scale.min)?;
        put_f64(&mut data, 8, scale.max)?;
        put_f64(&mut data, 16, scale.major)?;
        put_f64(&mut data, 24, scale.minor)?;
        put_f64(&mut data, 32, scale.crossing)?;
        put_u16(&mut data, 40, scale.flags)?;
        out.push(VALUE_RANGE, &data)?;
    }
    if let Some(tick) = axis.tick {
        let mut data = [0u8; 26];
        put_byte(&mut data, 0, tick.major)?;
        put_byte(&mut data, 1, tick.minor)?;
        put_byte(&mut data, 2, tick.label)?;
        put_byte(&mut data, 3, tick.background)?;
        put_slice(&mut data, 4, &tick.color, "Tick color")?;
        put_u16(&mut data, 24, tick.flags)?;
        out.push(TICK, &data)?;
    }
    for line in &axis.lines {
        out.push(AXIS_LINE, &u16::from(line_kind(line.kind)).to_le_bytes())?;
        out.push(LINE_FORMAT, &line_bytes(line.format))?;
    }
    out.push(END, &[])
}

fn encode_group(out: &mut Encoder, group: &Group) -> Result<()> {
    let mut header = [0u8; 20];
    put_u16(&mut header, 16, u16::from(group.vary_colors))?;
    put_u16(&mut header, 18, u16::from(group.order.get()))?;
    out.push(CHART_FORMAT, &header)?;
    out.push(BEGIN, &[])?;
    match group.family {
        Family::Line { flags } => out.push(LINE, &flags.to_le_bytes())?,
        Family::Area { flags } => out.push(AREA, &flags.to_le_bytes())?,
        Family::Bar {
            overlap,
            gap,
            flags,
        } => {
            let mut data = [0u8; 6];
            put_i16(&mut data, 0, overlap.get())?;
            put_u16(&mut data, 2, gap.get())?;
            put_u16(&mut data, 4, flags)?;
            out.push(BAR, &data)?;
        },
        Family::Pie {
            rotation,
            hole,
            flags,
        } => {
            let mut data = [0u8; 6];
            put_u16(&mut data, 0, rotation)?;
            put_u16(&mut data, 2, hole)?;
            put_u16(&mut data, 4, flags)?;
            out.push(PIE, &data)?;
        },
        Family::Scatter {
            bubble_percent,
            bubble_kind,
            flags,
        } => {
            let mut data = [0u8; 6];
            put_u16(&mut data, 0, bubble_percent.get())?;
            put_u16(&mut data, 2, bubble_kind as u16)?;
            put_u16(&mut data, 4, flags)?;
            out.push(SCATTER, &data)?;
        },
        Family::Radar { filled, flags } => {
            out.push(
                if filled { RADAR_AREA } else { RADAR },
                &flags.to_le_bytes(),
            )?;
        },
        Family::Surface { flags } => out.push(SURFACE, &flags.to_le_bytes())?,
    }
    for line in &group.lines {
        let marker = crate::record::line::Line::new(line.kind);
        out.push(CRT_LINE, &marker.payload())?;
        out.push(LINE_FORMAT, &line_bytes(line.format))?;
    }
    for drop in &group.drop_bars {
        out.push(DROP_BAR, &drop.gap.get().to_le_bytes())?;
        out.push(BEGIN, &[])?;
        out.push(LINE_FORMAT, &line_bytes(drop.line))?;
        out.push(AREA_FORMAT, &area_bytes(drop.area))?;
        out.push(END, &[])?;
    }
    out.push(END, &[])
}

fn encode_format(out: &mut Encoder, value: &Format) -> Result<()> {
    match value {
        Format::Line(value) => out.push(LINE_FORMAT, &line_bytes(*value)),
        Format::Area(value) => out.push(AREA_FORMAT, &area_bytes(*value)),
        Format::Marker { data } => out.push(MARKER_FORMAT, data),
        Format::Data {
            point,
            series,
            flags,
        } => {
            let mut data = [0u8; 8];
            put_u16(&mut data, 0, *point)?;
            put_u16(&mut data, 2, *series)?;
            put_u16(&mut data, 4, 0)?;
            put_u16(&mut data, 6, *flags)?;
            out.push(DATA_FORMAT, &data)
        },
        Format::Pie { explosion } => out.push(PIE_FORMAT, &explosion.to_le_bytes()),
    }
}

fn line_bytes(value: format::Line) -> [u8; 12] {
    let mut data = [0u8; 12];
    data[0..4].copy_from_slice(&value.color);
    data[4..6].copy_from_slice(&value.pattern.to_le_bytes());
    data[6..8].copy_from_slice(&value.weight.to_le_bytes());
    data[8..10].copy_from_slice(&value.flags.to_le_bytes());
    data[10..12].copy_from_slice(&value.color_index.to_le_bytes());
    data
}

fn area_bytes(value: format::Area) -> [u8; 16] {
    let mut data = [0u8; 16];
    data[0..4].copy_from_slice(&value.foreground);
    data[4..8].copy_from_slice(&value.background);
    data[8..10].copy_from_slice(&value.pattern.to_le_bytes());
    data[10..12].copy_from_slice(&value.flags.to_le_bytes());
    data[12..14].copy_from_slice(&value.foreground_index.to_le_bytes());
    data[14..16].copy_from_slice(&value.background_index.to_le_bytes());
    data
}

fn parse_short_text(record: RecordRef<'_>) -> Result<String> {
    let data = record.payload();
    if data.len() < 4 || u16_at(data, 0, record)? != 0 {
        return invalid(
            record,
            "SeriesText is truncated or its reserved field is nonzero",
        );
    }
    let string = data.get(2..).ok_or(Error::InvalidChart {
        offset: record.offset(),
        reason: "SeriesText string is truncated",
    })?;
    parse_string(string, record)
}

fn short_text(value: &str) -> Result<Vec<u8>> {
    let string = biff_string(value)?;
    let capacity = 2usize
        .checked_add(string.len())
        .ok_or(Error::SizeOverflow {
            resource: "chart text",
        })?;
    let mut data = vec_with_capacity(capacity, "chart text")?;
    data.extend_from_slice(&0u16.to_le_bytes());
    data.extend_from_slice(&string);
    Ok(data)
}

fn parse_string(data: &[u8], record: RecordRef<'_>) -> Result<String> {
    if data.len() < 2 {
        return invalid(record, "chart string is shorter than two bytes");
    }
    let count = usize::from(byte_at(data, 0, record)?);
    let flags = byte_at(data, 1, record)?;
    if flags & !1 != 0 {
        return invalid(record, "chart string uses unsupported option flags");
    }
    let wide = flags & 1 != 0;
    let width = if wide { 2usize } else { 1usize };
    let content = count.checked_mul(width).ok_or(Error::SizeOverflow {
        resource: "chart string",
    })?;
    let expected = 2usize.checked_add(content).ok_or(Error::SizeOverflow {
        resource: "chart string",
    })?;
    if data.len() != expected {
        return invalid(
            record,
            "chart string length does not match its character count",
        );
    }
    let bytes = data.get(2..).ok_or(Error::InvalidChart {
        offset: record.offset(),
        reason: "chart string content is truncated",
    })?;
    let reserve = count
        .checked_mul(if wide { 3 } else { 2 })
        .ok_or(Error::SizeOverflow {
            resource: "chart string",
        })?;
    let mut output = String::new();
    output
        .try_reserve_exact(reserve)
        .map_err(|_| Error::Allocation {
            resource: "chart string",
        })?;
    if wide {
        let units = bytes
            .chunks_exact(2)
            .map(|value| u16::from_le_bytes([value[0], value[1]]));
        for value in char::decode_utf16(units) {
            output.push(value.map_err(|_| Error::InvalidChart {
                offset: record.offset(),
                reason: "chart string contains invalid UTF-16",
            })?);
        }
    } else {
        for value in bytes {
            output.push(char::from(*value));
        }
    }
    Ok(output)
}

fn biff_string(value: &str) -> Result<Vec<u8>> {
    let count = value.encode_utf16().count();
    if count > usize::from(u8::MAX) {
        return invalid_model("text", "chart string exceeds 255 UTF-16 code units");
    }
    let wide = value.encode_utf16().any(|unit| unit > u16::from(u8::MAX));
    let width = if wide { 2usize } else { 1usize };
    let capacity = 2usize
        .checked_add(count.checked_mul(width).ok_or(Error::SizeOverflow {
            resource: "chart string",
        })?)
        .ok_or(Error::SizeOverflow {
            resource: "chart string",
        })?;
    let mut data = vec_with_capacity(capacity, "chart string")?;
    data.push(u8::try_from(count).map_err(|_| Error::InvalidModel {
        field: "text",
        reason: "chart string exceeds 255 UTF-16 code units",
    })?);
    data.push(u8::from(wide));
    for unit in value.encode_utf16() {
        if wide {
            data.extend_from_slice(&unit.to_le_bytes());
        } else {
            data.push(u8::try_from(unit).map_err(|_| Error::InvalidModel {
                field: "text",
                reason: "narrow chart string contains a wide code unit",
            })?);
        }
    }
    Ok(data)
}

fn parse_role(value: u8, record: RecordRef<'_>) -> Result<Role> {
    match value {
        0 => Ok(Role::Name),
        1 => Ok(Role::Values),
        2 => Ok(Role::Categories),
        3 => Ok(Role::Bubbles),
        _ => invalid(record, "BRAI role is outside the defined range"),
    }
}

fn parse_source(value: u8, record: RecordRef<'_>) -> Result<Source> {
    match value {
        0 => Ok(Source::Automatic),
        1 => Ok(Source::Literal),
        2 => Ok(Source::Cells),
        _ => invalid(record, "BRAI source is outside the defined range"),
    }
}

fn blank_kind(kind: Kind) -> RecordKind {
    match kind {
        Kind::Graph => GRAPH_BLANK,
        Kind::Excel => EXCEL_BLANK,
    }
}

fn number_kind(kind: Kind) -> RecordKind {
    match kind {
        Kind::Graph => GRAPH_NUMBER,
        Kind::Excel => EXCEL_NUMBER,
    }
}

fn cache_cell(data: &[u8], context: Context, record: RecordRef<'_>) -> Result<Cell> {
    let row = u16_at(data, 0, record)?;
    let col = u16_at(data, 2, record)?;
    match context.kind() {
        Kind::Graph => Ok(Cell::Graph {
            row: RowCol::new(row).ok_or(Error::InvalidChart {
                offset: record.offset(),
                reason: "Graph cache row exceeds 3,999",
            })?,
            col: RowCol::new(col).ok_or(Error::InvalidChart {
                offset: record.offset(),
                reason: "Graph cache column exceeds 3,999",
            })?,
        }),
        Kind::Excel => Ok(Cell::Excel {
            row,
            col: u8::try_from(col).map_err(|_| Error::InvalidChart {
                offset: record.offset(),
                reason: "Excel cache column exceeds the BIFF8 grid",
            })?,
        }),
    }
}

fn cache_matches(cell: Cell, kind: Kind) -> bool {
    matches!(
        (cell, kind),
        (Cell::Graph { .. }, Kind::Graph) | (Cell::Excel { .. }, Kind::Excel)
    )
}

fn check_string(value: Option<&str>, field: &'static str) -> Result<()> {
    let Some(value) = value else {
        return Ok(());
    };
    if value.encode_utf16().count() > usize::from(u8::MAX) {
        return invalid_model(field, "chart string exceeds 255 UTF-16 code units");
    }
    Ok(())
}

fn valid_props(flags: u32) -> bool {
    let reserved_clear = flags & 0x0000_FFE4 == 0 && flags & 0xFF00_0000 == 0;
    let blank = (flags >> 16) & 0xFF;
    let auto_plot = flags & (1 << 4) != 0;
    let manual_plot = flags & (1 << 3) != 0;
    reserved_clear && blank <= 2 && (!auto_plot || manual_plot)
}

fn line_kind(kind: axis::LineKind) -> u8 {
    match kind {
        axis::LineKind::Axis => 0,
        axis::LineKind::MajorGrid => 1,
        axis::LineKind::MinorGrid => 2,
        axis::LineKind::Wall => 3,
    }
}

fn count_at(data: &[u8], offset: usize, record: RecordRef<'_>) -> Result<Count> {
    Count::new(u16_at(data, offset, record)?).ok_or(Error::InvalidChart {
        offset: record.offset(),
        reason: "series point count exceeds 32,767",
    })
}

fn byte_at(data: &[u8], offset: usize, record: RecordRef<'_>) -> Result<u8> {
    data.get(offset).copied().ok_or(Error::InvalidChart {
        offset: record.offset(),
        reason: "record scalar is truncated",
    })
}

fn u16_at(data: &[u8], offset: usize, record: RecordRef<'_>) -> Result<u16> {
    let bytes = data
        .get(
            offset..offset.checked_add(2).ok_or(Error::SizeOverflow {
                resource: "record scalar",
            })?,
        )
        .ok_or(Error::InvalidChart {
            offset: record.offset(),
            reason: "record u16 is truncated",
        })?;
    let lo = bytes.first().copied().ok_or(Error::InvalidChart {
        offset: record.offset(),
        reason: "record u16 is truncated",
    })?;
    let hi = bytes.get(1).copied().ok_or(Error::InvalidChart {
        offset: record.offset(),
        reason: "record u16 is truncated",
    })?;
    Ok(u16::from_le_bytes([lo, hi]))
}

fn i16_at(data: &[u8], offset: usize, record: RecordRef<'_>) -> Result<i16> {
    Ok(i16::from_le_bytes(
        u16_at(data, offset, record)?.to_le_bytes(),
    ))
}

fn u32_at(data: &[u8], offset: usize, record: RecordRef<'_>) -> Result<u32> {
    let bytes = data
        .get(
            offset..offset.checked_add(4).ok_or(Error::SizeOverflow {
                resource: "record scalar",
            })?,
        )
        .ok_or(Error::InvalidChart {
            offset: record.offset(),
            reason: "record u32 is truncated",
        })?;
    let array = <[u8; 4]>::try_from(bytes).map_err(|_| Error::InvalidChart {
        offset: record.offset(),
        reason: "record u32 is truncated",
    })?;
    Ok(u32::from_le_bytes(array))
}

fn i32_at(data: &[u8], offset: usize, record: RecordRef<'_>) -> Result<i32> {
    Ok(i32::from_le_bytes(
        u32_at(data, offset, record)?.to_le_bytes(),
    ))
}

fn f64_at(data: &[u8], offset: usize, record: RecordRef<'_>) -> Result<f64> {
    let bytes = data
        .get(
            offset..offset.checked_add(8).ok_or(Error::SizeOverflow {
                resource: "record scalar",
            })?,
        )
        .ok_or(Error::InvalidChart {
            offset: record.offset(),
            reason: "record f64 is truncated",
        })?;
    let array = <[u8; 8]>::try_from(bytes).map_err(|_| Error::InvalidChart {
        offset: record.offset(),
        reason: "record f64 is truncated",
    })?;
    Ok(f64::from_le_bytes(array))
}

fn array4_at(data: &[u8], offset: usize, record: RecordRef<'_>) -> Result<[u8; 4]> {
    let bytes = data
        .get(
            offset..offset.checked_add(4).ok_or(Error::SizeOverflow {
                resource: "record array",
            })?,
        )
        .ok_or(Error::InvalidChart {
            offset: record.offset(),
            reason: "record array is truncated",
        })?;
    <[u8; 4]>::try_from(bytes).map_err(|_| Error::InvalidChart {
        offset: record.offset(),
        reason: "record array is truncated",
    })
}

fn exact(record: RecordRef<'_>, expected: usize) -> Result<()> {
    if record.payload().len() != expected {
        return invalid(record, "record payload has an invalid fixed length");
    }
    Ok(())
}

fn invalid<T>(record: RecordRef<'_>, reason: &'static str) -> Result<T> {
    Err(Error::InvalidChart {
        offset: record.offset(),
        reason,
    })
}

fn invalid_model<T>(field: &'static str, reason: &'static str) -> Result<T> {
    Err(Error::InvalidModel { field, reason })
}

fn limit<T>(resource: &'static str, observed: usize, maximum: usize) -> Result<T> {
    Err(Error::LimitExceeded {
        resource,
        observed: as_u64(observed),
        maximum: as_u64(maximum),
    })
}

fn check_add(current: usize, maximum: usize, resource: &'static str) -> Result<()> {
    let observed = current
        .checked_add(1)
        .ok_or(Error::SizeOverflow { resource })?;
    if observed > maximum {
        return limit(resource, observed, maximum);
    }
    Ok(())
}

fn push<T>(values: &mut Vec<T>, value: T, resource: &'static str) -> Result<()> {
    values
        .try_reserve(1)
        .map_err(|_| Error::Allocation { resource })?;
    values.push(value);
    Ok(())
}

fn copy(data: &[u8], resource: &'static str, maximum: usize) -> Result<Vec<u8>> {
    if data.len() > maximum {
        return limit(resource, data.len(), maximum);
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(data.len())
        .map_err(|_| Error::Allocation { resource })?;
    output.extend_from_slice(data);
    Ok(output)
}

fn vec_with_capacity(capacity: usize, resource: &'static str) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|_| Error::Allocation { resource })?;
    Ok(output)
}

fn put_slice(output: &mut [u8], offset: usize, value: &[u8], resource: &'static str) -> Result<()> {
    let end = offset
        .checked_add(value.len())
        .ok_or(Error::SizeOverflow { resource })?;
    output
        .get_mut(offset..end)
        .ok_or(Error::SizeOverflow { resource })?
        .copy_from_slice(value);
    Ok(())
}

fn put_byte(output: &mut [u8], offset: usize, value: u8) -> Result<()> {
    *output.get_mut(offset).ok_or(Error::SizeOverflow {
        resource: "encoded scalar",
    })? = value;
    Ok(())
}

fn put_u16(output: &mut [u8], offset: usize, value: u16) -> Result<()> {
    put_slice(output, offset, &value.to_le_bytes(), "encoded u16")
}

fn put_i16(output: &mut [u8], offset: usize, value: i16) -> Result<()> {
    put_slice(output, offset, &value.to_le_bytes(), "encoded i16")
}

fn put_i32(output: &mut [u8], offset: usize, value: i32) -> Result<()> {
    put_slice(output, offset, &value.to_le_bytes(), "encoded i32")
}

fn put_f64(output: &mut [u8], offset: usize, value: f64) -> Result<()> {
    put_slice(output, offset, &value.to_le_bytes(), "encoded f64")
}
