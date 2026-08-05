//! Chart BIFF record traversal and encoding.

use std::char;

use super::super::axis::{self, Axis};
use super::super::cache;
use super::super::format::{self, Format};
use super::super::group;
use super::super::layout;
use super::super::model::{
    Binding, Cache, CellRef, Chart, Context, Count, DataKind, Family, Group, GroupId, Label,
    Legend, Link, Order, Origin, Owner, Props, Raw, Rect, Role, RowCol, Series, Source, Value,
    XlValue, cache_dimensions, dimensions_cover,
};
use super::super::{
    BOF, BOF_BYTES, EOF, EXCEL_DOC_TYPE, EXCEL_VERSION, GRAPH_DOC_TYPE, GRAPH_VERSION,
    Kind as ChartKind, Ref,
};
use crate::limits::as_u64;
use crate::{Error, Limits, Result};
use litchi_biff::{Encoder, Kind as RecordKind, RecordRef};

const CONTINUE: RecordKind = RecordKind::from_wire(0x003C);
const SCL: RecordKind = RecordKind::from_wire(0x00A0);
const DIMENSIONS: RecordKind = RecordKind::from_wire(0x0200);
const GRAPH_BLANK: RecordKind = RecordKind::from_wire(0x0001);
const GRAPH_NUMBER: RecordKind = RecordKind::from_wire(0x0003);
const EXCEL_BLANK: RecordKind = RecordKind::from_wire(0x0201);
const EXCEL_NUMBER: RecordKind = RecordKind::from_wire(0x0203);
const EXCEL_BOOL_ERR: RecordKind = RecordKind::from_wire(0x0205);
const CELL_LABEL: RecordKind = RecordKind::from_wire(0x0204);
const DATA_LAB_EXT: RecordKind = RecordKind::from_wire(0x086A);
const DATA_LAB_EXT_CONTENTS: RecordKind = RecordKind::from_wire(0x086B);
const CHART_REC: RecordKind = RecordKind::from_wire(0x1002);
const SERIES: RecordKind = RecordKind::from_wire(0x1003);
const DATA_FORMAT: RecordKind = RecordKind::from_wire(0x1006);
const LINE_FORMAT: RecordKind = RecordKind::from_wire(0x1007);
const MARKER_FORMAT: RecordKind = RecordKind::from_wire(0x1009);
const AREA_FORMAT: RecordKind = RecordKind::from_wire(0x100A);
const PIE_FORMAT: RecordKind = RecordKind::from_wire(0x100B);
const SERIES_TEXT: RecordKind = RecordKind::from_wire(0x100D);
const CHART_FORMAT: RecordKind = RecordKind::from_wire(0x1014);
const LEGEND: RecordKind = RecordKind::from_wire(0x1015);
const SERIES_LIST: RecordKind = RecordKind::from_wire(0x1016);
const BAR: RecordKind = RecordKind::from_wire(0x1017);
const LINE: RecordKind = RecordKind::from_wire(0x1018);
const PIE: RecordKind = RecordKind::from_wire(0x1019);
const AREA: RecordKind = RecordKind::from_wire(0x101A);
const SCATTER: RecordKind = RecordKind::from_wire(0x101B);
const CRT_LINE: RecordKind = RecordKind::from_wire(0x101C);
const AXIS: RecordKind = RecordKind::from_wire(0x101D);
const TICK: RecordKind = RecordKind::from_wire(0x101E);
const VALUE_RANGE: RecordKind = RecordKind::from_wire(0x101F);
const CAT_SER_RANGE: RecordKind = RecordKind::from_wire(0x1020);
const AXIS_LINE: RecordKind = RecordKind::from_wire(0x1021);
const CRT_LINK: RecordKind = RecordKind::from_wire(0x1022);
const DEFAULT_TEXT: RecordKind = RecordKind::from_wire(0x1024);
const TEXT: RecordKind = RecordKind::from_wire(0x1025);
const FONT_X: RecordKind = RecordKind::from_wire(0x1026);
const OBJECT_LINK: RecordKind = RecordKind::from_wire(0x1027);
const FRAME: RecordKind = RecordKind::from_wire(0x1032);
const BEGIN: RecordKind = RecordKind::from_wire(0x1033);
const END: RecordKind = RecordKind::from_wire(0x1034);
const PLOT_AREA: RecordKind = RecordKind::from_wire(0x1035);
const DROP_BAR: RecordKind = RecordKind::from_wire(0x103D);
const RADAR: RecordKind = RecordKind::from_wire(0x103E);
const SURFACE: RecordKind = RecordKind::from_wire(0x103F);
const RADAR_AREA: RecordKind = RecordKind::from_wire(0x1040);
const AXIS_PARENT: RecordKind = RecordKind::from_wire(0x1041);
const SHT_PROPS: RecordKind = RecordKind::from_wire(0x1044);
const SER_TO_CRT: RecordKind = RecordKind::from_wire(0x1045);
const AXES_USED: RecordKind = RecordKind::from_wire(0x1046);
const SER_PARENT: RecordKind = RecordKind::from_wire(0x104A);
const SER_AUX_TREND: RecordKind = RecordKind::from_wire(0x104B);
const POS: RecordKind = RecordKind::from_wire(0x104F);
const BRAI: RecordKind = RecordKind::from_wire(0x1051);
const SER_AUX_ERR_BAR: RecordKind = RecordKind::from_wire(0x105B);
const PLOT_GROWTH: RecordKind = RecordKind::from_wire(0x1064);
const SI_INDEX: RecordKind = RecordKind::from_wire(0x1065);

use super::model::{PendingDrop, PendingLine};

pub(crate) fn parse(input: Ref<'_>, context: Context, limits: Limits) -> Result<Chart> {
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
        zoom: layout::Zoom::default(),
        growth: layout::Growth::default(),
        title: None,
        series: Vec::new(),
        groups: Vec::new(),
        axes: Vec::new(),
        parents: Vec::new(),
        legend: None,
        caches: Vec::new(),
        dimensions: cache::Dims::empty(context.kind()),
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
    let mut ai_next = 0usize;
    let mut last_ai = None;
    let mut series_owner_seen = false;
    let mut series_parent = None;
    let mut current_axis = None;
    let mut axis_depth = None;
    let mut group_depth = None;
    let mut group_family_seen = false;
    let mut group_link_seen = false;
    let mut parent_depth = None;
    let mut current_parent: Option<usize> = None;
    let mut parent_needs_pos = false;
    let mut parent_groups = 0usize;
    let mut parent_group_started = false;
    let mut pending_begin = false;
    let strict_excel = context.kind() == ChartKind::Excel;
    let mut cache_section = None;
    let mut next_cache_section = 0usize;
    let mut zoom_seen = false;
    let mut growth_seen = false;
    let mut dimensions_seen = false;
    let mut axes_used = None;
    let mut chart_seen = false;
    let mut props_seen = false;
    let mut chart_closed = false;
    let mut pending_axis_line = None;
    let mut pending_drop: Option<PendingDrop> = None;
    let mut unknown_bytes = 0usize;
    let mut first = true;

    for item in input.records() {
        let record = item?;
        if record.payload().len() > limits.biff.max_record_bytes {
            return limit(
                "record bytes",
                record.payload().len(),
                limits.biff.max_record_bytes,
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
        if last_ai.is_some() && record.kind() != SERIES_TEXT {
            last_ai = None;
        }
        if series_parent.is_some() && !matches!(record.kind(), SER_AUX_TREND | SER_AUX_ERR_BAR) {
            return invalid(
                record,
                "SerParent is not followed immediately by SerAuxTrend or SerAuxErrBar",
            );
        }
        if series_depth == Some(depth) && ai_next < Role::ALL.len() {
            let optional_text = record.kind() == SERIES_TEXT && last_ai.is_some();
            if record.kind() != BRAI && !optional_text {
                return invalid(
                    record,
                    "Series must begin with exactly four ordered AI bindings",
                );
            }
        }
        if strict_excel && parent_needs_pos && record.kind() != POS {
            return invalid(
                record,
                "AxisParent Begin is not followed immediately by Pos",
            );
        }
        if strict_excel && group_depth == Some(depth) {
            let family = matches!(
                record.kind(),
                BAR | LINE | PIE | AREA | SCATTER | RADAR | RADAR_AREA | SURFACE
            );
            if !group_family_seen && !family {
                return invalid(
                    record,
                    "ChartFormat Begin is not followed by a chart family",
                );
            }
            if group_family_seen && !group_link_seen && record.kind() != CRT_LINK {
                return invalid(
                    record,
                    "chart family is not followed immediately by CrtLink",
                );
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
                if strict_excel && depth == 0 && !pending_begin {
                    return invalid(record, "Begin record has no chart-level collection owner");
                }
                pending_begin = false;
                depth = depth.checked_add(1).ok_or(Error::SizeOverflow {
                    resource: "chart nesting",
                })?;
                if depth > limits.max_nesting {
                    return limit("chart nesting", depth, limits.max_nesting);
                }
                if parent_depth == Some(depth) {
                    parent_needs_pos = true;
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
                    if ai_next != Role::ALL.len() || !series_owner_seen || series_parent.is_some() {
                        return invalid(
                            record,
                            "Series requires exactly four AI bindings and one SerToCrt",
                        );
                    }
                    current_series = None;
                    series_depth = None;
                    ai_next = 0;
                    last_ai = None;
                    series_owner_seen = false;
                    series_parent = None;
                }
                if axis_depth == Some(depth) {
                    current_axis = None;
                    axis_depth = None;
                }
                if group_depth == Some(depth) {
                    if !group_family_seen || (strict_excel && !group_link_seen) {
                        return invalid(
                            record,
                            "ChartFormat collection is missing its family or required Excel CrtLink",
                        );
                    }
                    group_depth = None;
                    group_family_seen = false;
                    group_link_seen = false;
                    parent_groups = parent_groups.checked_add(1).ok_or(Error::SizeOverflow {
                        resource: "axis-parent chart groups",
                    })?;
                }
                if parent_depth == Some(depth) {
                    if strict_excel && (parent_needs_pos || !(1..=4).contains(&parent_groups)) {
                        return invalid(
                            record,
                            "AxisParent requires Pos and one through four ChartFormat collections",
                        );
                    }
                    parent_depth = None;
                    current_parent = None;
                    parent_groups = 0;
                    parent_group_started = false;
                }
                if depth == 1 {
                    let parent_count =
                        u16::try_from(chart.parents.len()).map_err(|_| Error::SizeOverflow {
                            resource: "axis-parent count",
                        })?;
                    if strict_excel
                        && (!zoom_seen
                            || !growth_seen
                            || !props_seen
                            || axes_used != Some(parent_count))
                    {
                        return invalid(
                            record,
                            "CHARTFOMATS is missing or misorders a mandatory collection",
                        );
                    }
                    chart_closed = true;
                }
                depth -= 1;
            },
            CHART_REC => {
                if chart_seen || depth != 0 {
                    return invalid(record, "Chart must occur once at chart-substream level");
                }
                exact(record, 16)?;
                chart.rect = Rect {
                    x: i32_at(data, 0, record)?,
                    y: i32_at(data, 4, record)?,
                    width: i32_at(data, 8, record)?,
                    height: i32_at(data, 12, record)?,
                };
                chart_seen = true;
                pending_begin = true;
            },
            SCL => {
                if zoom_seen || (strict_excel && (depth != 1 || growth_seen || !chart_seen)) {
                    return invalid(
                        record,
                        "Scl must occur once before PlotGrowth in CHARTFOMATS",
                    );
                }
                exact(record, 4)?;
                chart.zoom = layout::Zoom::new(u16_at(data, 0, record)?, u16_at(data, 2, record)?)
                    .ok_or(Error::InvalidChart {
                        offset: record.offset(),
                        reason: "Scl fraction is outside 1/10 through 4",
                    })?;
                zoom_seen = true;
            },
            PLOT_GROWTH => {
                if growth_seen || (strict_excel && (depth != 1 || !zoom_seen)) {
                    return invalid(
                        record,
                        "PlotGrowth must occur once after Scl in CHARTFOMATS",
                    );
                }
                exact(record, 8)?;
                chart.growth = layout::Growth {
                    x: layout::Fixed::from_raw(i32_at(data, 0, record)?),
                    y: layout::Fixed::from_raw(i32_at(data, 4, record)?),
                };
                growth_seen = true;
            },
            SHT_PROPS => {
                if props_seen
                    || (strict_excel && (depth != 1 || !growth_seen || axes_used.is_some()))
                {
                    return invalid(record, "ShtProps is duplicated or misplaced in CHARTFOMATS");
                }
                exact(record, 4)?;
                let flags = u32_at(data, 0, record)?;
                if !valid_props(flags) {
                    return invalid(record, "ShtProps reserved bits or blank mode are invalid");
                }
                chart.props.flags = flags;
                props_seen = true;
            },
            SERIES => {
                if current_series.is_some()
                    || (strict_excel && (depth != 1 || !growth_seen || props_seen))
                {
                    return invalid(
                        record,
                        "Series must own a non-overlapping CHARTFOMATS collection",
                    );
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
                let mut series = Series::new(context);
                series.category_kind = category_kind;
                series.category_count = count_at(data, 4, record)?;
                series.value_count = count_at(data, 6, record)?;
                series.bubble_count = count_at(data, 10, record)?;
                push(&mut chart.series, series, "chart series")?;
                current_series = chart.series.len().checked_sub(1);
                series_depth = depth.checked_add(1);
                ai_next = 0;
                last_ai = None;
                series_owner_seen = false;
                series_parent = None;
                pending_begin = true;
            },
            BRAI => {
                if series_depth != Some(depth) {
                    return invalid(record, "BRAI appears outside a Series collection");
                }
                let link = parse_link(record, context, limits)?;
                let expected = Role::ALL.get(ai_next).copied().ok_or(Error::InvalidChart {
                    offset: record.offset(),
                    reason: "Series collection contains more than four AI bindings",
                })?;
                if link.role() != expected {
                    return invalid(
                        record,
                        "Series AI bindings are missing or out of role order",
                    );
                }
                let index = current_series.ok_or(Error::InvalidChart {
                    offset: record.offset(),
                    reason: "BRAI appears outside a Series collection",
                })?;
                let series = chart.series.get_mut(index).ok_or(Error::InvalidChart {
                    offset: record.offset(),
                    reason: "BRAI refers to a missing Series",
                })?;
                series.ai.replace(Binding::new(link, None));
                ai_next = ai_next.checked_add(1).ok_or(Error::SizeOverflow {
                    resource: "series AI count",
                })?;
                last_ai = Some(expected);
            },
            SER_TO_CRT => {
                exact(record, 2)?;
                if series_depth != Some(depth) {
                    return invalid(record, "SerToCrt appears outside a Series collection");
                }
                if series_owner_seen || series_parent.is_some() {
                    return invalid(record, "Series contains more than one owner branch");
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
                    .owner = Owner::Group(group);
                series_owner_seen = true;
            },
            SER_PARENT => {
                if series_depth != Some(depth) || series_owner_seen || series_parent.is_some() {
                    return invalid(
                        record,
                        "SerParent is duplicated or outside a Series owner branch",
                    );
                }
                series_parent =
                    Some(crate::record::series::Parent::parse(record).map_err(|_| {
                        Error::InvalidChart {
                            offset: record.offset(),
                            reason: "SerParent series index is outside 1 through 254",
                        }
                    })?);
            },
            SER_AUX_TREND | SER_AUX_ERR_BAR => {
                if series_depth != Some(depth) || series_owner_seen {
                    return invalid(
                        record,
                        "auxiliary-series record is duplicated or outside Series",
                    );
                }
                let parent = series_parent.take().ok_or(Error::InvalidChart {
                    offset: record.offset(),
                    reason: "auxiliary-series record must immediately follow SerParent",
                })?;
                let owner = if record.kind() == SER_AUX_TREND {
                    exact(record, 28)?;
                    Owner::Trend {
                        parent,
                        data: data.try_into().map_err(|_| Error::InvalidChart {
                            offset: record.offset(),
                            reason: "SerAuxTrend payload is not 28 bytes",
                        })?,
                    }
                } else {
                    exact(record, 14)?;
                    Owner::ErrorBar {
                        parent,
                        data: data.try_into().map_err(|_| Error::InvalidChart {
                            offset: record.offset(),
                            reason: "SerAuxErrBar payload is not 14 bytes",
                        })?,
                    }
                };
                let series = current_series
                    .and_then(|index| chart.series.get_mut(index))
                    .ok_or(Error::InvalidChart {
                        offset: record.offset(),
                        reason: "auxiliary owner refers to a missing Series",
                    })?;
                series.owner = owner;
                series_owner_seen = true;
            },
            SERIES_TEXT => {
                if series_depth == Some(depth) {
                    let role = last_ai.take().ok_or(Error::InvalidChart {
                        offset: record.offset(),
                        reason: "SeriesText in Series must immediately follow one BRAI",
                    })?;
                    let text = parse_short_text(record)?;
                    let series = current_series
                        .and_then(|index| chart.series.get_mut(index))
                        .ok_or(Error::InvalidChart {
                            offset: record.offset(),
                            reason: "SeriesText refers to a missing Series",
                        })?;
                    series
                        .ai
                        .get_mut(role)
                        .set_text(text)
                        .map_err(|_| Error::InvalidChart {
                            offset: record.offset(),
                            reason: "one AI has more than one SeriesText",
                        })?;
                } else {
                    add_raw(&mut chart, record, &mut unknown_bytes, limits)?;
                }
            },
            CHART_FORMAT => {
                if group_depth.is_some() || parent_depth != Some(depth) {
                    return invalid(
                        record,
                        "ChartFormat must belong to one AxisParent collection",
                    );
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
                let order = Order::new(raw).ok_or(Error::InvalidChart {
                    offset: record.offset(),
                    reason: "chart-group order exceeds nine",
                })?;
                let parent = current_parent
                    .and_then(|index| chart.parents.get(index))
                    .map(|parent| parent.id())
                    .ok_or(Error::InvalidChart {
                        offset: record.offset(),
                        reason: "ChartFormat has no AxisParent owner",
                    })?;
                push(
                    &mut chart.groups,
                    Group {
                        parent,
                        order,
                        vary_colors: vary & 1 != 0,
                        family: Family::Line { flags: 0 },
                        link: crate::record::line::Link::new([0; 10]),
                        lines: Vec::new(),
                        drop_bars: Vec::new(),
                    },
                    "chart groups",
                )?;
                group_depth = depth.checked_add(1);
                group_family_seen = false;
                group_link_seen = false;
                parent_group_started = true;
                pending_begin = true;
            },
            BAR | LINE | PIE | AREA | SCATTER | RADAR | RADAR_AREA | SURFACE => {
                if group_depth != Some(depth) {
                    return invalid(record, "chart-family record appears outside ChartFormat");
                }
                if group_family_seen {
                    return invalid(record, "ChartFormat contains more than one chart family");
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
                group_family_seen = true;
            },
            CRT_LINK => {
                if group_depth != Some(depth) || !group_family_seen || group_link_seen {
                    return invalid(
                        record,
                        "CrtLink is missing, duplicated, or outside ChartFormat",
                    );
                }
                let link = crate::record::line::Link::from_payload(data).map_err(|_| {
                    Error::InvalidChart {
                        offset: record.offset(),
                        reason: "CrtLink payload is not ten bytes",
                    }
                })?;
                chart
                    .groups
                    .last_mut()
                    .ok_or(Error::InvalidChart {
                        offset: record.offset(),
                        reason: "CrtLink has no ChartFormat owner",
                    })?
                    .link = link;
                group_link_seen = true;
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
                if current_axis.is_some() || parent_depth != Some(depth) || parent_group_started {
                    return invalid(
                        record,
                        "Axis must own a non-overlapping AxisParent collection",
                    );
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
                let parent = current_parent
                    .and_then(|index| chart.parents.get(index))
                    .map(|parent| parent.id())
                    .ok_or(Error::InvalidChart {
                        offset: record.offset(),
                        reason: "Axis has no AxisParent owner",
                    })?;
                push(&mut chart.axes, Axis::in_parent(kind, parent), "chart axes")?;
                current_axis = chart.axes.len().checked_sub(1);
                axis_depth = depth.checked_add(1);
                pending_begin = true;
            },
            AXES_USED => {
                if axes_used.is_some() || (strict_excel && (depth != 1 || !props_seen)) {
                    return invalid(record, "AxesUsed must occur once in CHARTFOMATS");
                }
                exact(record, 2)?;
                let count = u16_at(data, 0, record)?;
                if !matches!(count, 1 | 2) {
                    return invalid(record, "AxesUsed must specify one or two axis groups");
                }
                axes_used = Some(count);
            },
            AXIS_PARENT => {
                let expected = match axes_used {
                    Some(value) => usize::from(value),
                    None if strict_excel => {
                        return invalid(record, "AxisParent appears before AxesUsed");
                    },
                    None => 2,
                };
                if parent_depth.is_some()
                    || chart.parents.len() >= expected
                    || (strict_excel && depth != 1)
                {
                    return invalid(
                        record,
                        "AxisParent is nested, misplaced, or exceeds two groups",
                    );
                }
                exact(record, 18)?;
                let secondary = match u16_at(data, 0, record)? {
                    0 if chart.parents.is_empty() => false,
                    1 if chart.parents.len() == 1 => true,
                    _ => {
                        return invalid(
                            record,
                            "AxisParent groups must be ordered primary then optional secondary",
                        );
                    },
                };
                if u16_at(data, 0, record)? > 1 {
                    return invalid(record, "AxisParent index must be primary or secondary");
                }
                check_add(chart.parents.len(), 2, "axis-parent count")?;
                push(
                    &mut chart.parents,
                    if secondary {
                        axis::Parent::secondary(layout::Pos::default())
                    } else {
                        axis::Parent::primary(layout::Pos::default())
                    },
                    "axis parents",
                )?;
                current_parent = chart.parents.len().checked_sub(1);
                parent_depth = depth.checked_add(1);
                parent_groups = 0;
                parent_group_started = false;
                pending_begin = true;
            },
            POS if parent_depth == Some(depth) => {
                if !parent_needs_pos {
                    return invalid(record, "AxisParent contains more than one Pos");
                }
                exact(record, 20)?;
                let top_left = layout::Mode::from_raw(u16_at(data, 0, record)?).ok_or(
                    Error::InvalidChart {
                        offset: record.offset(),
                        reason: "Pos upper-left mode is invalid",
                    },
                )?;
                let bottom_right = layout::Mode::from_raw(u16_at(data, 2, record)?).ok_or(
                    Error::InvalidChart {
                        offset: record.offset(),
                        reason: "Pos lower-right mode is invalid",
                    },
                )?;
                let pos = layout::Pos::parsed(
                    top_left,
                    bottom_right,
                    i16_at(data, 4, record)?,
                    i16_at(data, 8, record)?,
                    i16_at(data, 12, record)?,
                    i16_at(data, 16, record)?,
                );
                if !pos.is_plot() {
                    return invalid(record, "AxisParent Pos must use Parent/Parent modes");
                }
                let index = current_parent.ok_or(Error::InvalidChart {
                    offset: record.offset(),
                    reason: "Pos has no AxisParent owner",
                })?;
                let parent = chart.parents.get_mut(index).ok_or(Error::InvalidChart {
                    offset: record.offset(),
                    reason: "Pos AxisParent owner is missing",
                })?;
                *parent = if parent.is_secondary() {
                    axis::Parent::secondary(pos)
                } else {
                    axis::Parent::primary(pos)
                };
                parent_needs_pos = false;
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
                let data = copy(data, "marker payload", limits.biff.max_record_bytes)?;
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
                let data = copy(data, "data-label payload", limits.biff.max_record_bytes)?;
                push(
                    &mut chart.labels,
                    Label {
                        kind: record.kind(),
                        data,
                    },
                    "chart labels",
                )?;
            },
            DIMENSIONS => {
                if dimensions_seen || (strict_excel && (depth != 0 || !chart_closed)) {
                    return invalid(record, "Dimensions must occur once before SERIESDATA cells");
                }
                exact(record, 14)?;
                chart.dimensions = match context.kind() {
                    ChartKind::Excel => {
                        if u16_at(data, 12, record)? != 0 {
                            return invalid(record, "Excel Dimensions reserved field is nonzero");
                        }
                        let dims = cache::ExcelDims::new(
                            u32_at(data, 0, record)?,
                            u32_at(data, 4, record)?,
                            u16_at(data, 8, record)?,
                            u16_at(data, 10, record)?,
                        )
                        .ok_or(Error::InvalidChart {
                            offset: record.offset(),
                            reason: "Excel Dimensions range is reversed or outside the BIFF8 grid",
                        })?;
                        cache::Dims::Excel(dims)
                    },
                    ChartKind::Graph => {
                        if u32_at(data, 0, record)? != 0
                            || u16_at(data, 8, record)? != 0
                            || u16_at(data, 12, record)? != 0
                        {
                            return invalid(record, "Graph Dimensions reserved fields are nonzero");
                        }
                        let longest =
                            RowCol::new(u16::try_from(u32_at(data, 4, record)?).map_err(|_| {
                                Error::InvalidChart {
                                    offset: record.offset(),
                                    reason: "Graph Dimensions longest row exceeds u16",
                                }
                            })?)
                            .ok_or(Error::InvalidChart {
                                offset: record.offset(),
                                reason: "Graph Dimensions longest row exceeds 3,999",
                            })?;
                        let rows = u8::try_from(u16_at(data, 10, record)?).map_err(|_| {
                            Error::InvalidChart {
                                offset: record.offset(),
                                reason: "Graph Dimensions row count exceeds 255",
                            }
                        })?;
                        cache::Dims::Graph(cache::GraphDims::new(longest, rows).ok_or(
                            Error::InvalidChart {
                                offset: record.offset(),
                                reason: "Graph Dimensions empty fields disagree",
                            },
                        )?)
                    },
                };
                dimensions_seen = true;
            },
            SI_INDEX if context.kind() == ChartKind::Excel => {
                if depth != 0 || !dimensions_seen {
                    return invalid(record, "SIIndex must follow Dimensions at substream level");
                }
                exact(record, 2)?;
                let index = cache::Index::from_raw(u16_at(data, 0, record)?).ok_or(
                    Error::InvalidChart {
                        offset: record.offset(),
                        reason: "SIIndex must identify values, categories, or bubbles",
                    },
                )?;
                if cache::Index::ALL.get(next_cache_section).copied() != Some(index) {
                    return invalid(
                        record,
                        "SIIndex sections are missing, duplicated, or out of order",
                    );
                }
                next_cache_section =
                    next_cache_section
                        .checked_add(1)
                        .ok_or(Error::SizeOverflow {
                            resource: "SIIndex section count",
                        })?;
                cache_section = Some(index);
            },
            SI_INDEX => {
                return invalid(
                    record,
                    "SIIndex is not part of the standalone Graph grammar",
                );
            },
            kind if kind == number_kind(context.kind()) => {
                if depth != 0 || !dimensions_seen {
                    return invalid(record, "cached Number must follow Dimensions in SERIESDATA");
                }
                let value = match context.kind() {
                    ChartKind::Graph => {
                        exact(record, 15)?;
                        if byte_at(data, 4, record)? != 0 {
                            return invalid(record, "Graph Number reserved byte is nonzero");
                        }
                        graph_cache(
                            data,
                            cache::Ifmt::new(u16_at(data, 5, record)?),
                            Value::Number(f64_at(data, 7, record)?),
                            record,
                        )
                    },
                    ChartKind::Excel => {
                        exact(record, 14)?;
                        excel_cache(
                            data,
                            cache_section.ok_or(Error::InvalidChart {
                                offset: record.offset(),
                                reason: "Excel cache cell appears before its SIIndex",
                            })?,
                            cache::Xf::new(u16_at(data, 4, record)?),
                            XlValue::Number(f64_at(data, 6, record)?),
                            record,
                        )
                    },
                }?;
                check_add(
                    chart.caches.len(),
                    limits.max_cached_values,
                    "cached value count",
                )?;
                push(&mut chart.caches, value, "chart cache")?;
            },
            CELL_LABEL => {
                if depth != 0 || !dimensions_seen {
                    return invalid(record, "cached Label must follow Dimensions in SERIESDATA");
                }
                let value = match context.kind() {
                    ChartKind::Graph => {
                        if data.len() < 9 || byte_at(data, 4, record)? != 0 {
                            return invalid(
                                record,
                                "Graph Label is truncated or reserved byte is nonzero",
                            );
                        }
                        let string = data.get(7..).ok_or(Error::InvalidChart {
                            offset: record.offset(),
                            reason: "cached Label string is truncated",
                        })?;
                        graph_cache(
                            data,
                            cache::Ifmt::new(u16_at(data, 5, record)?),
                            Value::Text(parse_string(string, record)?),
                            record,
                        )
                    },
                    ChartKind::Excel => {
                        if data.len() < 9 {
                            return invalid(record, "Excel Label is shorter than nine bytes");
                        }
                        let string = data.get(6..).ok_or(Error::InvalidChart {
                            offset: record.offset(),
                            reason: "cached Label string is truncated",
                        })?;
                        excel_cache(
                            data,
                            cache_section.ok_or(Error::InvalidChart {
                                offset: record.offset(),
                                reason: "Excel cache cell appears before its SIIndex",
                            })?,
                            cache::Xf::new(u16_at(data, 4, record)?),
                            XlValue::Text(parse_xl_unicode_string(string, record)?),
                            record,
                        )
                    },
                }?;
                check_add(
                    chart.caches.len(),
                    limits.max_cached_values,
                    "cached value count",
                )?;
                push(&mut chart.caches, value, "chart cache")?;
            },
            kind if kind == blank_kind(context.kind()) => {
                if depth != 0 || !dimensions_seen {
                    return invalid(record, "cached Blank must follow Dimensions in SERIESDATA");
                }
                let value = match context.kind() {
                    ChartKind::Graph => {
                        exact(record, 7)?;
                        if byte_at(data, 4, record)? != 0 {
                            return invalid(record, "Graph Blank reserved byte is nonzero");
                        }
                        graph_cache(
                            data,
                            cache::Ifmt::new(u16_at(data, 5, record)?),
                            Value::Blank,
                            record,
                        )
                    },
                    ChartKind::Excel => {
                        exact(record, 6)?;
                        excel_cache(
                            data,
                            cache_section.ok_or(Error::InvalidChart {
                                offset: record.offset(),
                                reason: "Excel cache cell appears before its SIIndex",
                            })?,
                            cache::Xf::new(u16_at(data, 4, record)?),
                            XlValue::Blank,
                            record,
                        )
                    },
                }?;
                check_add(
                    chart.caches.len(),
                    limits.max_cached_values,
                    "cached value count",
                )?;
                push(&mut chart.caches, value, "chart cache")?;
            },
            EXCEL_BOOL_ERR if context.kind() == ChartKind::Excel => {
                if depth != 0 || !dimensions_seen {
                    return invalid(
                        record,
                        "cached BoolErr must follow Dimensions in SERIESDATA",
                    );
                }
                exact(record, 8)?;
                let value = match byte_at(data, 7, record)? {
                    0 => match byte_at(data, 6, record)? {
                        0 => XlValue::Bool(false),
                        1 => XlValue::Bool(true),
                        _ => return invalid(record, "BoolErr Boolean value is not zero or one"),
                    },
                    1 => XlValue::Error(cache::Fault::from_raw(byte_at(data, 6, record)?).ok_or(
                        Error::InvalidChart {
                            offset: record.offset(),
                            reason: "BoolErr contains an unknown BIFF error code",
                        },
                    )?),
                    _ => return invalid(record, "BoolErr discriminator is not zero or one"),
                };
                check_add(
                    chart.caches.len(),
                    limits.max_cached_values,
                    "cached value count",
                )?;
                let value = excel_cache(
                    data,
                    cache_section.ok_or(Error::InvalidChart {
                        offset: record.offset(),
                        reason: "Excel cache cell appears before its SIIndex",
                    })?,
                    cache::Xf::new(u16_at(data, 4, record)?),
                    value,
                    record,
                )?;
                push(&mut chart.caches, value, "chart cache")?;
            },
            BOF => return invalid(record, "nested BOF in chart substream"),
            CONTINUE | SERIES_LIST | CAT_SER_RANGE | DEFAULT_TEXT | FONT_X | OBJECT_LINK
            | FRAME | POS => {
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
        || parent_depth.is_some()
        || parent_needs_pos
    {
        return Err(Error::InvalidChart {
            offset: input.as_bytes().len(),
            reason: "chart collection owner is missing its complete Begin/End collection",
        });
    }
    if strict_excel
        && (!chart_seen || !chart_closed || !zoom_seen || !growth_seen || !dimensions_seen)
    {
        return Err(Error::InvalidChart {
            offset: input.as_bytes().len(),
            reason: "chart is missing Chart, Scl, PlotGrowth, or Dimensions",
        });
    }
    if (strict_excel || axes_used.is_some())
        && axes_used
            != Some(
                u16::try_from(chart.parents.len()).map_err(|_| Error::SizeOverflow {
                    resource: "axis-parent count",
                })?,
            )
    {
        return Err(Error::InvalidChart {
            offset: input.as_bytes().len(),
            reason: "AxesUsed does not match the AxisParent collection count",
        });
    }
    if context.kind() == ChartKind::Excel && next_cache_section != cache::Index::ALL.len() {
        return Err(Error::InvalidChart {
            offset: input.as_bytes().len(),
            reason: "Excel SERIESDATA does not contain exactly three SIIndex sections",
        });
    }
    validate(&chart, limits, strict_excel)?;
    Ok(chart)
}

fn parse_link(record: RecordRef<'_>, context: Context, limits: Limits) -> Result<Link> {
    let data = record.payload();
    match context.kind() {
        ChartKind::Graph => {
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
        ChartKind::Excel => {
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
        limits.biff.max_record_bytes,
    )?;
    push(
        &mut chart.unknown,
        Raw::parsed(record.kind(), data, record.offset()),
        "unknown chart records",
    )
}

fn validate(chart: &Chart, limits: Limits, require_topology: bool) -> Result<()> {
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
            validate_link(binding.link(), chart.context, limits)?;
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

fn validate_link(link: &Link, context: Context, limits: Limits) -> Result<()> {
    match (context.kind(), link) {
        (ChartKind::Graph, Link::Graph { .. }) => {},
        (
            ChartKind::Excel,
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
            let maximum = limits.biff.max_record_bytes.saturating_sub(8);
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
        (ChartKind::Graph, Link::Excel { .. }) => {
            return invalid_model("link", "Excel BRAI cannot be encoded in a Graph chart");
        },
        (ChartKind::Excel, Link::Graph { .. }) => {
            return invalid_model("link", "Graph BRAI cannot be encoded in an Excel chart");
        },
    }
    Ok(())
}

pub(crate) fn encode(chart: &Chart, limits: Limits) -> Result<Vec<u8>> {
    if !chart.authoring_proven {
        return Err(Error::UnsupportedAuthoring {
            reason: "the host prelude, attached-label/frame grammar, and complete axis ownership are not yet proven",
        });
    }
    if chart.parents.len() != 1 || !(1..=4).contains(&chart.groups.len()) {
        return Err(Error::UnsupportedAuthoring {
            reason: "secondary-axis chart-group ownership is not yet modeled",
        });
    }
    validate(chart, limits, true)?;
    let mut out = Encoder::with_limits(limits.biff)?;
    push_record(&mut out, BOF, &bof(chart.context.kind()))?;

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
    push_record(&mut out, CHART_REC, &rect)?;
    push_record(&mut out, BEGIN, &[])?;
    let mut zoom = [0u8; 4];
    put_u16(&mut zoom, 0, chart.zoom.numerator())?;
    put_u16(&mut zoom, 2, chart.zoom.denominator())?;
    push_record(&mut out, SCL, &zoom)?;
    let mut growth = [0u8; 8];
    put_i32(&mut growth, 0, chart.growth.x.raw())?;
    put_i32(&mut growth, 4, chart.growth.y.raw())?;
    push_record(&mut out, PLOT_GROWTH, &growth)?;

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
        push_record(&mut out, SERIES, &data)?;
        push_record(&mut out, BEGIN, &[])?;
        for binding in series.ai.ordered() {
            let data = encode_link(binding.link(), chart.context, limits)?;
            push_record(&mut out, BRAI, &data)?;
            if let Some(text) = binding.text() {
                push_record(&mut out, SERIES_TEXT, &short_text(text)?)?;
            }
        }
        match &series.owner {
            Owner::Group(group) => {
                push_record(&mut out, SER_TO_CRT, &u16::from(group.get()).to_le_bytes())?;
            },
            Owner::Trend { parent, data } => {
                parent.write(&mut out)?;
                push_record(&mut out, SER_AUX_TREND, data)?;
            },
            Owner::ErrorBar { parent, data } => {
                parent.write(&mut out)?;
                push_record(&mut out, SER_AUX_ERR_BAR, data)?;
            },
        }
        push_record(&mut out, END, &[])?;
    }

    push_record(&mut out, SHT_PROPS, &chart.props.flags.to_le_bytes())?;
    push_record(
        &mut out,
        AXES_USED,
        &u16::try_from(chart.parents.len())
            .map_err(|_| Error::SizeOverflow {
                resource: "axis-parent count",
            })?
            .to_le_bytes(),
    )?;
    let parent = chart.parents.first().copied().ok_or(Error::InvalidModel {
        field: "axis parents",
        reason: "a chart requires a primary AxisParent",
    })?;
    let mut parent_body = [0u8; 18];
    put_u16(&mut parent_body, 0, u16::from(parent.is_secondary()))?;
    push_record(&mut out, AXIS_PARENT, &parent_body)?;
    push_record(&mut out, BEGIN, &[])?;
    encode_pos(&mut out, parent.pos())?;
    for axis in &chart.axes {
        encode_axis(&mut out, axis)?;
    }
    for group in &chart.groups {
        encode_group(&mut out, group)?;
    }
    if chart.props.plot_area {
        push_record(&mut out, PLOT_AREA, &[])?;
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
        push_record(&mut out, LEGEND, &data)?;
    }
    if let Some(title) = &chart.title {
        push_record(&mut out, SERIES_TEXT, &short_text(title)?)?;
    }
    for value in &chart.formats {
        encode_format(&mut out, value)?;
    }
    for label in &chart.labels {
        push_record(&mut out, label.kind, &label.data)?;
    }

    push_record(&mut out, END, &[])?;
    push_record(&mut out, END, &[])?;
    encode_dimensions(&mut out, chart.dimensions)?;
    match chart.context.kind() {
        ChartKind::Excel => {
            for index in cache::Index::ALL {
                push_record(&mut out, SI_INDEX, &(index as u16).to_le_bytes())?;
                for value in chart.caches.iter().filter(
                    |value| matches!(value, Cache::Excel { section, .. } if *section == index),
                ) {
                    encode_cache(&mut out, value)?;
                }
            }
        },
        ChartKind::Graph => {
            for value in &chart.caches {
                encode_cache(&mut out, value)?;
            }
        },
    }
    push_record(&mut out, EOF, &[])?;
    Ok(out.finish())
}

fn push_record(out: &mut Encoder, kind: RecordKind, payload: &[u8]) -> Result<()> {
    out.push(kind, payload)?;
    Ok(())
}

fn encode_pos(out: &mut Encoder, pos: layout::Pos) -> Result<()> {
    let mut data = [0u8; 20];
    put_u16(&mut data, 0, pos.top_left() as u16)?;
    put_u16(&mut data, 2, pos.bottom_right() as u16)?;
    put_i16(&mut data, 4, pos.x())?;
    put_i16(&mut data, 8, pos.y())?;
    put_i16(&mut data, 12, pos.width())?;
    put_i16(&mut data, 16, pos.height())?;
    push_record(out, POS, &data)
}

fn encode_dimensions(out: &mut Encoder, dimensions: cache::Dims) -> Result<()> {
    let mut data = [0u8; 14];
    match dimensions {
        cache::Dims::Excel(value) => {
            put_u32(&mut data, 0, value.first_row())?;
            put_u32(&mut data, 4, value.row_after())?;
            put_u16(&mut data, 8, value.first_col())?;
            put_u16(&mut data, 10, value.col_after())?;
        },
        cache::Dims::Graph(value) => {
            put_u32(&mut data, 4, u32::from(value.longest_row().get()))?;
            put_u16(&mut data, 10, u16::from(value.rows()))?;
        },
    }
    push_record(out, DIMENSIONS, &data)
}

fn encode_cache(out: &mut Encoder, cache: &Cache) -> Result<()> {
    match cache {
        Cache::Excel {
            row,
            col,
            xf,
            value,
            ..
        } => encode_excel_cache(out, *row, *col, *xf, value),
        Cache::Graph {
            row,
            col,
            ifmt,
            value,
        } => encode_graph_cache(out, *row, *col, *ifmt, value),
    }
}

fn encode_excel_cache(
    out: &mut Encoder,
    row: u16,
    col: u8,
    xf: cache::Xf,
    value: &XlValue,
) -> Result<()> {
    match value {
        XlValue::Number(number) => {
            let mut data = [0u8; 14];
            put_u16(&mut data, 0, row)?;
            put_u16(&mut data, 2, u16::from(col))?;
            put_u16(&mut data, 4, xf.get())?;
            put_f64(&mut data, 6, *number)?;
            push_record(out, EXCEL_NUMBER, &data)
        },
        XlValue::Text(text) => {
            let string = xl_unicode_string(text)?;
            let capacity = 6usize
                .checked_add(string.len())
                .ok_or(Error::SizeOverflow {
                    resource: "cached chart string",
                })?;
            let mut data = vec_with_capacity(capacity, "cached chart string")?;
            data.extend_from_slice(&row.to_le_bytes());
            data.extend_from_slice(&u16::from(col).to_le_bytes());
            data.extend_from_slice(&xf.get().to_le_bytes());
            data.extend_from_slice(&string);
            push_record(out, CELL_LABEL, &data)
        },
        XlValue::Bool(value) => {
            let mut data = [0u8; 8];
            put_u16(&mut data, 0, row)?;
            put_u16(&mut data, 2, u16::from(col))?;
            put_u16(&mut data, 4, xf.get())?;
            put_byte(&mut data, 6, u8::from(*value))?;
            push_record(out, EXCEL_BOOL_ERR, &data)
        },
        XlValue::Error(value) => {
            let mut data = [0u8; 8];
            put_u16(&mut data, 0, row)?;
            put_u16(&mut data, 2, u16::from(col))?;
            put_u16(&mut data, 4, xf.get())?;
            put_byte(&mut data, 6, *value as u8)?;
            put_byte(&mut data, 7, 1)?;
            push_record(out, EXCEL_BOOL_ERR, &data)
        },
        XlValue::Blank => {
            let mut data = [0u8; 6];
            put_u16(&mut data, 0, row)?;
            put_u16(&mut data, 2, u16::from(col))?;
            put_u16(&mut data, 4, xf.get())?;
            push_record(out, EXCEL_BLANK, &data)
        },
    }
}

fn encode_graph_cache(
    out: &mut Encoder,
    row: RowCol,
    col: RowCol,
    ifmt: cache::Ifmt,
    value: &Value,
) -> Result<()> {
    match value {
        Value::Number(number) => {
            let mut data = [0u8; 15];
            put_u16(&mut data, 0, row.get())?;
            put_u16(&mut data, 2, col.get())?;
            put_u16(&mut data, 5, ifmt.get())?;
            put_f64(&mut data, 7, *number)?;
            push_record(out, GRAPH_NUMBER, &data)
        },
        Value::Text(text) => {
            let string = biff_string(text)?;
            let capacity = 7usize
                .checked_add(string.len())
                .ok_or(Error::SizeOverflow {
                    resource: "cached chart string",
                })?;
            let mut data = vec_with_capacity(capacity, "cached chart string")?;
            data.extend_from_slice(&row.get().to_le_bytes());
            data.extend_from_slice(&col.get().to_le_bytes());
            data.push(0);
            data.extend_from_slice(&ifmt.get().to_le_bytes());
            data.extend_from_slice(&string);
            push_record(out, CELL_LABEL, &data)
        },
        Value::Blank => {
            let mut data = [0u8; 7];
            put_u16(&mut data, 0, row.get())?;
            put_u16(&mut data, 2, col.get())?;
            put_u16(&mut data, 5, ifmt.get())?;
            push_record(out, GRAPH_BLANK, &data)
        },
    }
}

fn bof(kind: ChartKind) -> [u8; BOF_BYTES] {
    let mut data = [0u8; BOF_BYTES];
    match kind {
        ChartKind::Excel => {
            data[0..2].copy_from_slice(&EXCEL_VERSION.to_le_bytes());
            data[2..4].copy_from_slice(&EXCEL_DOC_TYPE.to_le_bytes());
            data[4..6].copy_from_slice(&0x0DBB_u16.to_le_bytes());
            data[6..8].copy_from_slice(&0x07CC_u16.to_le_bytes());
            data[8..12].copy_from_slice(&0x0000_0009_u32.to_le_bytes());
            data[12..16].copy_from_slice(&6u32.to_le_bytes());
        },
        ChartKind::Graph => {
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
    push_record(out, AXIS, &body)?;
    push_record(out, BEGIN, &[])?;
    if let Some(scale) = axis.scale {
        let mut data = [0u8; 42];
        put_f64(&mut data, 0, scale.min)?;
        put_f64(&mut data, 8, scale.max)?;
        put_f64(&mut data, 16, scale.major)?;
        put_f64(&mut data, 24, scale.minor)?;
        put_f64(&mut data, 32, scale.crossing)?;
        put_u16(&mut data, 40, scale.flags)?;
        push_record(out, VALUE_RANGE, &data)?;
    }
    if let Some(tick) = axis.tick {
        let mut data = [0u8; 26];
        put_byte(&mut data, 0, tick.major)?;
        put_byte(&mut data, 1, tick.minor)?;
        put_byte(&mut data, 2, tick.label)?;
        put_byte(&mut data, 3, tick.background)?;
        put_slice(&mut data, 4, &tick.color, "Tick color")?;
        put_u16(&mut data, 24, tick.flags)?;
        push_record(out, TICK, &data)?;
    }
    for line in &axis.lines {
        push_record(
            out,
            AXIS_LINE,
            &u16::from(line_kind(line.kind)).to_le_bytes(),
        )?;
        push_record(out, LINE_FORMAT, &line_bytes(line.format))?;
    }
    push_record(out, END, &[])
}

fn encode_group(out: &mut Encoder, group: &Group) -> Result<()> {
    let mut header = [0u8; 20];
    put_u16(&mut header, 16, u16::from(group.vary_colors))?;
    put_u16(&mut header, 18, u16::from(group.order.get()))?;
    push_record(out, CHART_FORMAT, &header)?;
    push_record(out, BEGIN, &[])?;
    match group.family {
        Family::Line { flags } => push_record(out, LINE, &flags.to_le_bytes())?,
        Family::Area { flags } => push_record(out, AREA, &flags.to_le_bytes())?,
        Family::Bar {
            overlap,
            gap,
            flags,
        } => {
            let mut data = [0u8; 6];
            put_i16(&mut data, 0, overlap.get())?;
            put_u16(&mut data, 2, gap.get())?;
            put_u16(&mut data, 4, flags)?;
            push_record(out, BAR, &data)?;
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
            push_record(out, PIE, &data)?;
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
            push_record(out, SCATTER, &data)?;
        },
        Family::Radar { filled, flags } => {
            push_record(
                out,
                if filled { RADAR_AREA } else { RADAR },
                &flags.to_le_bytes(),
            )?;
        },
        Family::Surface { flags } => push_record(out, SURFACE, &flags.to_le_bytes())?,
    }
    push_record(out, CRT_LINK, &group.link.payload())?;
    for line in &group.lines {
        let marker = crate::record::line::Line::new(line.kind);
        push_record(out, CRT_LINE, &marker.payload())?;
        push_record(out, LINE_FORMAT, &line_bytes(line.format))?;
    }
    for drop in &group.drop_bars {
        push_record(out, DROP_BAR, &drop.gap.get().to_le_bytes())?;
        push_record(out, BEGIN, &[])?;
        push_record(out, LINE_FORMAT, &line_bytes(drop.line))?;
        push_record(out, AREA_FORMAT, &area_bytes(drop.area))?;
        push_record(out, END, &[])?;
    }
    push_record(out, END, &[])
}

fn encode_format(out: &mut Encoder, value: &Format) -> Result<()> {
    match value {
        Format::Line(value) => push_record(out, LINE_FORMAT, &line_bytes(*value)),
        Format::Area(value) => push_record(out, AREA_FORMAT, &area_bytes(*value)),
        Format::Marker { data } => push_record(out, MARKER_FORMAT, data),
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
            push_record(out, DATA_FORMAT, &data)
        },
        Format::Pie { explosion } => push_record(out, PIE_FORMAT, &explosion.to_le_bytes()),
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
    parse_string_content(data, 2, count, flags, record)
}

fn parse_xl_unicode_string(data: &[u8], record: RecordRef<'_>) -> Result<String> {
    if data.len() < 3 {
        return invalid(record, "Excel chart string is shorter than three bytes");
    }
    let count = usize::from(u16_at(data, 0, record)?);
    let flags = byte_at(data, 2, record)?;
    parse_string_content(data, 3, count, flags, record)
}

fn parse_string_content(
    data: &[u8],
    header: usize,
    count: usize,
    flags: u8,
    record: RecordRef<'_>,
) -> Result<String> {
    if flags & !1 != 0 {
        return invalid(record, "chart string uses unsupported option flags");
    }
    let wide = flags & 1 != 0;
    let width = if wide { 2usize } else { 1usize };
    let content = count.checked_mul(width).ok_or(Error::SizeOverflow {
        resource: "chart string",
    })?;
    let expected = header.checked_add(content).ok_or(Error::SizeOverflow {
        resource: "chart string",
    })?;
    if data.len() != expected {
        return invalid(
            record,
            "chart string length does not match its character count",
        );
    }
    let bytes = data.get(header..).ok_or(Error::InvalidChart {
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
        let units = bytes.chunks_exact(2).map(|value| match value {
            [low, high] => u16::from_le_bytes([*low, *high]),
            _ => 0,
        });
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

fn xl_unicode_string(value: &str) -> Result<Vec<u8>> {
    let count = value.encode_utf16().count();
    let count = u16::try_from(count).map_err(|_| Error::InvalidModel {
        field: "cached text",
        reason: "Excel chart string exceeds 65,535 UTF-16 code units",
    })?;
    let wide = value.encode_utf16().any(|unit| unit > u16::from(u8::MAX));
    let width = if wide { 2usize } else { 1usize };
    let capacity = 3usize
        .checked_add(
            usize::from(count)
                .checked_mul(width)
                .ok_or(Error::SizeOverflow {
                    resource: "Excel chart string",
                })?,
        )
        .ok_or(Error::SizeOverflow {
            resource: "Excel chart string",
        })?;
    let mut data = vec_with_capacity(capacity, "Excel chart string")?;
    data.extend_from_slice(&count.to_le_bytes());
    data.push(u8::from(wide));
    for unit in value.encode_utf16() {
        if wide {
            data.extend_from_slice(&unit.to_le_bytes());
        } else {
            data.push(u8::try_from(unit).map_err(|_| Error::InvalidModel {
                field: "cached text",
                reason: "narrow Excel chart string contains a wide code unit",
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

fn blank_kind(kind: ChartKind) -> RecordKind {
    match kind {
        ChartKind::Graph => GRAPH_BLANK,
        ChartKind::Excel => EXCEL_BLANK,
    }
}

fn number_kind(kind: ChartKind) -> RecordKind {
    match kind {
        ChartKind::Graph => GRAPH_NUMBER,
        ChartKind::Excel => EXCEL_NUMBER,
    }
}

fn excel_cache(
    data: &[u8],
    section: cache::Index,
    xf: cache::Xf,
    value: XlValue,
    record: RecordRef<'_>,
) -> Result<Cache> {
    let row = u16_at(data, 0, record)?;
    let col = u16_at(data, 2, record)?;
    Ok(Cache::excel(
        section,
        row,
        u8::try_from(col).map_err(|_| Error::InvalidChart {
            offset: record.offset(),
            reason: "Excel cache column exceeds the BIFF8 grid",
        })?,
        xf,
        value,
    ))
}

fn graph_cache(
    data: &[u8],
    ifmt: cache::Ifmt,
    value: Value,
    record: RecordRef<'_>,
) -> Result<Cache> {
    Ok(Cache::graph(
        RowCol::new(u16_at(data, 0, record)?).ok_or(Error::InvalidChart {
            offset: record.offset(),
            reason: "Graph cache row exceeds 3,999",
        })?,
        RowCol::new(u16_at(data, 2, record)?).ok_or(Error::InvalidChart {
            offset: record.offset(),
            reason: "Graph cache column exceeds 3,999",
        })?,
        ifmt,
        value,
    ))
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

fn check_xl_string(value: &str, limits: Limits) -> Result<()> {
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

fn valid_props(flags: u32) -> bool {
    let reserved_clear = flags & 0x0000_FFE0 == 0 && flags & 0xFF00_0000 == 0;
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

fn put_u32(output: &mut [u8], offset: usize, value: u32) -> Result<()> {
    put_slice(output, offset, &value.to_le_bytes(), "encoded u32")
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
