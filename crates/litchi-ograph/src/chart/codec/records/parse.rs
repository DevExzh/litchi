//! MS-OGRAPH chart-substream traversal.

use super::super::super::axis::{self, Axis};
use super::super::super::cache as chart_cache;
use super::super::super::format::{self, Format};
use super::super::super::group;
use super::super::super::layout;
use super::super::super::model::{
    Binding, Chart, Context, DataKind, Family, Group, GroupId, Label, Legend, Order, Origin, Owner,
    Props, Raw, Rect, Role, RowCol, Series, Value, XlValue,
};
use super::super::super::{BOF, EOF, Kind as ChartKind, Ref};
use super::super::model::{PendingDrop, PendingLine};
use super::cache::{blank_kind, excel_cache, graph_cache, number_kind};
use super::links;
use super::text::{parse_short_text, parse_string, parse_xl_unicode_string};
use super::validate;
use super::validate::valid_props;
use super::wire::{
    AREA, AREA_FORMAT, AXC_EXT, AXES_USED, AXIS, AXIS_LINE, AXIS_PARENT, BAR, BEGIN, BOF_DATASHEET,
    BRAI, CAT_SER_RANGE, CELL_LABEL, CHART_FORMAT, CHART_REC, CONTINUE, COUNTRY, CRT_LINE,
    CRT_LINK, DATA_FORMAT, DATA_LAB_EXT, DATA_LAB_EXT_CONTENTS, DEFAULT_TEXT, DIMENSIONS, DROP_BAR,
    END, EXCEL_BOOL_ERR, EXCLUDE_COLUMNS, EXCLUDE_ROWS, FONT_X, FRAME, LEGEND, LINE, LINE_FORMAT,
    MAIN_WINDOW, MARKER_FORMAT, MAX_STATUS, OBJECT_LINK, ORIENT, PIE, PIE_FORMAT, PLOT_AREA,
    PLOT_GROWTH, POS, RADAR, RADAR_AREA, SCATTER, SCL, SELECTION, SER_AUX_ERR_BAR, SER_AUX_TREND,
    SER_PARENT, SER_TO_CRT, SERIES, SERIES_LIST, SERIES_TEXT, SHT_PROPS, SI_INDEX, SURFACE, TEXT,
    TICK, UNITS, VALUE_RANGE, WIN_DOC, WINDOW1, WINDOW2_GRAPH, array4_at, byte_at, check_add, copy,
    count_at, exact, f64_at, i16_at, i32_at, invalid, invalid_model, limit, push, u16_at, u32_at,
};
use crate::{Error, Limits, Result};
use litchi_biff::RecordRef;

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
        dimensions: chart_cache::Dims::empty(context.kind()),
        formats: Vec::new(),
        labels: Vec::new(),
        unknown: Vec::new(),
        origin: Origin::Fresh,
        dirty: false,
        limits,
        authoring_proven: false,
        graph_authoring_profile: false,
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
    let mut cache_sections = [false; chart_cache::Index::ALL.len()];
    let mut zoom_seen = false;
    let mut growth_seen = false;
    let mut dimensions_seen = false;
    let mut datasheet_seen = false;
    let mut datasheet_pending = false;
    let mut datasheet_active = false;
    let mut country_seen = false;
    let mut font_x_seen = false;
    let mut exclude_rows_seen = false;
    let mut exclude_columns_seen = false;
    let mut orient_seen = false;
    let mut windoc_seen = false;
    let mut main_window_seen = false;
    let mut window1_seen = false;
    let mut window2_seen = false;
    let mut max_status_seen = false;
    let mut selection_seen = false;
    let mut units_seen = false;
    let mut attached_label_pending = false;
    let mut attached_label_active = false;
    let mut attached_label_depth = None;
    let mut attached_label_brai_seen = false;
    let mut attached_label_text = None;
    let mut attached_label_link_object = None;
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
        if datasheet_pending && record.kind() != BEGIN {
            return invalid(record, "BOFDatasheet is not followed immediately by Begin");
        }
        if attached_label_pending && record.kind() != BEGIN {
            return invalid(record, "attached Text is not followed immediately by Begin");
        }
        if pending_begin && record.kind() != BEGIN {
            return invalid(
                record,
                "collection-owning record is not followed immediately by Begin",
            );
        }
        let data = record.payload();
        let graph_cache_open = context.kind() == ChartKind::Graph
            && ((datasheet_active && depth == 1 && orient_seen && !windoc_seen)
                || (!datasheet_active && depth == 0 && dimensions_seen));
        match record.kind() {
            BEGIN => {
                exact(record, 0)?;
                if datasheet_pending {
                    if context.kind() != ChartKind::Graph || depth != 0 || datasheet_active {
                        return invalid(record, "datasheet Begin is misplaced");
                    }
                    datasheet_pending = false;
                    datasheet_active = true;
                } else if datasheet_active {
                    return invalid(record, "datasheet collection cannot contain nested Begin");
                }
                if attached_label_pending {
                    if depth != 1 || attached_label_active {
                        return invalid(record, "attached-label Begin is misplaced");
                    }
                    attached_label_pending = false;
                    attached_label_active = true;
                    attached_label_depth = depth.checked_add(1);
                    attached_label_brai_seen = false;
                    attached_label_text = None;
                    attached_label_link_object = None;
                }
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
                if datasheet_active && depth == 1 {
                    if !orient_seen
                        || !country_seen
                        || !font_x_seen
                        || !exclude_rows_seen
                        || !exclude_columns_seen
                        || !windoc_seen
                        || !main_window_seen
                        || !window1_seen
                        || !window2_seen
                        || !max_status_seen
                        || !selection_seen
                    {
                        return invalid(
                            record,
                            "datasheet is missing Orient or mandatory window/selection records",
                        );
                    }
                    datasheet_active = false;
                    depth -= 1;
                    continue;
                }
                if attached_label_active && attached_label_depth == Some(depth) {
                    if attached_label_link_object == Some(1)
                        && attached_label_brai_seen
                        && attached_label_text.is_some()
                    {
                        if chart.title.is_some() {
                            return invalid(record, "chart contains more than one title label");
                        }
                        chart.title = attached_label_text.take();
                    }
                    attached_label_active = false;
                    attached_label_depth = None;
                    attached_label_brai_seen = false;
                    attached_label_text = None;
                    attached_label_link_object = None;
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
                        u16::try_from(chart.parents.len())
                            .ok()
                            .ok_or(Error::SizeOverflow {
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
            BOF_DATASHEET if context.kind() == ChartKind::Graph => {
                if datasheet_seen
                    || datasheet_pending
                    || datasheet_active
                    || chart_seen
                    || depth != 0
                {
                    return invalid(record, "BOFDatasheet is duplicated or misplaced");
                }
                exact(record, 4)?;
                datasheet_seen = true;
                datasheet_pending = true;
            },
            COUNTRY if context.kind() == ChartKind::Graph => {
                if !datasheet_active
                    || depth != 1
                    || country_seen
                    || font_x_seen
                    || exclude_rows_seen
                    || exclude_columns_seen
                    || orient_seen
                {
                    return invalid(record, "Country is duplicated or out of Datasheet order");
                }
                exact(record, 4)?;
                country_seen = true;
            },
            FONT_X if context.kind() == ChartKind::Graph && datasheet_active => {
                if depth != 1
                    || !country_seen
                    || font_x_seen
                    || exclude_rows_seen
                    || exclude_columns_seen
                    || orient_seen
                {
                    return invalid(record, "Datasheet FontX is duplicated or out of order");
                }
                exact(record, 2)?;
                font_x_seen = true;
            },
            EXCLUDE_ROWS if context.kind() == ChartKind::Graph => {
                if !datasheet_active
                    || depth != 1
                    || !font_x_seen
                    || exclude_rows_seen
                    || exclude_columns_seen
                    || orient_seen
                {
                    return invalid(
                        record,
                        "ExcludeRows is duplicated or out of Datasheet order",
                    );
                }
                if data.len() % 2 != 0 || data.len() / 2 > 255 {
                    return invalid(record, "ExcludeRows has an invalid transition array");
                }
                exclude_rows_seen = true;
            },
            EXCLUDE_COLUMNS if context.kind() == ChartKind::Graph => {
                if !datasheet_active
                    || depth != 1
                    || !exclude_rows_seen
                    || exclude_columns_seen
                    || orient_seen
                {
                    return invalid(
                        record,
                        "ExcludeColumns is duplicated or out of Datasheet order",
                    );
                }
                if data.len() % 2 != 0 {
                    return invalid(record, "ExcludeColumns has an invalid transition array");
                }
                exclude_columns_seen = true;
            },
            ORIENT if context.kind() == ChartKind::Graph => {
                if !datasheet_active
                    || depth != 1
                    || !country_seen
                    || !font_x_seen
                    || !exclude_rows_seen
                    || !exclude_columns_seen
                    || orient_seen
                {
                    return invalid(record, "Orient is duplicated or outside Datasheet");
                }
                exact(record, 6)?;
                let rows = byte_at(data, 0, record)?;
                if rows > 1 || byte_at(data, 5, record)? != 1 {
                    return invalid(record, "Orient Boolean or reserved field is invalid");
                }
                let row_series = u16_at(data, 1, record)?;
                let col_series = u16_at(data, 3, record)?;
                if row_series > 0x0F9F || col_series > 0x0F9F {
                    return invalid(record, "Orient series coordinate exceeds the Graph grid");
                }
                orient_seen = true;
            },
            WIN_DOC if context.kind() == ChartKind::Graph => {
                if !datasheet_active || depth != 1 || !orient_seen || windoc_seen {
                    return invalid(record, "WinDoc is duplicated or out of Datasheet order");
                }
                exact(record, 1)?;
                if byte_at(data, 0, record)? > 1 {
                    return invalid(record, "WinDoc Boolean is invalid");
                }
                windoc_seen = true;
            },
            MAIN_WINDOW if context.kind() == ChartKind::Graph => {
                if !datasheet_active || depth != 1 || !windoc_seen || main_window_seen {
                    return invalid(record, "MainWindow is duplicated or out of Datasheet order");
                }
                exact(record, 4)?;
                if i16_at(data, 0, record)? < 1 || i16_at(data, 2, record)? < 1 {
                    return invalid(record, "MainWindow scale must be positive");
                }
                main_window_seen = true;
            },
            WINDOW1 if context.kind() == ChartKind::Graph => {
                if !datasheet_active || depth != 1 || !main_window_seen || window1_seen {
                    return invalid(record, "Window1_10 is duplicated or out of Datasheet order");
                }
                exact(record, 10)?;
                if u16_at(data, 2, record)? == 0
                    || u16_at(data, 4, record)? == 0
                    || u16_at(data, 8, record)? != 0
                {
                    return invalid(
                        record,
                        "Window1_10 dimensions or reserved bytes are invalid",
                    );
                }
                window1_seen = true;
            },
            WINDOW2_GRAPH if context.kind() == ChartKind::Graph => {
                if !datasheet_active || depth != 1 || !window1_seen || window2_seen {
                    return invalid(
                        record,
                        "Window2Graph is duplicated or out of Datasheet order",
                    );
                }
                exact(record, 14)?;
                if data.get(0..5) != Some(&[1, 1, 1, 0, 1])
                    || byte_at(data, 9, record)? != 1
                    || u16_at(data, 10, record)? != 0
                    || u16_at(data, 12, record)? != 0
                    || u16_at(data, 5, record)? < 1
                    || u16_at(data, 7, record)? < 1
                {
                    return invalid(record, "Window2Graph fields are invalid");
                }
                window2_seen = true;
            },
            MAX_STATUS if context.kind() == ChartKind::Graph => {
                if !datasheet_active || depth != 1 || !window2_seen || max_status_seen {
                    return invalid(record, "MaxStatus is duplicated or out of Datasheet order");
                }
                exact(record, 2)?;
                max_status_seen = true;
            },
            SELECTION if context.kind() == ChartKind::Graph => {
                if !datasheet_active || depth != 1 || !max_status_seen || selection_seen {
                    return invalid(record, "Selection is duplicated or out of Datasheet order");
                }
                exact(record, 17)?;
                if byte_at(data, 0, record)? != 3 || u16_at(data, 5, record)? != 0 {
                    return invalid(
                        record,
                        "Selection discriminator or reserved fields are invalid",
                    );
                }
                let rw_act = u16_at(data, 1, record)?;
                let col_act = u16_at(data, 3, record)?;
                let rw_first = u16_at(data, 9, record)?;
                let rw_last = u16_at(data, 11, record)?;
                let col_first = u16_at(data, 13, record)?;
                let col_last = u16_at(data, 15, record)?;
                if rw_first > rw_last
                    || col_first > col_last
                    || rw_act < rw_first
                    || rw_act > rw_last
                    || col_act < col_first
                    || col_act > col_last
                    || RowCol::new(rw_last).is_none()
                    || RowCol::new(col_last).is_none()
                {
                    return invalid(
                        record,
                        "Selection range is reversed or outside the Graph grid",
                    );
                }
                selection_seen = true;
            },
            UNITS if context.kind() == ChartKind::Graph => {
                if !datasheet_seen
                    || datasheet_pending
                    || datasheet_active
                    || units_seen
                    || chart_seen
                    || depth != 0
                {
                    return invalid(record, "Units is duplicated or misplaced");
                }
                exact(record, 2)?;
                if data.iter().any(|value| *value != 0) {
                    return invalid(record, "Units reserved field is nonzero");
                }
                units_seen = true;
            },
            CHART_REC => {
                if chart_seen
                    || depth != 0
                    || datasheet_pending
                    || datasheet_active
                    || (datasheet_seen && !units_seen)
                {
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
                    // `[MS-XLS]` `TEXTPROPS` may carry a BRAI in a label or
                    // other text collection. It is not one of this semantic
                    // chart model's four series bindings, so preserve it
                    // losslessly instead of rejecting the host chart.
                    if attached_label_active && attached_label_depth == Some(depth) {
                        if attached_label_brai_seen {
                            return invalid(
                                record,
                                "attached label contains more than one AI binding",
                            );
                        }
                        let link = links::parse_link(record, context, limits)?;
                        if link.role() != Role::Name {
                            return invalid(
                                record,
                                "chart title attached label does not use a Name AI",
                            );
                        }
                        attached_label_brai_seen = true;
                    }
                    add_raw(&mut chart, record, &mut unknown_bytes, limits)?;
                    continue;
                }
                let link = links::parse_link(record, context, limits)?;
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
                let raw = u8::try_from(raw).ok().ok_or(Error::InvalidChart {
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
                series_parent = Some(crate::record::series::Parent::parse(record).ok().ok_or({
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
                        data: data.try_into().ok().ok_or(Error::InvalidChart {
                            offset: record.offset(),
                            reason: "SerAuxTrend payload is not 28 bytes",
                        })?,
                    }
                } else {
                    exact(record, 14)?;
                    Owner::ErrorBar {
                        parent,
                        data: data.try_into().ok().ok_or(Error::InvalidChart {
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
                        .ok()
                        .ok_or(Error::InvalidChart {
                            offset: record.offset(),
                            reason: "one AI has more than one SeriesText",
                        })?;
                } else {
                    if attached_label_active && attached_label_depth == Some(depth) {
                        if !attached_label_brai_seen || attached_label_text.is_some() {
                            return invalid(record, "attached label SeriesText is out of AI order");
                        }
                        attached_label_text = Some(parse_short_text(record)?);
                    }
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
                let raw = u8::try_from(raw).ok().ok_or(Error::InvalidChart {
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
                let link = crate::record::line::Link::from_payload(data).ok().ok_or({
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
                let value = crate::record::line::Line::from_payload(data).ok().ok_or({
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
            AXC_EXT if context.kind() == ChartKind::Graph => {
                if axis_depth != Some(depth)
                    || current_axis
                        .and_then(|index| chart.axes.get(index))
                        .is_none_or(|axis| axis.kind != axis::Kind::Category)
                {
                    return invalid(record, "AxcExt appears outside a category Axis collection");
                }
                exact(record, 18)?;
                let flags = u16_at(data, 16, record)?;
                if flags & 0xFF00 != 0 {
                    return invalid(record, "AxcExt reserved flags are nonzero");
                }
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
                if record.kind() == TEXT && context.kind() == ChartKind::Graph && depth == 1 {
                    if attached_label_pending || attached_label_active {
                        return invalid(record, "chart attached labels overlap");
                    }
                    attached_label_pending = true;
                }
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
                if dimensions_seen
                    || (strict_excel && (depth != 0 || !chart_closed))
                    || (context.kind() == ChartKind::Graph
                        && datasheet_seen
                        && (!units_seen || depth != 0 || !chart_closed))
                {
                    return invalid(record, "Dimensions must occur once before SERIESDATA cells");
                }
                exact(record, 14)?;
                chart.dimensions = match context.kind() {
                    ChartKind::Excel => {
                        if u16_at(data, 12, record)? != 0 {
                            return invalid(record, "Excel Dimensions reserved field is nonzero");
                        }
                        let dims = chart_cache::ExcelDims::new(
                            u32_at(data, 0, record)?,
                            u32_at(data, 4, record)?,
                            u16_at(data, 8, record)?,
                            u16_at(data, 10, record)?,
                        )
                        .ok_or(Error::InvalidChart {
                            offset: record.offset(),
                            reason: "Excel Dimensions range is reversed or outside the BIFF8 grid",
                        })?;
                        chart_cache::Dims::Excel(dims)
                    },
                    ChartKind::Graph => {
                        if u32_at(data, 0, record)? != 0
                            || u16_at(data, 8, record)? != 0
                            || u16_at(data, 12, record)? != 0
                        {
                            return invalid(record, "Graph Dimensions reserved fields are nonzero");
                        }
                        let longest =
                            RowCol::new(u16::try_from(u32_at(data, 4, record)?).ok().ok_or({
                                Error::InvalidChart {
                                    offset: record.offset(),
                                    reason: "Graph Dimensions longest row exceeds u16",
                                }
                            })?)
                            .ok_or(Error::InvalidChart {
                                offset: record.offset(),
                                reason: "Graph Dimensions longest row exceeds 3,999",
                            })?;
                        let rows = u8::try_from(u16_at(data, 10, record)?).ok().ok_or({
                            Error::InvalidChart {
                                offset: record.offset(),
                                reason: "Graph Dimensions row count exceeds 255",
                            }
                        })?;
                        chart_cache::Dims::Graph(chart_cache::GraphDims::new(longest, rows).ok_or(
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
                let index = chart_cache::Index::from_raw(u16_at(data, 0, record)?).ok_or(
                    Error::InvalidChart {
                        offset: record.offset(),
                        reason: "SIIndex must identify values, categories, or bubbles",
                    },
                )?;
                if cache_sections[index.slot()] {
                    return invalid(record, "SIIndex section is duplicated");
                }
                cache_sections[index.slot()] = true;
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
                if (context.kind() == ChartKind::Graph && !graph_cache_open)
                    || (context.kind() == ChartKind::Excel && (depth != 0 || !dimensions_seen))
                {
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
                            chart_cache::Ifmt::new(u16_at(data, 5, record)?),
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
                            chart_cache::Xf::new(u16_at(data, 4, record)?),
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
                if (context.kind() == ChartKind::Graph && !graph_cache_open)
                    || (context.kind() == ChartKind::Excel && (depth != 0 || !dimensions_seen))
                {
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
                            chart_cache::Ifmt::new(u16_at(data, 5, record)?),
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
                            chart_cache::Xf::new(u16_at(data, 4, record)?),
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
                if (context.kind() == ChartKind::Graph && !graph_cache_open)
                    || (context.kind() == ChartKind::Excel && (depth != 0 || !dimensions_seen))
                {
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
                            chart_cache::Ifmt::new(u16_at(data, 5, record)?),
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
                            chart_cache::Xf::new(u16_at(data, 4, record)?),
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
                    1 => XlValue::Error(
                        chart_cache::Fault::from_raw(byte_at(data, 6, record)?).ok_or(
                            Error::InvalidChart {
                                offset: record.offset(),
                                reason: "BoolErr contains an unknown BIFF error code",
                            },
                        )?,
                    ),
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
                    chart_cache::Xf::new(u16_at(data, 4, record)?),
                    value,
                    record,
                )?;
                push(&mut chart.caches, value, "chart cache")?;
            },
            BOF => return invalid(record, "nested BOF in chart substream"),
            OBJECT_LINK => {
                if attached_label_active && attached_label_depth == Some(depth) {
                    exact(record, 6)?;
                    attached_label_link_object = Some(u16_at(data, 0, record)?);
                }
                add_raw(&mut chart, record, &mut unknown_bytes, limits)?;
            },
            CONTINUE | SERIES_LIST | CAT_SER_RANGE | DEFAULT_TEXT | FONT_X | FRAME | POS => {
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
    if datasheet_pending || datasheet_active || attached_label_pending || attached_label_active {
        return Err(Error::InvalidChart {
            offset: input.as_bytes().len(),
            reason: "Datasheet or attached-label collection is incomplete",
        });
    }
    if datasheet_seen
        && (!units_seen
            || !orient_seen
            || !country_seen
            || !font_x_seen
            || !exclude_rows_seen
            || !exclude_columns_seen
            || !windoc_seen
            || !main_window_seen
            || !window1_seen
            || !window2_seen
            || !max_status_seen
            || !selection_seen
            || !chart_seen
            || !chart_closed
            || !dimensions_seen)
    {
        return Err(Error::InvalidChart {
            offset: input.as_bytes().len(),
            reason: "Graph chart sheet is missing a mandatory Datasheet or chart record",
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
                u16::try_from(chart.parents.len())
                    .ok()
                    .ok_or(Error::SizeOverflow {
                        resource: "axis-parent count",
                    })?,
            )
    {
        return Err(Error::InvalidChart {
            offset: input.as_bytes().len(),
            reason: "AxesUsed does not match the AxisParent collection count",
        });
    }
    if context.kind() == ChartKind::Excel && next_cache_section != chart_cache::Index::ALL.len() {
        return Err(Error::InvalidChart {
            offset: input.as_bytes().len(),
            reason: "Excel SERIESDATA does not contain exactly three SIIndex sections",
        });
    }
    validate::validate(&chart, limits, strict_excel)?;
    Ok(chart)
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
