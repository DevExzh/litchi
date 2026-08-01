//! Lossless BIFF8 chart-sheet and embedded-chart metadata and mutation.
//!
//! Chart formulas and cached values are inert. This module never evaluates a
//! formula, opens an external workbook, refreshes a cache, or renders a chart.

use std::collections::{HashMap, HashSet};

use litchi_ole_common::object::{
    Editor as ObjectEditor, Format as ObjectFormat, Limits as ObjectLimits,
};

use super::{XlsError, XlsResult};

const BOF: u16 = 0x0809;
const EOF: u16 = 0x000a;
const CONTINUE: u16 = 0x003c;
const OBJ: u16 = 0x005d;
const BOUNDSHEET: u16 = 0x0085;
const WINDOW1: u16 = 0x003d;
const RR_TAB_ID: u16 = 0x013d;
const SUP_BOOK: u16 = 0x01ae;
const EXTERN_SHEET: u16 = 0x0017;
const LBL: u16 = 0x0018;
const NUMBER: u16 = 0x0203;
const LABEL: u16 = 0x0204;
const CHART: u16 = 0x1002;
const SERIES: u16 = 0x1003;
const DATA_FORMAT: u16 = 0x1006;
const LINE_FORMAT: u16 = 0x1007;
const MARKER_FORMAT: u16 = 0x1009;
const AREA_FORMAT: u16 = 0x100a;
const PIE_FORMAT: u16 = 0x100b;
const SERIES_TEXT: u16 = 0x100d;
const CHART_FORMAT: u16 = 0x1014;
const LEGEND: u16 = 0x1015;
const SERIES_LIST: u16 = 0x1016;
const BAR: u16 = 0x1017;
const LINE: u16 = 0x1018;
const PIE: u16 = 0x1019;
const AREA: u16 = 0x101a;
const SCATTER: u16 = 0x101b;
const CRT_LINE: u16 = 0x101c;
const AXIS: u16 = 0x101d;
const TICK: u16 = 0x101e;
const VALUE_RANGE: u16 = 0x101f;
const CAT_SER_RANGE: u16 = 0x1020;
const AXIS_LINE: u16 = 0x1021;
const DEFAULT_TEXT: u16 = 0x1024;
const TEXT: u16 = 0x1025;
const FONT_X: u16 = 0x1026;
const OBJECT_LINK: u16 = 0x1027;
const FRAME: u16 = 0x1032;
const BEGIN: u16 = 0x1033;
const END: u16 = 0x1034;
const PLOT_AREA: u16 = 0x1035;
const DROP_BAR: u16 = 0x103d;
const RADAR: u16 = 0x103e;
const SURFACE: u16 = 0x103f;
const RADAR_AREA: u16 = 0x1040;
const AXIS_PARENT: u16 = 0x1041;
const SHT_PROPS: u16 = 0x1044;
const SER_TO_CRT: u16 = 0x1045;
const AXES_USED: u16 = 0x1046;
const DATA_LAB_EXT: u16 = 0x086a;
const DATA_LAB_EXT_CONTENTS: u16 = 0x086b;
const PLOT_GROWTH: u16 = 0x1064;
const SI_INDEX: u16 = 0x1065;
const MAX_BIFF_DATA: usize = 8_224;

/// Hard resource bounds for chart discovery and authoring.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct XlsChartLimits {
    pub max_workbook_bytes: usize,
    pub max_charts: usize,
    pub max_records_per_chart: usize,
    pub max_series: usize,
    pub max_groups: usize,
    pub max_axes: usize,
    pub max_formula_bytes: usize,
    pub max_cached_values: usize,
    pub max_unknown_bytes: usize,
}

impl Default for XlsChartLimits {
    fn default() -> Self {
        Self {
            max_workbook_bytes: 128 * 1024 * 1024,
            max_charts: 512,
            max_records_per_chart: 8_192,
            max_series: 255,
            max_groups: 10,
            max_axes: 6,
            max_formula_bytes: MAX_BIFF_DATA - 8,
            max_cached_values: 32_000,
            max_unknown_bytes: 16 * 1024 * 1024,
        }
    }
}

/// Stable location of one chart in the current workbook revision.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub enum XlsChartLocation {
    ChartSheet { sheet_index: usize },
    Embedded { sheet_index: usize, object_id: u16 },
}

impl XlsChartLocation {
    pub fn sheet_index(&self) -> usize {
        match self {
            Self::ChartSheet { sheet_index } | Self::Embedded { sheet_index, .. } => *sheet_index,
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct XlsChartRawRecord {
    pub record_type: u16,
    pub data: Vec<u8>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum XlsChartDataKind {
    Numeric,
    Text,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct XlsChartCellReference {
    pub extern_sheet_index: u16,
    pub first_row: u16,
    pub last_row: u16,
    pub first_column: u16,
    pub last_column: u16,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct XlsChartDataLink {
    /// 0=name, 1=values/X, 2=categories/Y, 3=bubble size.
    pub role: u8,
    /// 0=automatic, 1=literal/formula, 2=cell range.
    pub source_type: u8,
    pub unlinked_number_format: bool,
    pub number_format: u16,
    pub formula_tokens: Vec<u8>,
    pub references: Vec<XlsChartCellReference>,
}

#[derive(Clone, Debug, PartialEq)]
pub enum XlsChartCachedValue {
    Number(f64),
    Text(String),
    Blank,
}

#[derive(Clone, Debug, PartialEq)]
pub struct XlsChartCacheEntry {
    pub cache_index: u16,
    pub row: u16,
    pub column: u16,
    pub value: XlsChartCachedValue,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct XlsChartSeries {
    pub category_kind: XlsChartDataKind,
    pub category_count: u16,
    pub value_count: u16,
    pub bubble_count: u16,
    pub chart_group: u16,
    pub name: Option<String>,
    pub links: Vec<XlsChartDataLink>,
}

impl Default for XlsChartSeries {
    fn default() -> Self {
        Self {
            category_kind: XlsChartDataKind::Text,
            category_count: 0,
            value_count: 0,
            bubble_count: 0,
            chart_group: 0,
            name: None,
            links: Vec::new(),
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum XlsChartGroupKind {
    Line {
        flags: u16,
    },
    Bar {
        overlap: i16,
        gap: u16,
        flags: u16,
    },
    Area {
        flags: u16,
    },
    Pie {
        rotation: u16,
        hole_size: u16,
        flags: u16,
    },
    Scatter {
        bubble_size_percent: u16,
        bubble_size_type: u16,
        flags: u16,
    },
    Radar {
        filled: bool,
        flags: u16,
    },
    Surface {
        flags: u16,
    },
    Stock {
        flags: u16,
    },
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct XlsChartGroup {
    pub order: u16,
    pub vary_colors: bool,
    pub kind: XlsChartGroupKind,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum XlsChartType {
    Empty,
    Single(XlsChartGroupKind),
    Stock,
    Combo(Vec<XlsChartGroupKind>),
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum XlsChartAxisKind {
    CategoryOrHorizontal,
    ValueOrVertical,
    Series,
}

#[derive(Clone, Debug, PartialEq)]
pub struct XlsChartAxisScale {
    pub minimum: f64,
    pub maximum: f64,
    pub major: f64,
    pub minor: f64,
    pub crossing: f64,
    pub flags: u16,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct XlsChartTick {
    pub major: u8,
    pub minor: u8,
    pub label_position: u8,
    pub background: u8,
    pub color: [u8; 4],
    pub flags: u16,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum XlsChartAxisLineKind {
    Axis,
    MajorGridlines,
    MinorGridlines,
    WallsOrFloor,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct XlsChartLineFormat {
    pub color: [u8; 4],
    pub pattern: u16,
    pub weight: i16,
    pub flags: u16,
    pub color_index: u16,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct XlsChartAxisLine {
    pub kind: XlsChartAxisLineKind,
    pub format: Option<XlsChartLineFormat>,
}

#[derive(Clone, Debug, PartialEq)]
pub struct XlsChartAxis {
    pub kind: XlsChartAxisKind,
    pub scale: Option<XlsChartAxisScale>,
    pub tick: Option<XlsChartTick>,
    pub lines: Vec<XlsChartAxisLine>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct XlsChartLegend {
    pub x: i32,
    pub y: i32,
    pub width: i32,
    pub height: i32,
    pub position: u8,
    pub spacing: u8,
    pub flags: u16,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct XlsChartAreaFormat {
    pub foreground: [u8; 4],
    pub background: [u8; 4],
    pub pattern: u16,
    pub flags: u16,
    pub foreground_index: u16,
    pub background_index: u16,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum XlsChartFormatting {
    Line(XlsChartLineFormat),
    Area(XlsChartAreaFormat),
    Marker { data: Vec<u8> },
    Data { point: u16, series: u16, flags: u16 },
    Pie { explosion: u16 },
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct XlsChartDataLabel {
    pub record_type: u16,
    pub data: Vec<u8>,
}

#[derive(Clone, Debug, PartialEq)]
pub struct XlsChart {
    pub x: i32,
    pub y: i32,
    pub width: i32,
    pub height: i32,
    pub sheet_properties: u32,
    pub plot_area_present: bool,
    pub title: Option<String>,
    pub series: Vec<XlsChartSeries>,
    pub groups: Vec<XlsChartGroup>,
    pub axes: Vec<XlsChartAxis>,
    pub legend: Option<XlsChartLegend>,
    pub cached_values: Vec<XlsChartCacheEntry>,
    pub formatting: Vec<XlsChartFormatting>,
    pub data_labels: Vec<XlsChartDataLabel>,
    /// Records not interpreted by this implementation, retained byte-for-byte.
    pub unknown_records: Vec<XlsChartRawRecord>,
}

impl Default for XlsChart {
    fn default() -> Self {
        Self {
            x: 0,
            y: 0,
            width: 4000 << 16,
            height: 3000 << 16,
            sheet_properties: 0x0000_0002,
            plot_area_present: true,
            title: None,
            series: Vec::new(),
            groups: vec![XlsChartGroup {
                order: 0,
                vary_colors: false,
                kind: XlsChartGroupKind::Line { flags: 0 },
            }],
            axes: Vec::new(),
            legend: None,
            cached_values: Vec::new(),
            formatting: Vec::new(),
            data_labels: Vec::new(),
            unknown_records: Vec::new(),
        }
    }
}

impl XlsChart {
    pub fn chart_type(&self) -> XlsChartType {
        if self.groups.is_empty() {
            XlsChartType::Empty
        } else if self.groups.len() > 1 {
            XlsChartType::Combo(self.groups.iter().map(|v| v.kind.clone()).collect())
        } else if matches!(self.groups[0].kind, XlsChartGroupKind::Stock { .. }) {
            XlsChartType::Stock
        } else {
            XlsChartType::Single(self.groups[0].kind.clone())
        }
    }

    pub fn validate(&self, sheet_count: usize, limits: XlsChartLimits) -> XlsResult<()> {
        validate_limits(limits)?;
        if self.series.len() > limits.max_series
            || self.groups.len() > limits.max_groups
            || self.axes.len() > limits.max_axes
            || self.cached_values.len() > limits.max_cached_values
        {
            return invalid(CHART, "chart resource limit exceeded");
        }
        if self.groups.len() > 10 {
            return invalid(CHART_FORMAT, "BIFF8 permits at most ten chart groups");
        }
        let mut orders = HashSet::new();
        for group in &self.groups {
            if group.order > 9 || !orders.insert(group.order) {
                return invalid(
                    CHART_FORMAT,
                    "chart group order is duplicated or exceeds nine",
                );
            }
            match group.kind {
                XlsChartGroupKind::Area { flags } | XlsChartGroupKind::Line { flags }
                    if flags & !7 != 0 =>
                {
                    return invalid(CHART_FORMAT, "area/line chart uses reserved flags");
                },
                XlsChartGroupKind::Bar { flags, .. } if flags & !0xf != 0 => {
                    return invalid(BAR, "bar chart uses reserved flags");
                },
                XlsChartGroupKind::Pie {
                    rotation,
                    hole_size,
                    flags,
                } if rotation > 360 || hole_size > 90 || flags & !3 != 0 => {
                    return invalid(PIE, "pie/doughnut settings are out of range");
                },
                XlsChartGroupKind::Radar { flags, .. } | XlsChartGroupKind::Surface { flags }
                    if flags & !3 != 0 =>
                {
                    return invalid(CHART_FORMAT, "radar/surface chart uses reserved flags");
                },
                _ => {},
            }
        }
        for series in &self.series {
            if usize::from(series.chart_group) >= self.groups.len() {
                return invalid(SER_TO_CRT, "series references a missing chart group");
            }
            for link in &series.links {
                validate_link(link, sheet_count, limits)?;
            }
        }
        for axis in &self.axes {
            if let Some(scale) = &axis.scale
                && (![
                    scale.minimum,
                    scale.maximum,
                    scale.major,
                    scale.minor,
                    scale.crossing,
                ]
                .into_iter()
                .all(f64::is_finite)
                    || scale.maximum < scale.minimum
                    || scale.major < 0.0
                    || scale.minor < 0.0)
            {
                return invalid(VALUE_RANGE, "axis scale is not finite or ordered");
            }
            let mut line_order = None;
            for line in &axis.lines {
                let value = match line.kind {
                    XlsChartAxisLineKind::Axis => 0,
                    XlsChartAxisLineKind::MajorGridlines => 1,
                    XlsChartAxisLineKind::MinorGridlines => 2,
                    XlsChartAxisLineKind::WallsOrFloor => 3,
                };
                if line_order.is_some_and(|previous| value <= previous) {
                    return invalid(AXIS_LINE, "axis line records are duplicated or not ordered");
                }
                line_order = Some(value);
            }
        }
        let unknown = self
            .unknown_records
            .iter()
            .try_fold(0usize, |sum, value| sum.checked_add(value.data.len()))
            .ok_or_else(|| XlsError::InvalidData("chart unknown-record size overflow".into()))?;
        if unknown > limits.max_unknown_bytes {
            return invalid(CHART, "opaque chart data exceeds limit");
        }
        for record in &self.unknown_records {
            if record.data.len() > MAX_BIFF_DATA {
                return invalid(
                    record.record_type,
                    "opaque BIFF record exceeds maximum length",
                );
            }
        }
        Ok(())
    }
}

#[derive(Clone, Debug, PartialEq)]
pub struct XlsChartEntry {
    pub location: XlsChartLocation,
    pub chart: XlsChart,
}

#[derive(Clone)]
struct StoredChart {
    entry: XlsChartEntry,
    start: usize,
    end: usize,
    object: Option<(usize, usize)>,
}

/// Transactional editor for existing BIFF8 chart substreams.
#[derive(Clone)]
pub struct XlsChartEditor {
    package: ObjectEditor,
    workbook_path: Vec<String>,
    workbook: Vec<u8>,
    limits: XlsChartLimits,
    charts: Vec<StoredChart>,
}

impl XlsChartEditor {
    pub fn open(bytes: Vec<u8>, limits: XlsChartLimits) -> XlsResult<Self> {
        validate_limits(limits)?;
        let package = ObjectEditor::open(bytes, ObjectFormat::Xls, ObjectLimits::default())?;
        let workbook_path = [vec!["Workbook".into()], vec!["Book".into()]]
            .into_iter()
            .find(|path| package.stream(path).is_some())
            .ok_or_else(|| XlsError::InvalidData("Workbook stream not found".into()))?;
        let workbook = package
            .stream(&workbook_path)
            .ok_or_else(|| XlsError::InvalidData("selected Workbook stream disappeared".into()))?
            .to_vec();
        if workbook.len() > limits.max_workbook_bytes {
            return invalid(CHART, "Workbook stream exceeds chart editor limit");
        }
        let charts = parse_workbook_charts(&workbook, limits)?;
        Ok(Self {
            package,
            workbook_path,
            workbook,
            limits,
            charts,
        })
    }

    pub fn charts(&self) -> Vec<XlsChartEntry> {
        self.charts
            .iter()
            .map(|value| value.entry.clone())
            .collect()
    }

    /// Consume the editor and return the parsed chart inventory without persisting.
    pub fn into_charts(self) -> Vec<XlsChartEntry> {
        self.charts.into_iter().map(|value| value.entry).collect()
    }

    pub fn find(&self, location: &XlsChartLocation) -> Option<&XlsChart> {
        self.charts
            .iter()
            .find(|value| &value.entry.location == location)
            .map(|value| &value.entry.chart)
    }

    /// Add an Obj-linked embedded chart at `index` among charts on the worksheet.
    pub fn add(
        &mut self,
        sheet_index: usize,
        object_id: u16,
        index: usize,
        chart: XlsChart,
    ) -> XlsResult<()> {
        if object_id == 0 || self.charts.iter().any(|value| matches!(value.entry.location, XlsChartLocation::Embedded { object_id: id, .. } if id == object_id)) {
            return invalid(OBJ, "embedded chart object ID is zero or duplicated");
        }
        let sheet_count = bindings(&self.workbook)?.1.len();
        chart.validate(sheet_count, self.limits)?;
        let original = self.charts.clone();
        let mut desired = original.iter().map(|v| v.entry.clone()).collect::<Vec<_>>();
        let positions = desired
            .iter()
            .enumerate()
            .filter_map(|(i, value)| (value.location.sheet_index() == sheet_index).then_some(i))
            .collect::<Vec<_>>();
        if index > positions.len() {
            return invalid(CHART, "embedded chart insertion index is out of range");
        }
        let insert = positions
            .get(index)
            .copied()
            .unwrap_or_else(|| positions.last().map_or(desired.len(), |v| v + 1));
        desired.insert(
            insert,
            XlsChartEntry {
                location: XlsChartLocation::Embedded {
                    sheet_index,
                    object_id,
                },
                chart,
            },
        );
        self.commit(&original, desired)
    }

    /// Insert a new chart-sheet tab and its chart BOF..EOF substream.
    pub fn add_chart_sheet(
        &mut self,
        index: usize,
        name: impl Into<String>,
        chart: XlsChart,
    ) -> XlsResult<()> {
        let name = name.into();
        let (_, sheets) = bindings(&self.workbook)?;
        if index > sheets.len() {
            return invalid(BOUNDSHEET, "chart-sheet insertion index is out of range");
        }
        validate_sheet_name(&name)?;
        chart.validate(sheets.len() + 1, self.limits)?;
        let globals_end = sheets
            .iter()
            .map(|value| value.start)
            .min()
            .map_or(self.workbook.len(), |value| value);
        for value in ranges(&self.workbook[..globals_end])? {
            if value.kind == BOUNDSHEET
                && bound_sheet_name(&self.workbook[value.body_start..value.body_end])?
                    .eq_ignore_ascii_case(&name)
            {
                return invalid(BOUNDSHEET, "chart-sheet name duplicates an existing tab");
            }
        }
        let bound = bound_sheet_body(&name, 2)?;
        let stream = serialize_chart(&chart, sheets.len() + 1, self.limits)?;
        let order = (0..sheets.len()).map(Some).collect::<Vec<_>>();
        let workbook =
            rewrite_sheet_directory(&self.workbook, &order, Some((index, bound, stream)))?;
        self.install_workbook(workbook)
    }

    /// Remove a chart-sheet tab. References to that tab cause atomic failure.
    pub fn remove_chart_sheet(&mut self, sheet_index: usize) -> XlsResult<XlsChart> {
        let (_, sheets) = bindings(&self.workbook)?;
        let sheet = sheets
            .iter()
            .find(|value| value.index == sheet_index)
            .ok_or_else(|| invalid_error(BOUNDSHEET, "sheet index was not found"))?;
        if sheet.kind != 2 {
            return invalid(BOUNDSHEET, "selected tab is not a chart sheet");
        }
        let chart = self
            .find(&XlsChartLocation::ChartSheet { sheet_index })
            .cloned()
            .ok_or_else(|| invalid_error(CHART, "chart sheet has no chart"))?;
        let order = (0..sheets.len())
            .filter(|value| *value != sheet_index)
            .map(Some)
            .collect::<Vec<_>>();
        let workbook = rewrite_sheet_directory(&self.workbook, &order, None)?;
        self.install_workbook(workbook)?;
        Ok(chart)
    }

    /// Reorder all workbook tabs by their previous zero-based indexes.
    pub fn reorder_sheets(&mut self, order: &[usize]) -> XlsResult<()> {
        let count = bindings(&self.workbook)?.1.len();
        if order.len() != count {
            return invalid(BOUNDSHEET, "sheet reorder must contain every tab");
        }
        let mut seen = HashSet::new();
        if order
            .iter()
            .any(|value| *value >= count || !seen.insert(*value))
        {
            return invalid(
                BOUNDSHEET,
                "sheet reorder contains an invalid or repeated tab",
            );
        }
        let workbook = rewrite_sheet_directory(
            &self.workbook,
            &order.iter().copied().map(Some).collect::<Vec<_>>(),
            None,
        )?;
        self.install_workbook(workbook)
    }

    pub fn update(&mut self, location: &XlsChartLocation, chart: XlsChart) -> XlsResult<()> {
        self.replace(location, chart)
    }

    pub fn replace(&mut self, location: &XlsChartLocation, chart: XlsChart) -> XlsResult<()> {
        let sheet_count = bindings(&self.workbook)?.1.len();
        chart.validate(sheet_count, self.limits)?;
        let original = self.charts.clone();
        let mut desired = original.iter().map(|v| v.entry.clone()).collect::<Vec<_>>();
        let value = desired
            .iter_mut()
            .find(|value| &value.location == location)
            .ok_or_else(|| invalid_error(CHART, "chart location was not found"))?;
        value.chart = chart;
        self.commit(&original, desired)
    }

    pub fn remove(&mut self, location: &XlsChartLocation) -> XlsResult<XlsChart> {
        if let XlsChartLocation::ChartSheet { sheet_index } = location {
            return self.remove_chart_sheet(*sheet_index);
        }
        let original = self.charts.clone();
        let mut desired = original.iter().map(|v| v.entry.clone()).collect::<Vec<_>>();
        let index = desired
            .iter()
            .position(|value| &value.location == location)
            .ok_or_else(|| invalid_error(CHART, "chart location was not found"))?;
        let removed = desired.remove(index).chart;
        self.commit(&original, desired)?;
        Ok(removed)
    }

    /// Reorder every embedded chart on one worksheet by Obj identifier.
    pub fn reorder(&mut self, sheet_index: usize, object_ids: &[u16]) -> XlsResult<()> {
        let original = self.charts.clone();
        let mut desired = original.iter().map(|v| v.entry.clone()).collect::<Vec<_>>();
        let slots = desired
            .iter()
            .enumerate()
            .filter_map(|(i, value)| match value.location {
                XlsChartLocation::Embedded {
                    sheet_index: sheet, ..
                } if sheet == sheet_index => Some(i),
                _ => None,
            })
            .collect::<Vec<_>>();
        if slots.len() != object_ids.len() {
            return invalid(
                CHART,
                "reorder must include every embedded chart on the worksheet",
            );
        }
        let mut values = slots
            .iter()
            .map(|index| desired[*index].clone())
            .collect::<Vec<_>>();
        let mut ordered = Vec::new();
        for id in object_ids {
            let index = values.iter().position(|value| matches!(value.location, XlsChartLocation::Embedded { object_id, .. } if object_id == *id))
            .ok_or_else(|| invalid_error(CHART, "reorder contains an unknown or repeated object ID"))?;
            ordered.push(values.remove(index));
        }
        for (slot, value) in slots.into_iter().zip(ordered) {
            desired[slot] = value;
        }
        self.commit(&original, desired)
    }

    pub fn finish(self) -> XlsResult<Vec<u8>> {
        self.package.finish().map_err(Into::into)
    }

    fn commit(&mut self, original: &[StoredChart], desired: Vec<XlsChartEntry>) -> XlsResult<()> {
        if desired.len() > self.limits.max_charts {
            return invalid(CHART, "chart count exceeds limit");
        }
        let workbook = rewrite_workbook_charts(&self.workbook, original, &desired, self.limits)?;
        let reparsed = parse_workbook_charts(&workbook, self.limits)?;
        let actual = reparsed.iter().map(|v| v.entry.clone()).collect::<Vec<_>>();
        if actual != desired {
            return invalid(
                CHART,
                "rewritten chart substreams failed typed round-trip validation",
            );
        }
        let mut package = self.package.clone();
        package.put_stream(&self.workbook_path, workbook.clone())?;
        self.package = package;
        self.workbook = workbook;
        self.charts = reparsed;
        Ok(())
    }

    fn install_workbook(&mut self, workbook: Vec<u8>) -> XlsResult<()> {
        if workbook.len() > self.limits.max_workbook_bytes {
            return invalid(CHART, "rewritten Workbook exceeds limit");
        }
        let charts = parse_workbook_charts(&workbook, self.limits)?;
        let mut package = self.package.clone();
        package.put_stream(&self.workbook_path, workbook.clone())?;
        self.package = package;
        self.workbook = workbook;
        self.charts = charts;
        Ok(())
    }
}

/// The Obj identifier assigned to the chart embedded in a generated workbook.
const GENERATED_CHART_OBJECT_ID: u16 = 1;
/// BIFF record type marking the workbook-globals substream.
const BOF_WORKBOOK_GLOBALS: u16 = 0x0005;
/// BIFF record type marking a worksheet substream.
const BOF_WORKSHEET: u16 = 0x0010;
/// Sheet name of the single worksheet hosting a generated embedded chart.
const GENERATED_SHEET_NAME: &str = "Sheet1";

/// Build a standalone BIFF8 workbook compound file holding `chart` embedded
/// on a single otherwise-empty worksheet.
///
/// The generated worksheet carries no cell records; the chart is fully
/// described by its own cached values and inert data links. No formula is
/// evaluated and no external workbook is referenced or opened.
pub fn build_chart_workbook(chart: XlsChart, limits: XlsChartLimits) -> XlsResult<Vec<u8>> {
    validate_limits(limits)?;
    let mut package = crate::OleWriter::new();
    package.create_stream(&["Workbook"], &minimal_workbook_stream()?)?;
    let mut bytes = std::io::Cursor::new(Vec::new());
    package.write_to(&mut bytes)?;
    let mut editor = XlsChartEditor::open(bytes.into_inner(), limits)?;
    editor.add(0, GENERATED_CHART_OBJECT_ID, 0, chart)?;
    editor.finish()
}

/// A minimal one-worksheet BIFF8 `Workbook` stream accepted by the chart
/// editor: workbook globals with a single `BoundSheet` directory entry
/// followed by an empty worksheet substream.
fn minimal_workbook_stream() -> XlsResult<Vec<u8>> {
    let mut output = record(BOF, &bof_body(BOF_WORKBOOK_GLOBALS))?;
    let bound_offset_position = output.len() + 4;
    output.extend(record(
        BOUNDSHEET,
        &bound_sheet_body(GENERATED_SHEET_NAME, 0)?,
    )?);
    output.extend(record(EOF, &[])?);
    let sheet_offset = u32::try_from(output.len())
        .map_err(|_| XlsError::InvalidData("BoundSheet offset exceeds u32".into()))?;
    output[bound_offset_position..bound_offset_position + 4]
        .copy_from_slice(&sheet_offset.to_le_bytes());
    output.extend(record(BOF, &bof_body(BOF_WORKSHEET))?);
    output.extend(record(EOF, &[])?);
    Ok(output)
}

fn rewrite_sheet_directory(
    input: &[u8],
    order: &[Option<usize>],
    insert: Option<(usize, Vec<u8>, Vec<u8>)>,
) -> XlsResult<Vec<u8>> {
    let (_, sheets) = bindings(input)?;
    let old_count = sheets.len();
    let mut logical = vec![None; old_count];
    for sheet in &sheets {
        logical[sheet.index] = Some((sheet.kind, input[sheet.start..sheet.end].to_vec()));
    }
    let globals_end = sheets
        .iter()
        .map(|value| value.start)
        .min()
        .ok_or_else(|| invalid_error(BOUNDSHEET, "workbook has no sheet substreams"))?;
    let global_ranges = ranges(&input[..globals_end])?;
    let old_bounds = global_ranges
        .iter()
        .filter(|value| value.kind == BOUNDSHEET)
        .map(|value| input[value.body_start..value.body_end].to_vec())
        .collect::<Vec<_>>();
    if old_bounds.len() != old_count {
        return invalid(BOUNDSHEET, "BoundSheet directory count mismatch");
    }
    let mut tabs = Vec::new();
    for old in order {
        let index = old
            .ok_or_else(|| invalid_error(BOUNDSHEET, "unexpected empty sheet permutation slot"))?;
        let (_, stream) = logical
            .get(index)
            .and_then(|value| value.clone())
            .ok_or_else(|| invalid_error(BOUNDSHEET, "sheet permutation target is missing"))?;
        tabs.push((Some(index), old_bounds[index].clone(), stream));
    }
    if let Some((index, bound, stream)) = insert {
        if index > tabs.len() {
            return invalid(BOUNDSHEET, "inserted chart-sheet index is out of range");
        }
        tabs.insert(index, (None, bound, stream));
    }
    if tabs.is_empty() {
        return invalid(BOUNDSHEET, "workbook must retain at least one sheet");
    }
    let mut old_to_new = vec![None; old_count];
    for (new, (old, _, _)) in tabs.iter().enumerate() {
        if let Some(old) = old {
            old_to_new[*old] = Some(new);
        }
    }
    let insert_index = tabs.iter().position(|value| value.0.is_none());
    let globals = rewrite_chart_globals(&input[..globals_end], &tabs, &old_to_new, insert_index)?;
    let mut output = globals.bytes;
    let mut offsets = Vec::with_capacity(tabs.len());
    for (_, _, stream) in &tabs {
        offsets.push(output.len());
        output.extend_from_slice(stream);
    }
    for (position, offset) in globals.bound_positions.into_iter().zip(offsets) {
        output[position..position + 4].copy_from_slice(
            &u32::try_from(offset)
                .map_err(|_| XlsError::InvalidData("BoundSheet offset exceeds u32".into()))?
                .to_le_bytes(),
        );
    }
    Ok(output)
}

struct RewrittenGlobals {
    bytes: Vec<u8>,
    bound_positions: Vec<usize>,
}

fn rewrite_chart_globals(
    input: &[u8],
    tabs: &[(Option<usize>, Vec<u8>, Vec<u8>)],
    old_to_new: &[Option<usize>],
    insert_index: Option<usize>,
) -> XlsResult<RewrittenGlobals> {
    let records = ranges(input)?;
    let mut output = Vec::new();
    let mut bound_positions = Vec::new();
    let mut bounds_written = false;
    let internal_books = internal_sup_books(input, &records)?;
    let rr_ids = records
        .iter()
        .find(|value| value.kind == RR_TAB_ID)
        .map(|value| parse_rr_tab_ids(&input[value.body_start..value.body_end], old_to_new.len()))
        .transpose()?;
    for value in records {
        let data = &input[value.body_start..value.body_end];
        if value.kind == BOUNDSHEET {
            if bounds_written {
                continue;
            }
            bounds_written = true;
            for (_, body, _) in tabs {
                let mut body = body.clone();
                body[..4].fill(0);
                bound_positions.push(output.len() + 4);
                output.extend(record(BOUNDSHEET, &body)?);
            }
            continue;
        }
        let rewritten = match value.kind {
            WINDOW1 => remap_window1(data, old_to_new)?,
            RR_TAB_ID => write_rr_tab_ids(
                rr_ids.as_ref().ok_or_else(|| {
                    invalid_error(RR_TAB_ID, "RRTabId record inventory is missing")
                })?,
                tabs,
            )?,
            SUP_BOOK if data.len() == 4 && u16_at(data, 2)? == 0x0401 => {
                let mut value = data.to_vec();
                value[..2].copy_from_slice(
                    &u16::try_from(tabs.len())
                        .map_err(|_| XlsError::InvalidData("sheet count exceeds u16".into()))?
                        .to_le_bytes(),
                );
                value
            },
            EXTERN_SHEET => remap_extern_sheet(data, &internal_books, old_to_new, insert_index)?,
            LBL => remap_lbl(data, old_to_new)?,
            _ => data.to_vec(),
        };
        output.extend(record(value.kind, &rewritten)?);
    }
    if !bounds_written {
        return invalid(BOUNDSHEET, "workbook globals contain no BoundSheet records");
    }
    Ok(RewrittenGlobals {
        bytes: output,
        bound_positions,
    })
}

fn internal_sup_books(input: &[u8], records: &[Range]) -> XlsResult<HashSet<u16>> {
    let mut result = HashSet::new();
    let mut ordinal = 0u16;
    for value in records {
        if value.kind != SUP_BOOK {
            continue;
        }
        let data = &input[value.body_start..value.body_end];
        if data.len() == 4 && u16_at(data, 2)? == 0x0401 {
            result.insert(ordinal);
        }
        ordinal = ordinal
            .checked_add(1)
            .ok_or_else(|| XlsError::InvalidData("SupBook count overflow".into()))?;
    }
    Ok(result)
}

fn remap_extern_sheet(
    data: &[u8],
    internal: &HashSet<u16>,
    old_to_new: &[Option<usize>],
    insert_index: Option<usize>,
) -> XlsResult<Vec<u8>> {
    if data.len() < 2 {
        return invalid(EXTERN_SHEET, "ExternSheet is truncated");
    }
    let count = usize::from(u16_at(data, 0)?);
    if data.len() != 2 + count * 6 {
        return invalid(EXTERN_SHEET, "ExternSheet count does not match payload");
    }
    let mut output = data.to_vec();
    for index in 0..count {
        let offset = 2 + index * 6;
        if !internal.contains(&u16_at(data, offset)?) {
            continue;
        }
        let first = u16_at(data, offset + 2)?;
        let last = u16_at(data, offset + 4)?;
        if matches!(first, 0xfffe | 0xffff) || matches!(last, 0xfffe | 0xffff) {
            continue;
        }
        let first = usize::from(first);
        let last = usize::from(last);
        if first > last || last >= old_to_new.len() {
            return invalid(EXTERN_SHEET, "internal ExternSheet range is invalid");
        }
        if insert_index.is_some_and(|insert| first < insert && insert <= last) {
            return invalid(
                EXTERN_SHEET,
                "cannot insert a sheet inside an existing 3-D formula range",
            );
        }
        let mapped = (first..=last)
            .map(|old| {
                old_to_new[old].ok_or_else(|| {
                    invalid_error(
                        EXTERN_SHEET,
                        "cannot remove a sheet referenced by a formula",
                    )
                })
            })
            .collect::<XlsResult<Vec<_>>>()?;
        let minimum = *mapped
            .iter()
            .min()
            .ok_or_else(|| invalid_error(EXTERN_SHEET, "empty 3-D formula range"))?;
        let maximum = *mapped
            .iter()
            .max()
            .ok_or_else(|| invalid_error(EXTERN_SHEET, "empty 3-D formula range"))?;
        if maximum - minimum + 1 != mapped.len() {
            return invalid(
                EXTERN_SHEET,
                "sheet reorder would make a 3-D formula range noncontiguous",
            );
        }
        output[offset + 2..offset + 4].copy_from_slice(&(minimum as u16).to_le_bytes());
        output[offset + 4..offset + 6].copy_from_slice(&(maximum as u16).to_le_bytes());
    }
    Ok(output)
}

fn remap_lbl(data: &[u8], old_to_new: &[Option<usize>]) -> XlsResult<Vec<u8>> {
    if data.len() < 10 {
        return invalid(LBL, "Lbl is truncated");
    }
    let scope = usize::from(u16_at(data, 8)?);
    if scope == 0 {
        return Ok(data.to_vec());
    }
    let old = scope - 1;
    let new =
        old_to_new.get(old).copied().flatten().ok_or_else(|| {
            invalid_error(LBL, "cannot remove a sheet owning a scoped defined name")
        })?;
    let mut output = data.to_vec();
    output[8..10].copy_from_slice(
        &u16::try_from(new + 1)
            .map_err(|_| XlsError::InvalidData("Lbl sheet scope exceeds u16".into()))?
            .to_le_bytes(),
    );
    Ok(output)
}

fn remap_window1(data: &[u8], old_to_new: &[Option<usize>]) -> XlsResult<Vec<u8>> {
    if data.len() != 18 {
        return invalid(WINDOW1, "Window1 must contain 18 bytes");
    }
    let mut output = data.to_vec();
    for offset in [10usize, 12] {
        let old = usize::from(u16_at(data, offset)?);
        let new = old_to_new
            .get(old)
            .copied()
            .flatten()
            .unwrap_or_else(|| old_to_new.iter().flatten().copied().min().unwrap_or(0));
        output[offset..offset + 2].copy_from_slice(&(new as u16).to_le_bytes());
    }
    let selected = usize::from(u16_at(data, 14)?).clamp(1, old_to_new.iter().flatten().count());
    output[14..16].copy_from_slice(&(selected as u16).to_le_bytes());
    Ok(output)
}

fn parse_rr_tab_ids(data: &[u8], count: usize) -> XlsResult<Vec<u16>> {
    if data.len() != count * 2 {
        return invalid(RR_TAB_ID, "RRTabId count does not match BoundSheet count");
    }
    (0..count).map(|index| u16_at(data, index * 2)).collect()
}
fn write_rr_tab_ids(old: &[u16], tabs: &[(Option<usize>, Vec<u8>, Vec<u8>)]) -> XlsResult<Vec<u8>> {
    let next = old
        .iter()
        .copied()
        .max()
        .unwrap_or(0)
        .checked_add(1)
        .ok_or_else(|| XlsError::InvalidData("RRTabId identifier overflow".into()))?;
    let mut output = Vec::new();
    for (old_index, _, _) in tabs {
        output.extend(
            old_index
                .map(|index| old[index])
                .unwrap_or(next)
                .to_le_bytes(),
        );
    }
    Ok(output)
}
fn validate_sheet_name(name: &str) -> XlsResult<()> {
    let count = name.encode_utf16().count();
    if !(1..=31).contains(&count)
        || name.chars().any(|value| {
            matches!(
                value,
                '\0' | '\u{0003}' | ':' | '\\' | '*' | '?' | '/' | '[' | ']'
            )
        })
        || name.starts_with('\'')
        || name.ends_with('\'')
    {
        return invalid(BOUNDSHEET, "invalid chart-sheet name");
    }
    Ok(())
}
fn bound_sheet_body(name: &str, kind: u8) -> XlsResult<Vec<u8>> {
    validate_sheet_name(name)?;
    let units = name.encode_utf16().collect::<Vec<_>>();
    let wide = units.iter().any(|v| *v > 255);
    let mut output = vec![0; 6];
    output[5] = kind;
    output.push(units.len() as u8);
    output.push(u8::from(wide));
    if wide {
        for value in units {
            output.extend(value.to_le_bytes());
        }
    } else {
        output.extend(units.into_iter().map(|v| v as u8));
    }
    Ok(output)
}
fn bound_sheet_name(data: &[u8]) -> XlsResult<String> {
    if data.len() < 8 {
        return invalid(BOUNDSHEET, "BoundSheet is truncated");
    }
    parse_biff8_string(&data[6..]).map_err(|_| XlsError::InvalidRecord {
        record_type: BOUNDSHEET,
        message: "invalid BoundSheet name".into(),
    })
}

fn parse_workbook_charts(input: &[u8], limits: XlsChartLimits) -> XlsResult<Vec<StoredChart>> {
    let (_, sheets) = bindings(input)?;
    let mut output = Vec::new();
    for sheet in &sheets {
        if sheet.kind == 2 {
            let chart = parse_chart(&input[sheet.start..sheet.end], sheets.len(), limits)?;
            output.push(StoredChart {
                entry: XlsChartEntry {
                    location: XlsChartLocation::ChartSheet {
                        sheet_index: sheet.index,
                    },
                    chart,
                },
                start: sheet.start,
                end: sheet.end,
                object: None,
            });
            continue;
        }
        if sheet.kind != 0 {
            continue;
        }
        let records = ranges(&input[sheet.start..sheet.end])?;
        let mut chart_objects = Vec::new();
        let mut used = HashSet::new();
        let mut index = 0;
        while index < records.len() {
            let value = records[index];
            if value.kind == OBJ
                && let Some(id) = parse_chart_object(
                    &input[sheet.start + value.body_start..sheet.start + value.body_end],
                )?
            {
                chart_objects.push((id, sheet.start + value.start, sheet.start + value.end));
            }
            if value.kind == BOF
                && is_chart_bof(
                    &input[sheet.start + value.body_start..sheet.start + value.body_end],
                )
            {
                let end_index = (index + 1..records.len())
                    .find(|candidate| records[*candidate].kind == EOF)
                    .ok_or_else(|| invalid_error(BOF, "chart BOF has no EOF"))?;
                let end = sheet.start + records[end_index].end;
                let start = sheet.start + value.start;
                let object = chart_objects
                    .iter()
                    .rev()
                    .find(|(id, _, object_end)| *object_end <= start && !used.contains(id))
                    .copied()
                    .ok_or_else(|| {
                        invalid_error(OBJ, "embedded chart BOF has no preceding chart Obj/FtCmo")
                    })?;
                used.insert(object.0);
                let chart = parse_chart(&input[start..end], sheets.len(), limits)?;
                output.push(StoredChart {
                    entry: XlsChartEntry {
                        location: XlsChartLocation::Embedded {
                            sheet_index: sheet.index,
                            object_id: object.0,
                        },
                        chart,
                    },
                    start,
                    end,
                    object: Some((object.1, object.2)),
                });
                index = end_index;
            }
            index += 1;
        }
    }
    if output.len() > limits.max_charts {
        return invalid(CHART, "chart count exceeds limit");
    }
    Ok(output)
}

fn parse_chart(input: &[u8], sheet_count: usize, limits: XlsChartLimits) -> XlsResult<XlsChart> {
    let records = ranges(input)?;
    if records.len() > limits.max_records_per_chart {
        return invalid(CHART, "chart record count exceeds limit");
    }
    if records
        .first()
        .is_none_or(|v| v.kind != BOF || !is_chart_bof(&input[v.body_start..v.body_end]))
        || records.last().is_none_or(|v| v.kind != EOF)
    {
        return invalid(BOF, "chart substream must be bounded by chart BOF and EOF");
    }
    let mut chart = XlsChart {
        groups: Vec::new(),
        ..Default::default()
    };
    let mut depth = 0usize;
    let mut current_series = None;
    let mut series_depth = None;
    let mut current_axis = None;
    let mut cache_index = 0u16;
    let mut pending_axis_line = None;
    for value in records.iter().skip(1).take(records.len() - 2) {
        let data = &input[value.body_start..value.body_end];
        match value.kind {
            BEGIN => {
                if !data.is_empty() {
                    return invalid(BEGIN, "Begin record must be empty");
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| XlsError::InvalidData("chart nesting overflow".into()))?;
                if depth > 128 {
                    return invalid(BEGIN, "chart nesting exceeds limit");
                }
            },
            END => {
                if !data.is_empty() || depth == 0 {
                    return invalid(END, "unbalanced or nonempty End record");
                }
                if series_depth == Some(depth) {
                    current_series = None;
                    series_depth = None;
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
                if chart.sheet_properties & 0xffff_ffe0 != 0 || data[2] > 2 || data[3] != 0 {
                    return invalid(
                        SHT_PROPS,
                        "ShtProps reserved bits or blank mode are invalid",
                    );
                }
            },
            SERIES => {
                exact(data, 12, SERIES)?;
                if chart.series.len() >= limits.max_series {
                    return invalid(SERIES, "series count exceeds limit");
                }
                let kind = match u16_at(data, 0)? {
                    1 => XlsChartDataKind::Numeric,
                    3 => XlsChartDataKind::Text,
                    _ => return invalid(SERIES, "invalid category data type"),
                };
                if u16_at(data, 2)? != 1 || u16_at(data, 8)? != 1 {
                    return invalid(SERIES, "series numeric data type fields are invalid");
                }
                chart.series.push(XlsChartSeries {
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
            0x1051 => {
                let link = parse_link(data, sheet_count, limits)?;
                let series = current_series.ok_or_else(|| {
                    invalid_error(0x1051, "BRAI appears outside a Series collection")
                })?;
                chart.series[series].links.push(link);
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
                let text = parse_short_text(data)?;
                if let Some(series) =
                    current_series.filter(|index| chart.series[*index].name.is_none())
                {
                    chart.series[series].name = Some(text);
                } else if chart.title.is_none() {
                    chart.title = Some(text);
                }
            },
            CHART_FORMAT => {
                exact(data, 20, CHART_FORMAT)?;
                if data[..16].iter().any(|v| *v != 0) || u16_at(data, 16)? & !1 != 0 {
                    return invalid(CHART_FORMAT, "ChartFormat reserved fields are nonzero");
                }
                let order = u16_at(data, 18)?;
                chart.groups.push(XlsChartGroup {
                    order,
                    vary_colors: u16_at(data, 16)? & 1 != 0,
                    kind: XlsChartGroupKind::Line { flags: 0 },
                });
            },
            BAR | LINE | PIE | AREA | SCATTER | RADAR | RADAR_AREA | SURFACE => {
                let kind = parse_group(value.kind, data)?;
                if let Some(group) = chart.groups.last_mut() {
                    group.kind = kind;
                } else {
                    chart.groups.push(XlsChartGroup {
                        order: 0,
                        vary_colors: false,
                        kind,
                    });
                }
            },
            CRT_LINE | DROP_BAR => {
                if let Some(group) = chart.groups.last_mut()
                    && matches!(group.kind, XlsChartGroupKind::Line { .. })
                {
                    let flags = match group.kind {
                        XlsChartGroupKind::Line { flags } => flags,
                        _ => 0,
                    };
                    group.kind = XlsChartGroupKind::Stock { flags };
                }
            },
            AXIS => {
                exact(data, 18, AXIS)?;
                if data[2..].iter().any(|v| *v != 0) {
                    return invalid(AXIS, "Axis reserved fields are nonzero");
                }
                let kind = match u16_at(data, 0)? {
                    0 => XlsChartAxisKind::CategoryOrHorizontal,
                    1 => XlsChartAxisKind::ValueOrVertical,
                    2 => XlsChartAxisKind::Series,
                    _ => return invalid(AXIS, "invalid axis kind"),
                };
                chart.axes.push(XlsChartAxis {
                    kind,
                    scale: None,
                    tick: None,
                    lines: Vec::new(),
                });
                current_axis = Some(chart.axes.len() - 1);
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
                    .scale = Some(XlsChartAxisScale {
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
                    .tick = Some(XlsChartTick {
                    major: data[0],
                    minor: data[1],
                    label_position: data[2],
                    background: data[3],
                    color: array_at(data, 4)?,
                    flags: u16_at(data, 24)?,
                });
            },
            AXIS_LINE => {
                exact(data, 2, AXIS_LINE)?;
                let kind = match u16_at(data, 0)? {
                    0 => XlsChartAxisLineKind::Axis,
                    1 => XlsChartAxisLineKind::MajorGridlines,
                    2 => XlsChartAxisLineKind::MinorGridlines,
                    3 => XlsChartAxisLineKind::WallsOrFloor,
                    _ => return invalid(AXIS_LINE, "invalid AxisLine kind"),
                };
                let axis = current_axis
                    .ok_or_else(|| invalid_error(AXIS_LINE, "AxisLine appears before Axis"))?;
                chart
                    .axes
                    .get_mut(axis)
                    .ok_or_else(|| invalid_error(AXIS_LINE, "Axis index is invalid"))?
                    .lines
                    .push(XlsChartAxisLine { kind, format: None });
                pending_axis_line = Some(axis);
            },
            LINE_FORMAT => {
                let format = parse_line_format(data)?;
                if let Some(axis) = pending_axis_line.take() {
                    let line = chart
                        .axes
                        .get_mut(axis)
                        .and_then(|axis| axis.lines.last_mut())
                        .ok_or_else(|| {
                            invalid_error(LINE_FORMAT, "pending axis line is missing")
                        })?;
                    line.format = Some(format.clone());
                }
                chart.formatting.push(XlsChartFormatting::Line(format));
            },
            AREA_FORMAT => chart
                .formatting
                .push(XlsChartFormatting::Area(parse_area_format(data)?)),
            MARKER_FORMAT => chart.formatting.push(XlsChartFormatting::Marker {
                data: data.to_vec(),
            }),
            DATA_FORMAT => {
                exact(data, 8, DATA_FORMAT)?;
                chart.formatting.push(XlsChartFormatting::Data {
                    point: u16_at(data, 0)?,
                    series: u16_at(data, 2)?,
                    flags: u16_at(data, 6)?,
                });
            },
            PIE_FORMAT => {
                exact(data, 2, PIE_FORMAT)?;
                chart.formatting.push(XlsChartFormatting::Pie {
                    explosion: u16_at(data, 0)?,
                });
            },
            LEGEND => {
                exact(data, 20, LEGEND)?;
                chart.legend = Some(XlsChartLegend {
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
                exact(data, 0, PLOT_AREA)?;
                chart.plot_area_present = true;
            },
            DATA_LAB_EXT | DATA_LAB_EXT_CONTENTS | TEXT => {
                chart.data_labels.push(XlsChartDataLabel {
                    record_type: value.kind,
                    data: data.to_vec(),
                })
            },
            CONTINUE | SERIES_LIST | CAT_SER_RANGE | DEFAULT_TEXT | FONT_X | OBJECT_LINK
            | FRAME | PLOT_GROWTH => chart.unknown_records.push(XlsChartRawRecord {
                record_type: value.kind,
                data: data.to_vec(),
            }),
            SI_INDEX => {
                exact(data, 2, SI_INDEX)?;
                cache_index = u16_at(data, 0)?;
            },
            NUMBER => {
                exact(data, 14, NUMBER)?;
                chart.cached_values.push(XlsChartCacheEntry {
                    cache_index,
                    row: u16_at(data, 0)?,
                    column: u16_at(data, 2)?,
                    value: XlsChartCachedValue::Number(f64_at(data, 6)?),
                });
            },
            LABEL => {
                if data.len() < 8 {
                    return invalid(LABEL, "cached Label is truncated");
                }
                chart.cached_values.push(XlsChartCacheEntry {
                    cache_index,
                    row: u16_at(data, 0)?,
                    column: u16_at(data, 2)?,
                    value: XlsChartCachedValue::Text(parse_biff8_string(&data[6..])?),
                });
            },
            BOF | EOF => return invalid(value.kind, "nested BOF/EOF is invalid in a chart"),
            _ => chart.unknown_records.push(XlsChartRawRecord {
                record_type: value.kind,
                data: data.to_vec(),
            }),
        }
    }
    if depth != 0 {
        return invalid(BEGIN, "chart Begin/End collections are unbalanced");
    }
    chart.validate(sheet_count, limits)?;
    Ok(chart)
}

fn validate_link(
    link: &XlsChartDataLink,
    _sheet_count: usize,
    limits: XlsChartLimits,
) -> XlsResult<()> {
    if link.role > 3 || link.source_type > 2 || link.formula_tokens.len() > limits.max_formula_bytes
    {
        return invalid(
            0x1051,
            "BRAI role, source type, or formula length is invalid",
        );
    }
    if link.source_type == 0 && !link.formula_tokens.is_empty() {
        return invalid(0x1051, "automatic BRAI must have an empty formula");
    }
    for value in &link.references {
        if value.first_row > value.last_row
            || value.first_column > value.last_column
            || value.last_column > 255
        {
            return invalid(
                0x1051,
                "chart formula reference is outside workbook or BIFF8 grid bounds",
            );
        }
    }
    Ok(())
}

fn parse_link(
    data: &[u8],
    sheet_count: usize,
    limits: XlsChartLimits,
) -> XlsResult<XlsChartDataLink> {
    if data.len() < 8 {
        return invalid(0x1051, "BRAI is truncated");
    }
    let flags = u16_at(data, 2)?;
    if flags & !1 != 0 {
        return invalid(0x1051, "BRAI reserved flags are nonzero");
    }
    let len = usize::from(u16_at(data, 6)?);
    if data.len() != 8 + len {
        return invalid(0x1051, "BRAI formula length mismatch");
    }
    let formula_tokens = data[8..].to_vec();
    let references = parse_chart_references(&formula_tokens)?;
    let value = XlsChartDataLink {
        role: data[0],
        source_type: data[1],
        unlinked_number_format: flags & 1 != 0,
        number_format: u16_at(data, 4)?,
        formula_tokens,
        references,
    };
    validate_link(&value, sheet_count, limits)?;
    Ok(value)
}

fn parse_chart_references(tokens: &[u8]) -> XlsResult<Vec<XlsChartCellReference>> {
    if tokens.is_empty() {
        return Ok(Vec::new());
    }
    let opcode = tokens[0] & 0x1f;
    match (opcode, tokens.len()) {
        (0x1a, 7) => {
            let col = u16_at(tokens, 5)? & 0x3fff;
            Ok(vec![XlsChartCellReference {
                extern_sheet_index: u16_at(tokens, 1)?,
                first_row: u16_at(tokens, 3)?,
                last_row: u16_at(tokens, 3)?,
                first_column: col,
                last_column: col,
            }])
        },
        (0x1b, 11) => Ok(vec![XlsChartCellReference {
            extern_sheet_index: u16_at(tokens, 1)?,
            first_row: u16_at(tokens, 3)?,
            last_row: u16_at(tokens, 5)?,
            first_column: u16_at(tokens, 7)? & 0x3fff,
            last_column: u16_at(tokens, 9)? & 0x3fff,
        }]),
        _ => Ok(Vec::new()),
    }
}

fn parse_group(kind: u16, data: &[u8]) -> XlsResult<XlsChartGroupKind> {
    Ok(match kind {
        BAR => {
            exact(data, 6, BAR)?;
            XlsChartGroupKind::Bar {
                overlap: i16_at(data, 0)?,
                gap: u16_at(data, 2)?,
                flags: u16_at(data, 4)?,
            }
        },
        LINE => {
            exact(data, 2, LINE)?;
            XlsChartGroupKind::Line {
                flags: u16_at(data, 0)?,
            }
        },
        AREA => {
            exact(data, 2, AREA)?;
            XlsChartGroupKind::Area {
                flags: u16_at(data, 0)?,
            }
        },
        PIE => {
            exact(data, 6, PIE)?;
            XlsChartGroupKind::Pie {
                rotation: u16_at(data, 0)?,
                hole_size: u16_at(data, 2)?,
                flags: u16_at(data, 4)?,
            }
        },
        SCATTER => {
            exact(data, 6, SCATTER)?;
            XlsChartGroupKind::Scatter {
                bubble_size_percent: u16_at(data, 0)?,
                bubble_size_type: u16_at(data, 2)?,
                flags: u16_at(data, 4)?,
            }
        },
        RADAR | RADAR_AREA => {
            exact(data, 2, kind)?;
            XlsChartGroupKind::Radar {
                filled: kind == RADAR_AREA,
                flags: u16_at(data, 0)?,
            }
        },
        SURFACE => {
            exact(data, 2, SURFACE)?;
            XlsChartGroupKind::Surface {
                flags: u16_at(data, 0)?,
            }
        },
        _ => return invalid(kind, "unsupported chart group record"),
    })
}

fn serialize_chart(
    chart: &XlsChart,
    sheet_count: usize,
    limits: XlsChartLimits,
) -> XlsResult<Vec<u8>> {
    chart.validate(sheet_count, limits)?;
    let mut out = Vec::new();
    out.extend(record(BOF, &chart_bof())?);
    let mut geometry = Vec::new();
    for value in [chart.x, chart.y, chart.width, chart.height] {
        geometry.extend(value.to_le_bytes());
    }
    out.extend(record(CHART, &geometry)?);
    out.extend(record(BEGIN, &[])?);
    out.extend(record(SHT_PROPS, &chart.sheet_properties.to_le_bytes())?);
    for series in &chart.series {
        let mut body = Vec::new();
        body.extend(
            (if series.category_kind == XlsChartDataKind::Numeric {
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
        out.extend(record(SERIES, &body)?);
        out.extend(record(BEGIN, &[])?);
        for link in &series.links {
            let mut data = vec![link.role, link.source_type];
            data.extend(u16::from(link.unlinked_number_format).to_le_bytes());
            data.extend(link.number_format.to_le_bytes());
            data.extend(
                u16::try_from(link.formula_tokens.len())
                    .map_err(|_| XlsError::InvalidData("chart formula exceeds u16".into()))?
                    .to_le_bytes(),
            );
            data.extend(&link.formula_tokens);
            out.extend(record(0x1051, &data)?);
        }
        if let Some(name) = &series.name {
            out.extend(record(SERIES_TEXT, &short_text(name)?)?);
        }
        out.extend(record(SER_TO_CRT, &series.chart_group.to_le_bytes())?);
        out.extend(record(END, &[])?);
    }
    out.extend(record(
        AXES_USED,
        &(if chart.groups.len() > 1 { 2u16 } else { 1 }).to_le_bytes(),
    )?);
    out.extend(record(AXIS_PARENT, &[0; 18])?);
    out.extend(record(BEGIN, &[])?);
    for axis in &chart.axes {
        let mut body = vec![0; 18];
        body[..2].copy_from_slice(
            &(match axis.kind {
                XlsChartAxisKind::CategoryOrHorizontal => 0u16,
                XlsChartAxisKind::ValueOrVertical => 1,
                XlsChartAxisKind::Series => 2,
            })
            .to_le_bytes(),
        );
        out.extend(record(AXIS, &body)?);
        out.extend(record(BEGIN, &[])?);
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
            out.extend(record(VALUE_RANGE, &data)?);
        }
        if let Some(tick) = &axis.tick {
            let mut data = vec![0; 26];
            data[0] = tick.major;
            data[1] = tick.minor;
            data[2] = tick.label_position;
            data[3] = tick.background;
            data[4..8].copy_from_slice(&tick.color);
            data[24..26].copy_from_slice(&tick.flags.to_le_bytes());
            out.extend(record(TICK, &data)?);
        }
        for line in &axis.lines {
            let id = match line.kind {
                XlsChartAxisLineKind::Axis => 0u16,
                XlsChartAxisLineKind::MajorGridlines => 1,
                XlsChartAxisLineKind::MinorGridlines => 2,
                XlsChartAxisLineKind::WallsOrFloor => 3,
            };
            out.extend(record(AXIS_LINE, &id.to_le_bytes())?);
            if let Some(format) = &line.format {
                out.extend(record(LINE_FORMAT, &write_line(format))?);
            }
        }
        out.extend(record(END, &[])?);
    }
    for group in &chart.groups {
        let mut data = vec![0; 20];
        data[16..18].copy_from_slice(&u16::from(group.vary_colors).to_le_bytes());
        data[18..20].copy_from_slice(&group.order.to_le_bytes());
        out.extend(record(CHART_FORMAT, &data)?);
        out.extend(record(BEGIN, &[])?);
        write_group(&mut out, &group.kind)?;
        out.extend(record(END, &[])?);
    }
    if chart.plot_area_present {
        out.extend(record(PLOT_AREA, &[])?);
    }
    if let Some(legend) = &chart.legend {
        let mut data = Vec::new();
        for v in [legend.x, legend.y, legend.width, legend.height] {
            data.extend(v.to_le_bytes());
        }
        data.push(legend.position);
        data.push(legend.spacing);
        data.extend(legend.flags.to_le_bytes());
        out.extend(record(LEGEND, &data)?);
    }
    if let Some(title) = &chart.title {
        out.extend(record(SERIES_TEXT, &short_text(title)?)?);
    }
    for format in &chart.formatting {
        match format {
            XlsChartFormatting::Line(value) => out.extend(record(LINE_FORMAT, &write_line(value))?),
            XlsChartFormatting::Area(value) => out.extend(record(AREA_FORMAT, &write_area(value))?),
            XlsChartFormatting::Marker { data } => out.extend(record(MARKER_FORMAT, data)?),
            XlsChartFormatting::Data {
                point,
                series,
                flags,
            } => {
                let mut data = Vec::new();
                data.extend(point.to_le_bytes());
                data.extend(series.to_le_bytes());
                data.extend(0u16.to_le_bytes());
                data.extend(flags.to_le_bytes());
                out.extend(record(DATA_FORMAT, &data)?);
            },
            XlsChartFormatting::Pie { explosion } => {
                out.extend(record(PIE_FORMAT, &explosion.to_le_bytes())?)
            },
        }
    }
    for label in &chart.data_labels {
        out.extend(record(label.record_type, &label.data)?);
    }
    let mut active_cache = None;
    for value in &chart.cached_values {
        if active_cache != Some(value.cache_index) {
            out.extend(record(SI_INDEX, &value.cache_index.to_le_bytes())?);
            active_cache = Some(value.cache_index);
        }
        match &value.value {
            XlsChartCachedValue::Number(number) => {
                if !number.is_finite() {
                    return invalid(NUMBER, "cached chart number must be finite");
                }
                let mut data = Vec::new();
                data.extend(value.row.to_le_bytes());
                data.extend(value.column.to_le_bytes());
                data.extend(0u16.to_le_bytes());
                data.extend(number.to_le_bytes());
                out.extend(record(NUMBER, &data)?);
            },
            XlsChartCachedValue::Text(text) => {
                let mut data = Vec::new();
                data.extend(value.row.to_le_bytes());
                data.extend(value.column.to_le_bytes());
                data.extend(0u16.to_le_bytes());
                data.extend(biff8_string(text)?);
                out.extend(record(LABEL, &data)?);
            },
            XlsChartCachedValue::Blank => {},
        }
    }
    for value in &chart.unknown_records {
        if !known_record(value.record_type) {
            out.extend(record(value.record_type, &value.data)?);
        }
    }
    out.extend(record(END, &[])?);
    out.extend(record(END, &[])?);
    out.extend(record(EOF, &[])?);
    Ok(out)
}

fn write_group(out: &mut Vec<u8>, kind: &XlsChartGroupKind) -> XlsResult<()> {
    match kind {
        XlsChartGroupKind::Line { flags } => out.extend(record(LINE, &flags.to_le_bytes())?),
        XlsChartGroupKind::Area { flags } => out.extend(record(AREA, &flags.to_le_bytes())?),
        XlsChartGroupKind::Bar {
            overlap,
            gap,
            flags,
        } => {
            let mut d = overlap.to_le_bytes().to_vec();
            d.extend(gap.to_le_bytes());
            d.extend(flags.to_le_bytes());
            out.extend(record(BAR, &d)?);
        },
        XlsChartGroupKind::Pie {
            rotation,
            hole_size,
            flags,
        } => {
            let mut d = rotation.to_le_bytes().to_vec();
            d.extend(hole_size.to_le_bytes());
            d.extend(flags.to_le_bytes());
            out.extend(record(PIE, &d)?);
        },
        XlsChartGroupKind::Scatter {
            bubble_size_percent,
            bubble_size_type,
            flags,
        } => {
            let mut d = bubble_size_percent.to_le_bytes().to_vec();
            d.extend(bubble_size_type.to_le_bytes());
            d.extend(flags.to_le_bytes());
            out.extend(record(SCATTER, &d)?);
        },
        XlsChartGroupKind::Radar { filled, flags } => out.extend(record(
            if *filled { RADAR_AREA } else { RADAR },
            &flags.to_le_bytes(),
        )?),
        XlsChartGroupKind::Surface { flags } => out.extend(record(SURFACE, &flags.to_le_bytes())?),
        XlsChartGroupKind::Stock { flags } => {
            out.extend(record(LINE, &flags.to_le_bytes())?);
            out.extend(record(CRT_LINE, &0u16.to_le_bytes())?);
            out.extend(record(DROP_BAR, &[0; 4])?);
        },
    }
    Ok(())
}

fn rewrite_workbook_charts(
    input: &[u8],
    original: &[StoredChart],
    desired: &[XlsChartEntry],
    limits: XlsChartLimits,
) -> XlsResult<Vec<u8>> {
    let (refs, sheets) = bindings(input)?;
    let mut output = input[..sheets.first().map_or(input.len(), |v| v.start)].to_vec();
    let mut new_offsets = HashMap::new();
    for sheet in &sheets {
        new_offsets.insert(sheet.start, output.len());
        let existing = original
            .iter()
            .filter(|v| v.entry.location.sheet_index() == sheet.index)
            .collect::<Vec<_>>();
        let wanted = desired
            .iter()
            .filter(|v| v.location.sheet_index() == sheet.index)
            .collect::<Vec<_>>();
        if sheet.kind == 2 {
            let current = existing
                .first()
                .ok_or_else(|| invalid_error(CHART, "chart sheet has no parsed chart"))?;
            let replacement = wanted
                .iter()
                .find(|v| v.location == current.entry.location)
                .ok_or_else(|| invalid_error(CHART, "chart sheet removal is not supported"))?;
            output.extend(serialize_chart(&replacement.chart, sheets.len(), limits)?);
            continue;
        }
        let changed = existing.iter().map(|v| &v.entry).collect::<Vec<_>>() != wanted;
        if !changed {
            output.extend_from_slice(&input[sheet.start..sheet.end]);
            continue;
        }
        if sheet.kind != 0 {
            return invalid(CHART, "embedded charts can only be written to worksheets");
        }
        let mut remove = Vec::new();
        for value in &existing {
            remove.push((value.start, value.end));
            if let Some(range) = value.object {
                remove.push(range);
            }
        }
        remove.sort_unstable();
        let segment = &input[sheet.start..sheet.end];
        let records = ranges(segment)?;
        let eof = records
            .iter()
            .rfind(|v| v.kind == EOF)
            .ok_or_else(|| invalid_error(EOF, "worksheet has no EOF"))?
            .start
            + sheet.start;
        let mut cursor = sheet.start;
        for (start, end) in remove {
            if start < cursor || end > sheet.end {
                return invalid(CHART, "overlapping chart/object ranges");
            }
            if start >= eof {
                break;
            }
            output.extend_from_slice(&input[cursor..start]);
            cursor = end;
        }
        if cursor > eof {
            return invalid(CHART, "chart range crosses worksheet EOF");
        }
        output.extend_from_slice(&input[cursor..eof]);
        for value in wanted {
            let object_id = match value.location {
                XlsChartLocation::Embedded { object_id, .. } => object_id,
                _ => return invalid(CHART, "chart sheet cannot be embedded"),
            };
            output.extend(chart_object_record(object_id)?);
            output.extend(serialize_chart(&value.chart, sheets.len(), limits)?);
        }
        output.extend_from_slice(&input[eof..sheet.end]);
    }
    for (reference, old) in refs {
        let new = *new_offsets
            .get(&old)
            .ok_or_else(|| invalid_error(BOUNDSHEET, "BoundSheet target is missing"))?;
        output[reference..reference + 4].copy_from_slice(
            &u32::try_from(new)
                .map_err(|_| XlsError::InvalidData("BoundSheet offset exceeds u32".into()))?
                .to_le_bytes(),
        );
    }
    if output.len() > limits.max_workbook_bytes {
        return invalid(CHART, "rewritten Workbook exceeds limit");
    }
    Ok(output)
}

#[derive(Clone, Copy)]
struct Range {
    start: usize,
    end: usize,
    kind: u16,
    body_start: usize,
    body_end: usize,
}
#[derive(Clone, Copy)]
struct Sheet {
    index: usize,
    start: usize,
    end: usize,
    kind: u8,
}
#[allow(clippy::type_complexity)]
fn bindings(input: &[u8]) -> XlsResult<(Vec<(usize, usize)>, Vec<Sheet>)> {
    let mut refs = Vec::new();
    for value in ranges(input)? {
        if value.kind == BOUNDSHEET {
            let data = &input[value.body_start..value.body_end];
            if data.len() < 8 {
                return invalid(BOUNDSHEET, "BoundSheet is truncated");
            }
            refs.push((value.start + 4, u32_at(data, 0)? as usize, data[5]));
        }
    }
    let mut physical = refs
        .iter()
        .enumerate()
        .map(|(index, (_, start, kind))| (index, *start, *kind))
        .collect::<Vec<_>>();
    physical.sort_by_key(|v| v.1);
    if physical.is_empty()
        || physical.windows(2).any(|v| v[0].1 >= v[1].1)
        || physical.iter().any(|v| v.1 >= input.len())
    {
        return invalid(BOUNDSHEET, "invalid or missing BoundSheet offsets");
    }
    let sheets = physical
        .iter()
        .enumerate()
        .map(|(slot, (index, start, kind))| Sheet {
            index: *index,
            start: *start,
            end: physical.get(slot + 1).map_or(input.len(), |v| v.1),
            kind: *kind,
        })
        .collect();
    Ok((refs.into_iter().map(|(p, o, _)| (p, o)).collect(), sheets))
}
fn ranges(input: &[u8]) -> XlsResult<Vec<Range>> {
    let mut out = Vec::new();
    let mut offset = 0;
    while offset < input.len() {
        let h = input
            .get(offset..offset + 4)
            .ok_or(XlsError::InvalidLength {
                expected: offset + 4,
                found: input.len(),
            })?;
        let kind = u16::from_le_bytes([h[0], h[1]]);
        let len = usize::from(u16::from_le_bytes([h[2], h[3]]));
        if len > MAX_BIFF_DATA {
            return invalid(kind, "BIFF record exceeds 8224 bytes");
        }
        let end = offset
            .checked_add(4 + len)
            .ok_or_else(|| XlsError::InvalidData("BIFF record length overflow".into()))?;
        if end > input.len() {
            return Err(XlsError::InvalidLength {
                expected: end,
                found: input.len(),
            });
        }
        out.push(Range {
            start: offset,
            end,
            kind,
            body_start: offset + 4,
            body_end: end,
        });
        offset = end;
    }
    Ok(out)
}
fn parse_chart_object(data: &[u8]) -> XlsResult<Option<u16>> {
    let mut offset = 0;
    while offset < data.len() {
        let h = data
            .get(offset..offset + 4)
            .ok_or_else(|| invalid_error(OBJ, "truncated Obj subrecord"))?;
        let kind = u16::from_le_bytes([h[0], h[1]]);
        let len = usize::from(u16::from_le_bytes([h[2], h[3]]));
        offset += 4;
        let end = offset
            .checked_add(len)
            .ok_or_else(|| XlsError::InvalidData("Obj length overflow".into()))?;
        let body = data
            .get(offset..end)
            .ok_or_else(|| invalid_error(OBJ, "truncated Obj subrecord body"))?;
        if kind == 0x15 {
            if len != 18 {
                return invalid(OBJ, "FtCmo must contain 18 bytes");
            }
            return if u16_at(body, 0)? == 5 {
                Ok(Some(u16_at(body, 2)?))
            } else {
                Ok(None)
            };
        }
        offset = end;
    }
    Ok(None)
}
fn chart_object_record(id: u16) -> XlsResult<Vec<u8>> {
    let mut body = Vec::new();
    body.extend(0x15u16.to_le_bytes());
    body.extend(18u16.to_le_bytes());
    body.extend(5u16.to_le_bytes());
    body.extend(id.to_le_bytes());
    body.extend(0x6011u16.to_le_bytes());
    body.extend([0; 12]);
    body.extend(0u16.to_le_bytes());
    body.extend(0u16.to_le_bytes());
    record(OBJ, &body)
}
fn is_chart_bof(data: &[u8]) -> bool {
    data.len() >= 4
        && u16::from_le_bytes([data[0], data[1]]) == 0x0600
        && u16::from_le_bytes([data[2], data[3]]) == 0x0020
}
fn chart_bof() -> Vec<u8> {
    bof_body(0x0020)
}
fn bof_body(kind: u16) -> Vec<u8> {
    let mut d = Vec::new();
    d.extend(0x0600u16.to_le_bytes());
    d.extend(kind.to_le_bytes());
    d.extend(0x0dbbu16.to_le_bytes());
    d.extend(0x07ccu16.to_le_bytes());
    d.extend(0u32.to_le_bytes());
    d.extend(6u32.to_le_bytes());
    d
}
fn parse_line_format(data: &[u8]) -> XlsResult<XlsChartLineFormat> {
    exact(data, 12, LINE_FORMAT)?;
    Ok(XlsChartLineFormat {
        color: array_at(data, 0)?,
        pattern: u16_at(data, 4)?,
        weight: i16_at(data, 6)?,
        flags: u16_at(data, 8)?,
        color_index: u16_at(data, 10)?,
    })
}
fn write_line(v: &XlsChartLineFormat) -> Vec<u8> {
    let mut d = v.color.to_vec();
    d.extend(v.pattern.to_le_bytes());
    d.extend(v.weight.to_le_bytes());
    d.extend(v.flags.to_le_bytes());
    d.extend(v.color_index.to_le_bytes());
    d
}
fn parse_area_format(data: &[u8]) -> XlsResult<XlsChartAreaFormat> {
    exact(data, 16, AREA_FORMAT)?;
    Ok(XlsChartAreaFormat {
        foreground: array_at(data, 0)?,
        background: array_at(data, 4)?,
        pattern: u16_at(data, 8)?,
        flags: u16_at(data, 10)?,
        foreground_index: u16_at(data, 12)?,
        background_index: u16_at(data, 14)?,
    })
}
fn write_area(v: &XlsChartAreaFormat) -> Vec<u8> {
    let mut d = v.foreground.to_vec();
    d.extend(v.background);
    d.extend(v.pattern.to_le_bytes());
    d.extend(v.flags.to_le_bytes());
    d.extend(v.foreground_index.to_le_bytes());
    d.extend(v.background_index.to_le_bytes());
    d
}
fn parse_short_text(data: &[u8]) -> XlsResult<String> {
    if data.len() < 4 || u16_at(data, 0)? != 0 {
        return invalid(
            SERIES_TEXT,
            "SeriesText is truncated or reserved field is nonzero",
        );
    }
    parse_biff8_string(&data[2..])
}
fn short_text(value: &str) -> XlsResult<Vec<u8>> {
    let mut d = 0u16.to_le_bytes().to_vec();
    d.extend(biff8_string(value)?);
    Ok(d)
}
fn parse_biff8_string(data: &[u8]) -> XlsResult<String> {
    if data.len() < 2 {
        return invalid(SERIES_TEXT, "chart string is truncated");
    }
    let count = usize::from(data[0]);
    let wide = data[1] & 1 != 0;
    if data[1] & !1 != 0 {
        return invalid(SERIES_TEXT, "chart string uses unsupported option flags");
    }
    let need = 2 + count * (if wide { 2 } else { 1 });
    if data.len() != need {
        return invalid(SERIES_TEXT, "chart string length mismatch");
    }
    if wide {
        let units = data[2..]
            .chunks_exact(2)
            .map(|v| u16::from_le_bytes([v[0], v[1]]))
            .collect::<Vec<_>>();
        String::from_utf16(&units)
            .map_err(|_| invalid_error(SERIES_TEXT, "invalid UTF-16 chart string"))
    } else {
        Ok(data[2..].iter().map(|v| char::from(*v)).collect())
    }
}
fn biff8_string(value: &str) -> XlsResult<Vec<u8>> {
    let units = value.encode_utf16().collect::<Vec<_>>();
    if units.len() > 255 {
        return invalid(SERIES_TEXT, "chart string exceeds 255 UTF-16 code units");
    }
    let wide = units.iter().any(|v| *v > 255);
    let mut d = vec![units.len() as u8, u8::from(wide)];
    if wide {
        for v in units {
            d.extend(v.to_le_bytes());
        }
    } else {
        d.extend(units.into_iter().map(|v| v as u8));
    }
    Ok(d)
}
fn record(kind: u16, data: &[u8]) -> XlsResult<Vec<u8>> {
    if data.len() > MAX_BIFF_DATA {
        return invalid(kind, "BIFF record exceeds 8224 bytes");
    }
    let mut d = kind.to_le_bytes().to_vec();
    d.extend((data.len() as u16).to_le_bytes());
    d.extend(data);
    Ok(d)
}
fn known_record(kind: u16) -> bool {
    matches!(
        kind,
        BOF | EOF
            | CHART
            | SERIES
            | 0x1051
            | SER_TO_CRT
            | SERIES_TEXT
            | CHART_FORMAT
            | BAR
            | LINE
            | PIE
            | AREA
            | SCATTER
            | RADAR
            | RADAR_AREA
            | SURFACE
            | CRT_LINE
            | DROP_BAR
            | AXIS
            | VALUE_RANGE
            | TICK
            | AXIS_LINE
            | LINE_FORMAT
            | AREA_FORMAT
            | MARKER_FORMAT
            | DATA_FORMAT
            | PIE_FORMAT
            | LEGEND
            | PLOT_AREA
            | DATA_LAB_EXT
            | DATA_LAB_EXT_CONTENTS
            | TEXT
            | SI_INDEX
            | NUMBER
            | LABEL
            | BEGIN
            | END
            | SHT_PROPS
            | AXES_USED
            | AXIS_PARENT
    )
}
fn validate_limits(v: XlsChartLimits) -> XlsResult<()> {
    if v.max_workbook_bytes == 0
        || v.max_charts == 0
        || v.max_records_per_chart == 0
        || v.max_series == 0
        || v.max_groups == 0
        || v.max_axes == 0
        || v.max_formula_bytes == 0
        || v.max_cached_values == 0
        || v.max_unknown_bytes == 0
    {
        return Err(XlsError::InvalidData(
            "all chart limits must be nonzero".into(),
        ));
    }
    Ok(())
}
fn bounded_count(data: &[u8], offset: usize) -> XlsResult<u16> {
    let v = u16_at(data, offset)?;
    if v > 32767 {
        return invalid(SERIES, "series value count exceeds 32767");
    }
    Ok(v)
}
fn exact(data: &[u8], len: usize, kind: u16) -> XlsResult<()> {
    if data.len() != len {
        return invalid(kind, format!("record must contain {len} bytes"));
    }
    Ok(())
}
fn u16_at(data: &[u8], o: usize) -> XlsResult<u16> {
    Ok(u16::from_le_bytes(array_at(data, o)?))
}
fn i16_at(data: &[u8], o: usize) -> XlsResult<i16> {
    Ok(u16_at(data, o)? as i16)
}
fn u32_at(data: &[u8], o: usize) -> XlsResult<u32> {
    Ok(u32::from_le_bytes(array_at(data, o)?))
}
fn i32_at(data: &[u8], o: usize) -> XlsResult<i32> {
    Ok(u32_at(data, o)? as i32)
}
fn f64_at(data: &[u8], o: usize) -> XlsResult<f64> {
    Ok(f64::from_le_bytes(array_at(data, o)?))
}
fn array_at<const N: usize>(data: &[u8], offset: usize) -> XlsResult<[u8; N]> {
    let expected = offset
        .checked_add(N)
        .ok_or_else(|| XlsError::InvalidData("record field offset overflow".into()))?;
    data.get(offset..expected)
        .and_then(|bytes| bytes.try_into().ok())
        .ok_or(XlsError::InvalidLength {
            expected,
            found: data.len(),
        })
}
fn invalid_error(kind: u16, message: impl Into<String>) -> XlsError {
    XlsError::InvalidRecord {
        record_type: kind,
        message: message.into(),
    }
}
fn invalid<T>(kind: u16, message: impl Into<String>) -> XlsResult<T> {
    Err(invalid_error(kind, message))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn generated_workbook_round_trips_through_chart_editor() {
        let mut chart = XlsChart {
            title: Some("Sales".into()),
            ..Default::default()
        };
        chart.series.push(XlsChartSeries {
            category_count: 2,
            value_count: 2,
            links: vec![XlsChartDataLink {
                role: 1,
                source_type: 2,
                unlinked_number_format: false,
                number_format: 0,
                formula_tokens: vec![0x3b, 0, 0, 1, 0, 2, 0, 1, 0, 1, 0],
                references: vec![XlsChartCellReference {
                    extern_sheet_index: 0,
                    first_row: 1,
                    last_row: 2,
                    first_column: 1,
                    last_column: 1,
                }],
            }],
            ..Default::default()
        });
        chart.cached_values.push(XlsChartCacheEntry {
            cache_index: 0,
            row: 0,
            column: 0,
            value: XlsChartCachedValue::Number(42.0),
        });
        let bytes = build_chart_workbook(chart, XlsChartLimits::default()).unwrap();
        let editor = XlsChartEditor::open(bytes, XlsChartLimits::default()).unwrap();
        let charts = editor.charts();
        assert_eq!(charts.len(), 1);
        assert_eq!(
            charts[0].location,
            XlsChartLocation::Embedded {
                sheet_index: 0,
                object_id: GENERATED_CHART_OBJECT_ID
            }
        );
        assert_eq!(charts[0].chart.title.as_deref(), Some("Sales"));
        assert_eq!(charts[0].chart.series.len(), 1);
        assert!(
            charts[0]
                .chart
                .cached_values
                .iter()
                .any(|v| v.value == XlsChartCachedValue::Number(42.0))
        );
    }

    #[test]
    fn generated_combo_round_trips_and_preserves_opaque_records() {
        let mut chart = XlsChart::default();
        chart.groups.push(XlsChartGroup {
            order: 1,
            vary_colors: true,
            kind: XlsChartGroupKind::Pie {
                rotation: 45,
                hole_size: 50,
                flags: 0,
            },
        });
        chart.series.push(XlsChartSeries {
            chart_group: 1,
            ..Default::default()
        });
        chart.unknown_records.push(XlsChartRawRecord {
            record_type: 0x7777,
            data: b"opaque".to_vec(),
        });
        let bytes = serialize_chart(&chart, 2, XlsChartLimits::default()).unwrap();
        let parsed = parse_chart(&bytes, 2, XlsChartLimits::default()).unwrap();
        assert!(matches!(parsed.chart_type(), XlsChartType::Combo(_)));
        assert_eq!(parsed.unknown_records[0].data, b"opaque");
    }

    #[test]
    fn malformed_nesting_formula_and_axis_are_rejected() {
        let mut bytes = record(BOF, &chart_bof()).unwrap();
        bytes.extend(record(CHART, &[0; 16]).unwrap());
        bytes.extend(record(END, &[]).unwrap());
        bytes.extend(record(EOF, &[]).unwrap());
        assert!(parse_chart(&bytes, 1, XlsChartLimits::default()).is_err());
        let mut chart = XlsChart::default();
        chart.series.push(XlsChartSeries {
            links: vec![XlsChartDataLink {
                role: 1,
                source_type: 2,
                unlinked_number_format: false,
                number_format: 0,
                formula_tokens: vec![0x3b, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0],
                references: vec![XlsChartCellReference {
                    extern_sheet_index: 2,
                    first_row: 0,
                    last_row: 0,
                    first_column: 300,
                    last_column: 300,
                }],
            }],
            ..Default::default()
        });
        assert!(chart.validate(1, XlsChartLimits::default()).is_err());
    }

    #[test]
    fn referenced_sheet_removal_and_noncontiguous_reorder_are_rejected_atomically() {
        let mut extern_sheet = 1u16.to_le_bytes().to_vec();
        extern_sheet.extend(0u16.to_le_bytes());
        extern_sheet.extend(0u16.to_le_bytes());
        extern_sheet.extend(2u16.to_le_bytes());
        let internal = HashSet::from([0u16]);
        assert!(
            remap_extern_sheet(&extern_sheet, &internal, &[Some(0), None, Some(1)], None).is_err()
        );
        assert!(
            remap_extern_sheet(&extern_sheet, &internal, &[Some(0), Some(2), Some(1)], None)
                .is_ok()
        );
    }
}
