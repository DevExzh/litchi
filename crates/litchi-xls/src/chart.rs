//! Bounded BIFF8 chart-sheet and embedded-chart metadata and safe mutation.
//!
//! Chart formulas and cached values are inert. This module never evaluates a
//! formula, opens an external workbook, refreshes a cache, or renders a chart.
//! Unsupported records remain available for inspection; edits that cannot
//! prove their original placement are refused instead of being normalized.
//! Existing charts can be inspected, replayed exactly, removed, and reordered.
//! Fresh and replacement authoring is refused until the complete
//! Office-compatible BIFF chart grammar is implemented.

#[cfg(test)]
use std::collections::HashMap;
use std::collections::HashSet;
use std::sync::Arc;

use litchi_ole_common::object::{Editor as ObjectEditor, Limits as ObjectLimits, Targets};

use litchi_biff::MAX_RECORD_BYTES;
use litchi_biff::{Encoder as GraphEncoder, Kind as RecordKind, Limits as BiffLimits, Records};
use litchi_ograph::Limits as GraphLimits;
use litchi_ograph::chart::{Kind as GraphChartKind, Ref as GraphChartRef, Refs as GraphCharts};
pub use litchi_ograph::chart::{format, group};
pub use litchi_ograph::record::{chart3d, frame, line, marker, pie, series};

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
const BLANK: u16 = 0x0201;
const NUMBER: u16 = 0x0203;
const LABEL: u16 = 0x0204;
const CHART: u16 = 0x1002;
const SERIES: u16 = 0x1003;
const DATA_FORMAT: u16 = 0x1006;
const LINE_FORMAT: u16 = 0x1007;
const MARKER_FORMAT: u16 = 0x1009;
const AREA_FORMAT: u16 = 0x100a;
const PIE_FORMAT: u16 = pie::Format::KIND.get();
const SERIES_TEXT: u16 = 0x100d;
const CHART_FORMAT: u16 = 0x1014;
const LEGEND: u16 = 0x1015;
const SERIES_LIST: u16 = 0x1016;
const BAR: u16 = 0x1017;
const LINE: u16 = 0x1018;
const PIE: u16 = 0x1019;
const AREA: u16 = 0x101a;
const SCATTER: u16 = 0x101b;
const CRT_LINE: u16 = line::Line::KIND.get();
const CRT_LINK: u16 = line::Link::KIND.get();
const AXIS: u16 = 0x101d;
const TICK: u16 = 0x101e;
const VALUE_RANGE: u16 = 0x101f;
const CAT_SER_RANGE: u16 = 0x1020;
const AXIS_LINE: u16 = 0x1021;
const DEFAULT_TEXT: u16 = 0x1024;
const TEXT: u16 = 0x1025;
const FONT_X: u16 = 0x1026;
const OBJECT_LINK: u16 = 0x1027;
const FRAME: u16 = frame::Frame::KIND.get();
const BEGIN: u16 = marker::Begin::KIND.get();
const END: u16 = marker::End::KIND.get();
const PLOT_AREA: u16 = marker::PlotArea::KIND.get();
const DROP_BAR: u16 = 0x103d;
const RADAR: u16 = 0x103e;
const SURFACE: u16 = 0x103f;
const RADAR_AREA: u16 = 0x1040;
const AXIS_PARENT: u16 = 0x1041;
const SHT_PROPS: u16 = 0x1044;
const SER_TO_CRT: u16 = 0x1045;
const AXES_USED: u16 = 0x1046;
const SERIES_PARENT: u16 = series::Parent::KIND.get();
const SERIES_FORMAT: u16 = series::Format::KIND.get();
const BAR_SHAPE: u16 = chart3d::BarShape::KIND.get();
const DATA_LAB_EXT: u16 = 0x086a;
const DATA_LAB_EXT_CONTENTS: u16 = 0x086b;
const PLOT_GROWTH: u16 = 0x1064;
const SI_INDEX: u16 = 0x1065;

/// Hard resource bounds for chart discovery and safe mutation.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Limits {
    /// Maximum bytes accepted for the BIFF `Workbook` stream.
    pub max_workbook_bytes: usize,
    /// Maximum chart substreams in one workbook.
    pub max_charts: usize,
    /// Maximum BIFF records in one chart substream.
    pub max_records_per_chart: usize,
    /// Maximum data series in one chart.
    pub max_series: usize,
    /// Maximum chart groups in one chart.
    pub max_groups: usize,
    /// Maximum axes in one chart.
    pub max_axes: usize,
    /// Maximum bytes retained for one inert formula token array.
    pub max_formula_bytes: usize,
    /// Maximum cached cell values in one chart.
    pub max_cached_values: usize,
    /// Maximum aggregate payload bytes retained for unknown records.
    pub max_unknown_bytes: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_workbook_bytes: 128 * 1024 * 1024,
            max_charts: 512,
            max_records_per_chart: 8_192,
            max_series: 255,
            max_groups: 10,
            max_axes: 6,
            max_formula_bytes: MAX_RECORD_BYTES - 8,
            max_cached_values: 32_000,
            max_unknown_bytes: 16 * 1024 * 1024,
        }
    }
}

/// Stable location of one chart in the current workbook revision.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub enum Location {
    /// A chart-sheet tab.
    ChartSheet {
        /// Zero-based workbook tab index.
        sheet_index: usize,
    },
    /// An Obj-linked chart embedded in a worksheet.
    Embedded {
        /// Zero-based workbook tab index.
        sheet_index: usize,
        /// Host OBJ identifier; semantic selectors are preferred.
        object_id: u16,
    },
}

impl Location {
    /// Returns the zero-based host tab index.
    pub fn sheet_index(&self) -> usize {
        match self {
            Self::ChartSheet { sheet_index } | Self::Embedded { sheet_index, .. } => *sheet_index,
        }
    }
}

/// Semantic chart lookup key.
///
/// Names are compared case-insensitively using Unicode lowercase mappings.
/// Embedded chart indexes are zero-based in drawing order on that worksheet.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum Selector<'name> {
    /// The chart occupying a named chart-sheet tab.
    Sheet(&'name str),
    /// One embedded chart on a named worksheet.
    Embedded {
        /// Worksheet tab name.
        sheet: &'name str,
        /// Zero-based chart order on that worksheet.
        index: usize,
    },
}

/// Payload of an unsupported BIFF chart record retained for inspection.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Raw {
    /// BIFF record identifier.
    pub record_type: u16,
    /// Record payload without its four-byte BIFF header.
    pub data: Vec<u8>,
}

/// Scalar kind declared by a chart series.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum DataKind {
    /// IEEE-754 numeric values.
    Numeric,
    /// BIFF strings.
    Text,
}

/// One checked area reference extracted from an inert chart formula.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct CellRef {
    /// Index into the workbook's `ExternSheet` table.
    pub extern_sheet_index: u16,
    /// Inclusive first row.
    pub first_row: u16,
    /// Inclusive last row.
    pub last_row: u16,
    /// Inclusive first column.
    pub first_column: u16,
    /// Inclusive last column.
    pub last_column: u16,
}

/// Semantic part of a series referenced by a data link.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum Role {
    /// Series, legend-entry, or trendline name.
    Name = 0,
    /// Values, or horizontal values for scatter and bubble charts.
    Values = 1,
    /// Categories, or vertical values for scatter and bubble charts.
    Categories = 2,
    /// Bubble-size values.
    Bubbles = 3,
}

/// Kind of source referenced by a chart data link.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum Source {
    /// Excel generated the category, series name, or bubble size.
    Automatic = 0,
    /// A literal text or value is held by the formula field.
    Literal = 1,
    /// A formula references a range of worksheet cells.
    Cells = 2,
}

/// Inert chart formula and its validated workbook references.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DataLink {
    /// Semantic series part supplied by this link.
    pub role: Role,
    /// Source kind supplied by this link.
    pub source: Source,
    /// Whether the link does not inherit its source number format.
    pub unlinked_number_format: bool,
    /// BIFF number-format index.
    pub number_format: u16,
    /// Original formula tokens; the library never evaluates them.
    pub formula_tokens: Vec<u8>,
    /// Checked area references decoded from the token stream.
    pub references: Vec<CellRef>,
}

/// One cached chart cell value.
#[derive(Clone, Debug, PartialEq)]
pub enum Value {
    /// Finite numeric value.
    Number(f64),
    /// Text value.
    Text(String),
    /// Explicit blank.
    Blank,
}

/// Cached value address and content.
#[derive(Clone, Debug, PartialEq)]
pub struct Cache {
    /// `SIIndex` cache identifier.
    pub cache_index: u16,
    /// Zero-based cached row.
    pub row: u16,
    /// Zero-based cached column.
    pub column: u8,
    /// BIFF number-format index stored with the cached cell.
    pub format: u16,
    /// Cached cell content.
    pub value: Value,
}

/// One chart data series.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Series {
    /// Category scalar kind.
    pub category_kind: DataKind,
    /// Declared category value count.
    pub category_count: u16,
    /// Declared numeric value count.
    pub value_count: u16,
    /// Declared bubble-size value count.
    pub bubble_count: u16,
    /// Zero-based chart-group order referenced by this series.
    pub chart_group: u16,
    /// Optional series name.
    pub name: Option<String>,
    /// Inert source links in record order.
    pub links: Vec<DataLink>,
}

impl Default for Series {
    fn default() -> Self {
        Self {
            category_kind: DataKind::Text,
            category_count: 0,
            value_count: 0,
            bubble_count: 0,
            chart_group: 0,
            name: None,
            links: Vec::new(),
        }
    }
}

/// Rendering family and family-specific BIFF settings for one chart group.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum GroupKind {
    /// Line chart.
    Line {
        /// Validated BIFF option bits.
        flags: u16,
    },
    /// Bar or column chart.
    Bar {
        /// Signed series overlap percentage.
        overlap: i16,
        /// Inter-series gap width.
        gap: u16,
        /// Validated BIFF option bits.
        flags: u16,
    },
    /// Area chart.
    Area {
        /// Validated BIFF option bits.
        flags: u16,
    },
    /// Pie or doughnut chart.
    Pie {
        /// First-slice rotation in degrees.
        rotation: u16,
        /// Doughnut hole percentage; zero selects a pie.
        hole_size: u16,
        /// Validated BIFF option bits.
        flags: u16,
    },
    /// Scatter or bubble chart.
    Scatter {
        /// Bubble-size percentage.
        bubble_size_percent: u16,
        /// BIFF bubble sizing mode.
        bubble_size_type: u16,
        /// Validated BIFF option bits.
        flags: u16,
    },
    /// Radar chart.
    Radar {
        /// Whether the radar area is filled.
        filled: bool,
        /// Validated BIFF option bits.
        flags: u16,
    },
    /// Surface chart.
    Surface {
        /// Validated BIFF option bits.
        flags: u16,
    },
}

/// Ordered chart group.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Group {
    /// Stable BIFF group order in the range `0..=9`.
    pub order: u16,
    /// Whether each series receives a distinct color.
    pub vary_colors: bool,
    /// Rendering family and settings.
    pub kind: GroupKind,
    /// Ordered drop, high-low, series, or leader lines with required formatting.
    pub lines: Vec<group::Line>,
    /// Complete up/down-bar collections in source order.
    pub drop_bars: Vec<group::DropBar>,
}

/// Concise chart family derived from its groups.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum Kind {
    /// No chart group is present.
    Empty,
    /// Line chart.
    Line,
    /// Bar or column chart.
    Bar,
    /// Area chart.
    Area,
    /// Pie or doughnut chart.
    Pie,
    /// Scatter or bubble chart.
    Scatter,
    /// Radar chart.
    Radar,
    /// Surface chart.
    Surface,
    /// Stock chart.
    Stock,
    /// Multiple chart groups.
    Combo,
}

/// Semantic role of a chart axis.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum AxisKind {
    /// Category axis, or the horizontal axis in a scatter chart.
    CategoryOrHorizontal,
    /// Value axis, or the vertical axis in a scatter chart.
    ValueOrVertical,
    /// Series axis for a 3-D chart.
    Series,
}

/// Numeric value-axis scale.
#[derive(Clone, Debug, PartialEq)]
pub struct Scale {
    /// Minimum scale value.
    pub minimum: f64,
    /// Maximum scale value.
    pub maximum: f64,
    /// Major unit.
    pub major: f64,
    /// Minor unit.
    pub minor: f64,
    /// Crossing value.
    pub crossing: f64,
    /// Validated BIFF option bits.
    pub flags: u16,
}

/// Axis tick and label settings.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Tick {
    /// Major tick-mark mode.
    pub major: u8,
    /// Minor tick-mark mode.
    pub minor: u8,
    /// Axis-label position mode.
    pub label_position: u8,
    /// Text-background mode.
    pub background: u8,
    /// BIFF color bytes.
    pub color: [u8; 4],
    /// Validated BIFF option bits.
    pub flags: u16,
}

/// Line role within an axis block.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum AxisLineKind {
    /// The axis itself.
    Axis,
    /// Major gridlines.
    MajorGridlines,
    /// Minor gridlines.
    MinorGridlines,
    /// Plot walls or floor.
    WallsOrFloor,
}

/// BIFF line styling.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct LineFormat {
    /// BIFF color bytes.
    pub color: [u8; 4],
    /// Line-pattern code.
    pub pattern: u16,
    /// Line-weight code.
    pub weight: i16,
    /// Validated BIFF option bits.
    pub flags: u16,
    /// Palette or automatic-color index.
    pub color_index: u16,
}

/// Formatting for one axis-line role.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct AxisLine {
    /// Semantic role of the line.
    pub kind: AxisLineKind,
    /// Required styling from the immediately following `LineFormat` record.
    pub format: LineFormat,
}

/// One chart axis.
#[derive(Clone, Debug, PartialEq)]
pub struct Axis {
    /// Semantic axis role.
    pub kind: AxisKind,
    /// Optional numeric scale.
    pub scale: Option<Scale>,
    /// Optional tick and label settings.
    pub tick: Option<Tick>,
    /// Ordered axis-line roles and their styling.
    pub lines: Vec<AxisLine>,
}

/// Chart legend geometry and placement.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Legend {
    /// Horizontal position in chart units.
    pub x: i32,
    /// Vertical position in chart units.
    pub y: i32,
    /// Width in chart units.
    pub width: i32,
    /// Height in chart units.
    pub height: i32,
    /// BIFF legend-position code.
    pub position: u8,
    /// BIFF legend-spacing code.
    pub spacing: u8,
    /// Validated BIFF option bits.
    pub flags: u16,
}

/// BIFF area fill styling.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct AreaFormat {
    /// Foreground BIFF color bytes.
    pub foreground: [u8; 4],
    /// Background BIFF color bytes.
    pub background: [u8; 4],
    /// Fill-pattern code.
    pub pattern: u16,
    /// Validated BIFF option bits.
    pub flags: u16,
    /// Foreground palette index.
    pub foreground_index: u16,
    /// Background palette index.
    pub background_index: u16,
}

/// Formatting record retained in chart record order.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Format {
    /// Line styling.
    Line(LineFormat),
    /// Area fill styling.
    Area(AreaFormat),
    /// Marker payload not yet elevated to a semantic model.
    Marker {
        /// Original `MarkerFormat` payload.
        data: Vec<u8>,
    },
    /// Per-point or per-series formatting selector.
    Data {
        /// Point index, or BIFF's all-points sentinel.
        point: u16,
        /// Series index.
        series: u16,
        /// Validated BIFF option bits.
        flags: u16,
    },
    /// Pie-slice explosion formatting.
    Pie(pie::Format),
}

/// Opaque data-label record retained for lossless rewriting.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Label {
    /// Supported data-label BIFF record identifier.
    pub record_type: u16,
    /// Original record payload.
    pub data: Vec<u8>,
}

/// Owned semantic BIFF8 chart model.
///
/// `Clone` is retained while transactional workbook mutation snapshots the
/// current model; chart serialization itself writes into one move-owned buffer.
#[derive(Clone, Debug, PartialEq)]
pub struct Chart {
    /// Horizontal chart origin in BIFF chart units.
    pub x: i32,
    /// Vertical chart origin in BIFF chart units.
    pub y: i32,
    /// Chart width in BIFF chart units.
    pub width: i32,
    /// Chart height in BIFF chart units.
    pub height: i32,
    /// Validated `ShtProps` option bits.
    pub sheet_properties: u32,
    /// Whether the chart contains a `PlotArea` marker.
    pub plot_area_present: bool,
    /// Optional chart title.
    pub title: Option<String>,
    /// Data series in record order.
    pub series: Vec<Series>,
    /// Chart groups in rendering order.
    pub groups: Vec<Group>,
    /// Axes in record order.
    pub axes: Vec<Axis>,
    /// Optional legend.
    pub legend: Option<Legend>,
    /// Cached values used when linked source cells are unavailable.
    pub cached_values: Vec<Cache>,
    /// Supported formatting records in source order.
    pub formatting: Vec<Format>,
    /// Data-label extension records in source order.
    pub data_labels: Vec<Label>,
    /// Records not interpreted by this implementation, retained byte-for-byte.
    pub unknown_records: Vec<Raw>,
}

impl Default for Chart {
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
            groups: vec![Group {
                order: 0,
                vary_colors: false,
                kind: GroupKind::Line { flags: 0 },
                lines: Vec::new(),
                drop_bars: Vec::new(),
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

impl Chart {
    /// Derives the concise chart family from the configured groups.
    pub fn kind(&self) -> Kind {
        match self.groups.as_slice() {
            [] => Kind::Empty,
            [group] => match &group.kind {
                GroupKind::Line { .. }
                    if !group.drop_bars.is_empty()
                        && group
                            .lines
                            .iter()
                            .any(|value| value.kind == line::Kind::HighLow) =>
                {
                    Kind::Stock
                },
                GroupKind::Line { .. } => Kind::Line,
                GroupKind::Bar { .. } => Kind::Bar,
                GroupKind::Area { .. } => Kind::Area,
                GroupKind::Pie { .. } => Kind::Pie,
                GroupKind::Scatter { .. } => Kind::Scatter,
                GroupKind::Radar { .. } => Kind::Radar,
                GroupKind::Surface { .. } => Kind::Surface,
            },
            _ => Kind::Combo,
        }
    }

    /// Checks resource bounds, invariants, flags, and inert cell references.
    pub fn validate(&self, limits: Limits) -> XlsResult<()> {
        validate_limits(limits)?;
        validate_sheet_properties(self.sheet_properties)?;
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
            if group.drop_bars.len() > 2 {
                return invalid(
                    DROP_BAR,
                    "a chart group permits at most two DropBar records",
                );
            }
            if !group.drop_bars.is_empty() && !matches!(group.kind, GroupKind::Line { .. }) {
                return invalid(DROP_BAR, "DropBar records require a line chart group");
            }
            let mut prior_line = None;
            for value in &group.lines {
                let current = value.kind;
                if prior_line.is_some_and(|prior| current <= prior) {
                    return invalid(
                        CRT_LINE,
                        "chart-group lines are duplicated or not strictly ordered",
                    );
                }
                prior_line = Some(current);
            }
            match group.kind {
                GroupKind::Area { flags } | GroupKind::Line { flags } if flags & !7 != 0 => {
                    return invalid(CHART_FORMAT, "area/line chart uses reserved flags");
                },
                GroupKind::Bar {
                    overlap,
                    gap,
                    flags,
                } if !(-100..=100).contains(&overlap) || gap > 500 || flags & !0xf != 0 => {
                    return invalid(BAR, "bar chart settings are outside BIFF bounds");
                },
                GroupKind::Pie {
                    rotation,
                    hole_size,
                    flags,
                } if rotation > 360 || hole_size > 90 || flags & !3 != 0 => {
                    return invalid(PIE, "pie/doughnut settings are out of range");
                },
                GroupKind::Radar { flags, .. } | GroupKind::Surface { flags }
                    if flags & !3 != 0 =>
                {
                    return invalid(CHART_FORMAT, "radar/surface chart uses reserved flags");
                },
                GroupKind::Scatter {
                    bubble_size_percent,
                    bubble_size_type,
                    flags,
                } if bubble_size_percent > 300
                    || !(1..=2).contains(&bubble_size_type)
                    || flags & !7 != 0 =>
                {
                    return invalid(SCATTER, "scatter chart settings are outside BIFF bounds");
                },
                _ => {},
            }
        }
        for series in &self.series {
            if usize::from(series.chart_group) >= self.groups.len() {
                return invalid(SER_TO_CRT, "series references a missing chart group");
            }
            for link in &series.links {
                validate_link(link, limits)?;
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
                    AxisLineKind::Axis => 0,
                    AxisLineKind::MajorGridlines => 1,
                    AxisLineKind::MinorGridlines => 2,
                    AxisLineKind::WallsOrFloor => 3,
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
            if record.data.len() > MAX_RECORD_BYTES {
                return invalid(
                    record.record_type,
                    "opaque BIFF record exceeds maximum length",
                );
            }
        }
        Ok(())
    }
}

/// Chart plus its current workbook host location.
#[derive(Clone, Debug, PartialEq)]
pub struct Entry {
    /// Stable location within the current workbook revision.
    pub location: Location,
    /// Owned chart model.
    pub chart: Chart,
}

#[derive(Clone)]
struct StoredChart {
    entry: Entry,
    #[cfg(test)]
    start: usize,
    #[cfg(test)]
    end: usize,
    #[cfg(test)]
    object: Option<(usize, usize)>,
}

const UNSUPPORTED_AUTHORING_REASON: &str = "fresh and replacement XLS chart authoring requires the complete Office-compatible BIFF chart grammar";
const UNSUPPORTED_EMBEDDED_MUTATION_REASON: &str = "embedded XLS chart mutation requires complete MsoDrawing/Continue, Obj/Continue, chart-substream, and OfficeArt drawing-group ownership";

fn unsupported_authoring<T>() -> XlsResult<T> {
    Err(litchi_ograph::Error::UnsupportedAuthoring {
        reason: UNSUPPORTED_AUTHORING_REASON,
    }
    .into())
}

fn unsupported_embedded_mutation<T>() -> XlsResult<T> {
    Err(litchi_ograph::Error::UnsupportedMutation {
        operation: "embedded XLS chart drawing",
        reason: UNSUPPORTED_EMBEDDED_MUTATION_REASON,
    }
    .into())
}

/// Transactional editor for existing BIFF8 chart substreams.
pub struct Editor {
    package: ObjectEditor,
    workbook_path: Vec<String>,
    workbook: Arc<[u8]>,
    limits: Limits,
    charts: Vec<StoredChart>,
}

impl Editor {
    /// Takes ownership of an XLS compound file and validates its chart inventory.
    pub fn open(bytes: Vec<u8>, limits: Limits) -> XlsResult<Self> {
        validate_limits(limits)?;
        let package = ObjectEditor::open(bytes, Targets::default(), ObjectLimits::default())?;
        let workbook_path = [vec!["Workbook".into()], vec!["Book".into()]]
            .into_iter()
            .find(|path| package.stream(path).is_some())
            .ok_or_else(|| XlsError::InvalidData("Workbook stream not found".into()))?;
        let workbook = package
            .stream_shared(&workbook_path)
            .ok_or_else(|| XlsError::InvalidData("selected Workbook stream disappeared".into()))?;
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

    /// Iterates borrowed chart entries in workbook drawing order.
    pub fn charts(&self) -> impl ExactSizeIterator<Item = &Entry> {
        self.charts.iter().map(|value| &value.entry)
    }

    /// Consume the editor and return the parsed chart inventory without persisting.
    pub fn into_charts(self) -> Vec<Entry> {
        self.charts.into_iter().map(|value| value.entry).collect()
    }

    /// Looks up a chart by worksheet name and semantic position.
    pub fn get(&self, selector: Selector<'_>) -> XlsResult<Option<&Chart>> {
        let Some(location) = self.resolve(selector)? else {
            return Ok(None);
        };
        Ok(self.at(&location))
    }

    /// Looks up a chart using a checked low-level host location.
    pub fn at(&self, location: &Location) -> Option<&Chart> {
        self.charts
            .iter()
            .find(|value| &value.entry.location == location)
            .map(|value| &value.entry.chart)
    }

    /// Refuses fresh embedded-chart authoring until its complete BIFF grammar is available.
    ///
    /// The editor is unchanged when this returns
    /// [`litchi_ograph::Error::UnsupportedAuthoring`] through [`XlsError::Graph`].
    pub fn add(&mut self, _sheet: &str, _chart: Chart) -> XlsResult<Location> {
        unsupported_authoring()
    }

    /// Refuses fresh embedded-chart authoring at a checked raw host location.
    ///
    /// The editor is unchanged when this returns
    /// [`litchi_ograph::Error::UnsupportedAuthoring`] through [`XlsError::Graph`].
    pub fn insert_at(
        &mut self,
        _sheet_index: usize,
        _object_id: u16,
        _index: usize,
        _chart: Chart,
    ) -> XlsResult<()> {
        unsupported_authoring()
    }

    #[cfg(test)]
    fn insert_fixture_at(
        &mut self,
        sheet_index: usize,
        object_id: u16,
        index: usize,
        chart: Chart,
    ) -> XlsResult<()> {
        let (_, sheets) = bindings(&self.workbook)?;
        let sheet = sheets
            .iter()
            .find(|value| value.index == sheet_index)
            .ok_or_else(|| invalid_error(BOUNDSHEET, "worksheet index was not found"))?;
        if sheet.kind != 0 {
            return invalid(BOUNDSHEET, "embedded charts require a worksheet tab");
        }
        if object_id == 0 || sheet_object_ids(&self.workbook, sheet)?.contains(&object_id) {
            return invalid(OBJ, "embedded chart object ID is zero or duplicated");
        }
        chart.validate(self.limits)?;
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
            Entry {
                location: Location::Embedded {
                    sheet_index,
                    object_id,
                },
                chart,
            },
        );
        self.commit_fixture(&original, desired)
    }

    /// Refuses fresh chart-sheet authoring until its complete BIFF grammar is available.
    ///
    /// The editor is unchanged when this returns
    /// [`litchi_ograph::Error::UnsupportedAuthoring`] through [`XlsError::Graph`].
    pub fn add_sheet(&mut self, _name: impl Into<String>, _chart: Chart) -> XlsResult<()> {
        unsupported_authoring()
    }

    /// Refuses fresh chart-sheet authoring at a checked raw tab index.
    ///
    /// The editor is unchanged when this returns
    /// [`litchi_ograph::Error::UnsupportedAuthoring`] through [`XlsError::Graph`].
    pub fn insert_sheet_at(
        &mut self,
        _index: usize,
        _name: impl Into<String>,
        _chart: Chart,
    ) -> XlsResult<()> {
        unsupported_authoring()
    }

    /// Remove a chart-sheet tab. References to that tab cause atomic failure.
    pub fn remove_sheet_at(&mut self, sheet_index: usize) -> XlsResult<Chart> {
        let (_, sheets) = bindings(&self.workbook)?;
        let sheet = sheets
            .iter()
            .find(|value| value.index == sheet_index)
            .ok_or_else(|| invalid_error(BOUNDSHEET, "sheet index was not found"))?;
        if sheet.kind != 2 {
            return invalid(BOUNDSHEET, "selected tab is not a chart sheet");
        }
        let chart_index = self
            .charts
            .iter()
            .position(|value| value.entry.location == Location::ChartSheet { sheet_index })
            .ok_or_else(|| invalid_error(CHART, "chart sheet has no chart"))?;
        let order = (0..sheets.len())
            .filter(|value| *value != sheet_index)
            .map(Some)
            .collect::<Vec<_>>();
        let workbook = rewrite_sheet_directory(&self.workbook, &order, None)?;
        let mut previous = self.install_workbook(workbook)?;
        if chart_index >= previous.len() {
            return invalid(CHART, "removed chart-sheet inventory changed unexpectedly");
        }
        Ok(previous.swap_remove(chart_index).entry.chart)
    }

    /// Reorders workbook tabs by Unicode case-insensitive tab names.
    pub fn reorder_sheets(&mut self, order: &[&str]) -> XlsResult<()> {
        let (_, sheets) = bindings(&self.workbook)?;
        if order.len() != sheets.len() {
            return invalid(BOUNDSHEET, "sheet reorder must contain every tab");
        }
        let mut indexes = Vec::new();
        indexes
            .try_reserve(order.len())
            .map_err(|_| XlsError::InvalidData("could not allocate sheet reorder".into()))?;
        let mut seen = HashSet::new();
        for name in order {
            let sheet = sheets
                .iter()
                .find(|value| names_equal(&value.name, name))
                .ok_or_else(|| invalid_error(BOUNDSHEET, "sheet reorder name was not found"))?;
            if !seen.insert(sheet.index) {
                return invalid(BOUNDSHEET, "sheet reorder repeats a tab name");
            }
            indexes.push(sheet.index);
        }
        self.reorder_sheets_at(&indexes)
    }

    /// Reorders all workbook tabs by checked previous zero-based indexes.
    pub fn reorder_sheets_at(&mut self, order: &[usize]) -> XlsResult<()> {
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
        self.install_workbook(workbook).map(drop)
    }

    /// Refuses replacement authoring until its complete BIFF grammar is available.
    ///
    /// The editor is unchanged when this returns
    /// [`litchi_ograph::Error::UnsupportedAuthoring`] through [`XlsError::Graph`].
    pub fn replace(&mut self, _selector: Selector<'_>, _chart: Chart) -> XlsResult<()> {
        unsupported_authoring()
    }

    /// Refuses replacement authoring at a checked low-level host location.
    ///
    /// The editor is unchanged when this returns
    /// [`litchi_ograph::Error::UnsupportedAuthoring`] through [`XlsError::Graph`].
    pub fn replace_at(&mut self, _location: &Location, _chart: Chart) -> XlsResult<()> {
        unsupported_authoring()
    }

    /// Removes a chart sheet transactionally and refuses embedded-chart removal.
    ///
    /// Embedded charts participate in the worksheet OfficeArt drawing graph;
    /// until that complete ownership is modeled, the editor returns
    /// [`litchi_ograph::Error::UnsupportedMutation`] without mutation.
    pub fn remove(&mut self, selector: Selector<'_>) -> XlsResult<Chart> {
        let location = self
            .resolve(selector)?
            .ok_or_else(|| invalid_error(CHART, "chart selector was not found"))?;
        self.remove_at(&location)
    }

    /// Removes a chart sheet using a checked low-level host location.
    ///
    /// An existing embedded location is validated and then refused atomically
    /// until its complete OfficeArt drawing ownership can be rewritten.
    pub fn remove_at(&mut self, location: &Location) -> XlsResult<Chart> {
        if let Location::ChartSheet { sheet_index } = location {
            return self.remove_sheet_at(*sheet_index);
        }
        if self
            .charts
            .iter()
            .all(|value| &value.entry.location != location)
        {
            return Err(invalid_error(CHART, "chart location was not found"));
        }
        unsupported_embedded_mutation()
    }

    /// Validates embedded-chart order on a named worksheet.
    ///
    /// The current identity order is a no-op. A structural reorder is refused
    /// atomically until complete OfficeArt drawing ownership is modeled.
    pub fn reorder(&mut self, sheet: &str, order: &[usize]) -> XlsResult<()> {
        let (_, sheets) = bindings(&self.workbook)?;
        let sheet = sheets
            .iter()
            .find(|value| names_equal(&value.name, sheet))
            .ok_or_else(|| invalid_error(BOUNDSHEET, "worksheet name was not found"))?;
        let ids = self
            .charts
            .iter()
            .filter_map(|value| match value.entry.location {
                Location::Embedded {
                    sheet_index,
                    object_id,
                } if sheet_index == sheet.index => Some(object_id),
                _ => None,
            })
            .collect::<Vec<_>>();
        if order.len() != ids.len() {
            return invalid(CHART, "chart reorder must contain every chart");
        }
        let mut seen = HashSet::new();
        let mut object_ids = Vec::new();
        object_ids
            .try_reserve(order.len())
            .map_err(|_| XlsError::InvalidData("could not allocate chart reorder".into()))?;
        for index in order {
            let id = ids
                .get(*index)
                .copied()
                .filter(|_| seen.insert(*index))
                .ok_or_else(|| {
                    invalid_error(CHART, "chart reorder index is invalid or repeated")
                })?;
            object_ids.push(id);
        }
        self.reorder_at(sheet.index, &object_ids)
    }

    /// Validates embedded-chart order using checked worksheet and Obj identifiers.
    ///
    /// The current identity order is a no-op. A structural reorder is refused
    /// atomically until complete OfficeArt drawing ownership is modeled.
    pub fn reorder_at(&mut self, sheet_index: usize, object_ids: &[u16]) -> XlsResult<()> {
        let (_, sheets) = bindings(&self.workbook)?;
        let sheet = sheets
            .iter()
            .find(|value| value.index == sheet_index)
            .ok_or_else(|| invalid_error(BOUNDSHEET, "worksheet index was not found"))?;
        if sheet.kind != 0 {
            return invalid(BOUNDSHEET, "embedded charts require a worksheet tab");
        }
        let slots = self
            .charts
            .iter()
            .enumerate()
            .filter_map(|(i, value)| match value.entry.location {
                Location::Embedded {
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
        let mut available = slots.clone();
        let mut current = Vec::new();
        current
            .try_reserve_exact(slots.len())
            .map_err(|_| XlsError::InvalidData("could not allocate chart reorder".into()))?;
        for index in &slots {
            let object_id = self
                .charts
                .get(*index)
                .and_then(|value| match value.entry.location {
                    Location::Embedded { object_id, .. } => Some(object_id),
                    Location::ChartSheet { .. } => None,
                })
                .ok_or_else(|| invalid_error(CHART, "chart reorder slot is invalid"))?;
            current.push(object_id);
        }
        for id in object_ids {
            let position = available
                .iter()
                .position(|index| {
                    self.charts.get(*index).is_some_and(|value| {
                        matches!(
                            value.entry.location,
                            Location::Embedded { object_id, .. } if object_id == *id
                        )
                    })
                })
                .ok_or_else(|| {
                    invalid_error(CHART, "reorder contains an unknown or repeated object ID")
                })?;
            available.remove(position);
        }
        if current == object_ids {
            return Ok(());
        }
        unsupported_embedded_mutation()
    }

    /// Consumes the editor and returns the rewritten compound-file allocation.
    pub fn finish(self) -> XlsResult<Vec<u8>> {
        self.package.finish().map_err(Into::into)
    }

    #[cfg(test)]
    fn commit_fixture(&mut self, original: &[StoredChart], desired: Vec<Entry>) -> XlsResult<()> {
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
        let workbook: Arc<[u8]> = workbook.into();
        self.package
            .put_stream_shared(&self.workbook_path, Arc::clone(&workbook))?;
        self.workbook = workbook;
        self.charts = reparsed;
        Ok(())
    }

    fn install_workbook(&mut self, workbook: Vec<u8>) -> XlsResult<Vec<StoredChart>> {
        if workbook.len() > self.limits.max_workbook_bytes {
            return invalid(CHART, "rewritten Workbook exceeds limit");
        }
        let charts = parse_workbook_charts(&workbook, self.limits)?;
        let workbook: Arc<[u8]> = workbook.into();
        self.package
            .put_stream_shared(&self.workbook_path, Arc::clone(&workbook))?;
        self.workbook = workbook;
        Ok(std::mem::replace(&mut self.charts, charts))
    }

    fn resolve(&self, selector: Selector<'_>) -> XlsResult<Option<Location>> {
        let (_, sheets) = bindings(&self.workbook)?;
        match selector {
            Selector::Sheet(name) => Ok(sheets
                .iter()
                .find(|sheet| sheet.kind == 2 && names_equal(&sheet.name, name))
                .map(|sheet| Location::ChartSheet {
                    sheet_index: sheet.index,
                })),
            Selector::Embedded { sheet, index } => {
                let Some(sheet) = sheets
                    .iter()
                    .find(|value| value.kind == 0 && names_equal(&value.name, sheet))
                else {
                    return Ok(None);
                };
                Ok(self
                    .charts
                    .iter()
                    .filter(|value| value.entry.location.sheet_index() == sheet.index)
                    .filter_map(|value| match value.entry.location {
                        Location::Embedded { .. } => Some(value.entry.location.clone()),
                        Location::ChartSheet { .. } => None,
                    })
                    .nth(index))
            },
        }
    }
}

/// The Obj identifier assigned to the chart embedded in a test-only workbook.
#[cfg(test)]
const GENERATED_CHART_OBJECT_ID: u16 = 1;
/// BIFF record type marking the workbook-globals substream.
#[cfg(test)]
const BOF_WORKBOOK_GLOBALS: u16 = 0x0005;
/// BIFF record type marking a worksheet substream.
#[cfg(test)]
const BOF_WORKSHEET: u16 = 0x0010;
/// Sheet name of the single worksheet hosting a test-only embedded chart.
#[cfg(test)]
const GENERATED_SHEET_NAME: &str = "Sheet1";

/// Refuses fresh standalone BIFF8 chart-workbook authoring.
///
/// This public entry point returns [`litchi_ograph::Error::UnsupportedAuthoring`]
/// through [`XlsError::Graph`] until the complete Office-compatible chart
/// grammar is implemented.
pub fn build_workbook(_chart: Chart, _limits: Limits) -> XlsResult<Vec<u8>> {
    unsupported_authoring()
}

/// Builds the abbreviated workbook used only to exercise the private parser.
#[cfg(test)]
pub(crate) fn build_workbook_fixture(chart: Chart, limits: Limits) -> XlsResult<Vec<u8>> {
    validate_limits(limits)?;
    let mut package = litchi_cfb::OleWriter::new();
    package.create_stream(&["Workbook"], &minimal_workbook_stream()?)?;
    let mut bytes = std::io::Cursor::new(Vec::new());
    package.write_to(&mut bytes)?;
    let mut editor = Editor::open(bytes.into_inner(), limits)?;
    editor.insert_fixture_at(0, GENERATED_CHART_OBJECT_ID, 0, chart)?;
    editor.finish()
}

/// A minimal one-worksheet BIFF8 `Workbook` stream accepted by the chart
/// editor: workbook globals with a single `BoundSheet` directory entry
/// followed by an empty worksheet substream.
#[cfg(test)]
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
        *logical
            .get_mut(sheet.index)
            .ok_or_else(|| invalid_error(BOUNDSHEET, "sheet directory index is invalid"))? =
            Some((sheet.kind, input[sheet.start..sheet.end].to_vec()));
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
        let bound = old_bounds
            .get(index)
            .cloned()
            .ok_or_else(|| invalid_error(BOUNDSHEET, "BoundSheet index is invalid"))?;
        tabs.push((Some(index), bound, stream));
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
            *old_to_new
                .get_mut(*old)
                .ok_or_else(|| invalid_error(BOUNDSHEET, "sheet permutation index is invalid"))? =
                Some(new);
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
                old_to_new.get(old).copied().flatten().ok_or_else(|| {
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
        let value = match old_index {
            Some(index) => old
                .get(*index)
                .copied()
                .ok_or_else(|| invalid_error(RR_TAB_ID, "RRTabId index is invalid"))?,
            None => next,
        };
        output.extend(value.to_le_bytes());
    }
    Ok(output)
}
fn names_equal(left: &str, right: &str) -> bool {
    left.chars()
        .flat_map(char::to_lowercase)
        .eq(right.chars().flat_map(char::to_lowercase))
}
#[cfg(test)]
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
#[cfg(test)]
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

fn parse_workbook_charts(input: &[u8], limits: Limits) -> XlsResult<Vec<StoredChart>> {
    let (_, sheets) = bindings(input)?;
    let mut output = Vec::new();
    for sheet in &sheets {
        if sheet.kind == 2 {
            let bytes = &input[sheet.start..sheet.end];
            let chart_ref = GraphChartRef::with_limits(bytes, chart_scan_limits(limits))
                .map_err(|error| graph_error(BOF, error))?;
            if chart_ref.kind() != GraphChartKind::Excel {
                return invalid(BOF, "chart sheet uses a non-Excel chart grammar");
            }
            let chart = parse_chart(chart_ref.as_bytes(), limits)?;
            output.push(StoredChart {
                entry: Entry {
                    location: Location::ChartSheet {
                        sheet_index: sheet.index,
                    },
                    chart,
                },
                #[cfg(test)]
                start: sheet.start,
                #[cfg(test)]
                end: sheet.end,
                #[cfg(test)]
                object: None,
            });
            continue;
        }
        if sheet.kind != 0 {
            continue;
        }
        let bytes = &input[sheet.start..sheet.end];
        let records = ranges(bytes)?;
        let mut chart_objects = Vec::new();
        let mut used = HashSet::new();
        for value in records {
            if value.kind == OBJ
                && let Some(id) = parse_chart_object(&bytes[value.body_start..value.body_end])?
            {
                let start = sheet
                    .start
                    .checked_add(value.start)
                    .ok_or_else(|| XlsError::InvalidData("object start offset overflow".into()))?;
                let end = sheet
                    .start
                    .checked_add(value.end)
                    .ok_or_else(|| XlsError::InvalidData("object end offset overflow".into()))?;
                chart_objects.push((id, start, end));
            }
        }
        let charts = GraphCharts::with_limits(bytes, chart_scan_limits(limits))
            .map_err(|error| graph_error(BOF, error))?;
        for chart_ref in charts {
            let chart_ref = chart_ref.map_err(|error| graph_error(BOF, error))?;
            if chart_ref.kind() != GraphChartKind::Excel {
                return invalid(BOF, "embedded chart uses a non-Excel chart grammar");
            }
            let start = sheet
                .start
                .checked_add(chart_ref.offset())
                .ok_or_else(|| XlsError::InvalidData("chart start offset overflow".into()))?;
            let end = start
                .checked_add(chart_ref.as_bytes().len())
                .ok_or_else(|| XlsError::InvalidData("chart end offset overflow".into()))?;
            #[cfg(not(test))]
            let _ = end;
            let object = chart_objects
                .iter()
                .rev()
                .find(|(id, _, object_end)| *object_end <= start && !used.contains(id))
                .copied()
                .ok_or_else(|| {
                    invalid_error(OBJ, "embedded chart BOF has no preceding chart Obj/FtCmo")
                })?;
            used.insert(object.0);
            let chart = parse_chart(chart_ref.as_bytes(), limits)?;
            output.push(StoredChart {
                entry: Entry {
                    location: Location::Embedded {
                        sheet_index: sheet.index,
                        object_id: object.0,
                    },
                    chart,
                },
                #[cfg(test)]
                start,
                #[cfg(test)]
                end,
                #[cfg(test)]
                object: Some((object.1, object.2)),
            });
        }
    }
    if output.len() > limits.max_charts {
        return invalid(CHART, "chart count exceeds limit");
    }
    Ok(output)
}

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

fn parse_chart(input: &[u8], limits: Limits) -> XlsResult<Chart> {
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
    let mut cache_index = 0u16;
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
                    .ok_or_else(|| XlsError::InvalidData("chart nesting overflow".into()))?;
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
            0x1051 => {
                if let Some(series) = current_series {
                    let link = parse_link(data, limits)?;
                    chart
                        .series
                        .get_mut(series)
                        .ok_or_else(|| invalid_error(0x1051, "Series index is invalid"))?
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
                        .ok_or_else(|| XlsError::InvalidData("DropBar nesting overflow".into()))?,
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
                cache_index = u16_at(data, 0)?;
            },
            BLANK => {
                exact(data, 6, BLANK)?;
                chart.cached_values.push(Cache {
                    cache_index,
                    row: u16_at(data, 0)?,
                    column: u8::try_from(u16_at(data, 2)?)
                        .map_err(|_| invalid_error(BLANK, "cached column exceeds BIFF8 grid"))?,
                    format: u16_at(data, 4)?,
                    value: Value::Blank,
                });
            },
            NUMBER => {
                exact(data, 14, NUMBER)?;
                chart.cached_values.push(Cache {
                    cache_index,
                    row: u16_at(data, 0)?,
                    column: u8::try_from(u16_at(data, 2)?)
                        .map_err(|_| invalid_error(NUMBER, "cached column exceeds BIFF8 grid"))?,
                    format: u16_at(data, 4)?,
                    value: Value::Number(f64_at(data, 6)?),
                });
            },
            LABEL => {
                if data.len() < 8 {
                    return invalid(LABEL, "cached Label is truncated");
                }
                chart.cached_values.push(Cache {
                    cache_index,
                    row: u16_at(data, 0)?,
                    column: u8::try_from(u16_at(data, 2)?)
                        .map_err(|_| invalid_error(LABEL, "cached column exceeds BIFF8 grid"))?,
                    format: u16_at(data, 4)?,
                    value: Value::Text(parse_biff8_string(&data[6..])?),
                });
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

fn validate_link(link: &DataLink, limits: Limits) -> XlsResult<()> {
    if link.formula_tokens.len() > limits.max_formula_bytes {
        return invalid(0x1051, "BRAI formula length exceeds the configured limit");
    }
    if link.source == Source::Automatic && !link.formula_tokens.is_empty() {
        return invalid(0x1051, "automatic BRAI must have an empty formula");
    }
    if parse_chart_references(&link.formula_tokens)? != link.references {
        return invalid(
            0x1051,
            "BRAI references do not match its inert formula tokens",
        );
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

fn parse_link(data: &[u8], limits: Limits) -> XlsResult<DataLink> {
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
    let value = DataLink {
        role: match data[0] {
            0 => Role::Name,
            1 => Role::Values,
            2 => Role::Categories,
            3 => Role::Bubbles,
            _ => return invalid(0x1051, "BRAI role is invalid"),
        },
        source: match data[1] {
            0 => Source::Automatic,
            1 => Source::Literal,
            2 => Source::Cells,
            _ => return invalid(0x1051, "BRAI source kind is invalid"),
        },
        unlinked_number_format: flags & 1 != 0,
        number_format: u16_at(data, 4)?,
        formula_tokens,
        references,
    };
    validate_link(&value, limits)?;
    Ok(value)
}

fn parse_chart_references(tokens: &[u8]) -> XlsResult<Vec<CellRef>> {
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

fn parse_group(kind: u16, data: &[u8]) -> XlsResult<GroupKind> {
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
fn serialize_chart(chart: &Chart, limits: Limits) -> XlsResult<Vec<u8>> {
    chart.validate(limits)?;
    if !chart.unknown_records.is_empty() {
        return Err(XlsError::UnsafeEdit(
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
            let mut data = vec![link.role as u8, link.source as u8];
            data.extend(u16::from(link.unlinked_number_format).to_le_bytes());
            data.extend(link.number_format.to_le_bytes());
            data.extend(
                u16::try_from(link.formula_tokens.len())
                    .map_err(|_| XlsError::InvalidData("chart formula exceeds u16".into()))?
                    .to_le_bytes(),
            );
            data.extend(&link.formula_tokens);
            push_record(&mut out, 0x1051, &data)?;
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
        if active_cache != Some(value.cache_index) {
            push_record(&mut out, SI_INDEX, &value.cache_index.to_le_bytes())?;
            active_cache = Some(value.cache_index);
        }
        match &value.value {
            Value::Number(number) => {
                if !number.is_finite() {
                    return invalid(NUMBER, "cached chart number must be finite");
                }
                let mut data = Vec::new();
                data.extend(value.row.to_le_bytes());
                data.extend(u16::from(value.column).to_le_bytes());
                data.extend(value.format.to_le_bytes());
                data.extend(number.to_le_bytes());
                push_record(&mut out, NUMBER, &data)?;
            },
            Value::Text(text) => {
                let mut data = Vec::new();
                data.extend(value.row.to_le_bytes());
                data.extend(u16::from(value.column).to_le_bytes());
                data.extend(value.format.to_le_bytes());
                data.extend(biff8_string(text)?);
                push_record(&mut out, LABEL, &data)?;
            },
            Value::Blank => {
                let mut data = Vec::with_capacity(6);
                data.extend(value.row.to_le_bytes());
                data.extend(u16::from(value.column).to_le_bytes());
                data.extend(value.format.to_le_bytes());
                push_record(&mut out, BLANK, &data)?;
            },
        }
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
fn write_group(out: &mut GraphEncoder, group: &Group) -> XlsResult<()> {
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

#[cfg(test)]
fn rewrite_workbook_charts(
    input: &[u8],
    original: &[StoredChart],
    desired: &[Entry],
    limits: Limits,
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
        let changed = existing.iter().map(|v| &v.entry).collect::<Vec<_>>() != wanted;
        if sheet.kind == 2 {
            if !changed {
                output.extend_from_slice(&input[sheet.start..sheet.end]);
                continue;
            }
            let current = existing
                .first()
                .ok_or_else(|| invalid_error(CHART, "chart sheet has no parsed chart"))?;
            let replacement = wanted
                .iter()
                .find(|v| v.location == current.entry.location)
                .ok_or_else(|| invalid_error(CHART, "chart sheet removal is not supported"))?;
            output.extend(serialize_chart(&replacement.chart, limits)?);
            continue;
        }
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
                Location::Embedded { object_id, .. } => object_id,
                _ => return invalid(CHART, "chart sheet cannot be embedded"),
            };
            output.extend(chart_object_record(object_id)?);
            output.extend(serialize_chart(&value.chart, limits)?);
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
#[derive(Clone)]
struct Sheet {
    index: usize,
    start: usize,
    end: usize,
    kind: u8,
    name: String,
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
            refs.push((
                value.start + 4,
                u32_at(data, 0)? as usize,
                data[5],
                bound_sheet_name(data)?,
            ));
        }
    }
    let mut physical = refs
        .iter()
        .enumerate()
        .map(|(index, (_, start, kind, name))| (index, *start, *kind, name.clone()))
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
        .map(|(slot, (index, start, kind, name))| Sheet {
            index: *index,
            start: *start,
            end: physical.get(slot + 1).map_or(input.len(), |v| v.1),
            kind: *kind,
            name: name.clone(),
        })
        .collect();
    Ok((
        refs.into_iter().map(|(p, o, _, _)| (p, o)).collect(),
        sheets,
    ))
}
fn ranges(input: &[u8]) -> XlsResult<Vec<Range>> {
    ranges_with(input, BiffLimits::default().max_records)
}
fn ranges_with(input: &[u8], max_records: usize) -> XlsResult<Vec<Range>> {
    let mut out = Vec::new();
    let biff_limits = BiffLimits {
        max_records,
        max_input_bytes: input.len().max(1),
        ..BiffLimits::default()
    };
    let records = Records::with_limits(input, biff_limits)
        .map_err(|error| XlsError::InvalidData(error.to_string()))?;
    for record in records {
        let record = record.map_err(|error| XlsError::InvalidData(error.to_string()))?;
        let start = record.offset();
        let body_start = start
            .checked_add(4)
            .ok_or_else(|| XlsError::InvalidData("BIFF body offset overflow".into()))?;
        let end = start
            .checked_add(record.encoded().len())
            .ok_or_else(|| XlsError::InvalidData("BIFF record length overflow".into()))?;
        out.try_reserve(1)
            .map_err(|_| XlsError::InvalidData("could not allocate BIFF ranges".into()))?;
        out.push(Range {
            start,
            end,
            kind: record.kind().get(),
            body_start,
            body_end: end,
        });
    }
    Ok(out)
}
#[cfg(test)]
fn sheet_object_ids(input: &[u8], sheet: &Sheet) -> XlsResult<HashSet<u16>> {
    let bytes = input
        .get(sheet.start..sheet.end)
        .ok_or_else(|| invalid_error(BOUNDSHEET, "worksheet range is out of bounds"))?;
    let mut ids = HashSet::new();
    for value in ranges(bytes)? {
        if value.kind != OBJ {
            continue;
        }
        if let Some((_, id)) = parse_object(&bytes[value.body_start..value.body_end])? {
            ids.insert(id);
        }
    }
    Ok(ids)
}
fn parse_chart_object(data: &[u8]) -> XlsResult<Option<u16>> {
    Ok(parse_object(data)?.and_then(|(kind, id)| (kind == 5).then_some(id)))
}
fn parse_object(data: &[u8]) -> XlsResult<Option<(u16, u16)>> {
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
            return Ok(Some((u16_at(body, 0)?, u16_at(body, 2)?)));
        }
        offset = end;
    }
    Ok(None)
}
#[cfg(test)]
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
#[cfg(test)]
fn chart_bof() -> Vec<u8> {
    bof_body(0x0020)
}
#[cfg(test)]
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
fn parse_line_format(data: &[u8]) -> XlsResult<LineFormat> {
    exact(data, 12, LINE_FORMAT)?;
    Ok(LineFormat {
        color: array_at(data, 0)?,
        pattern: u16_at(data, 4)?,
        weight: i16_at(data, 6)?,
        flags: u16_at(data, 8)?,
        color_index: u16_at(data, 10)?,
    })
}
#[cfg(test)]
fn write_line(v: &LineFormat) -> Vec<u8> {
    let mut d = v.color.to_vec();
    d.extend(v.pattern.to_le_bytes());
    d.extend(v.weight.to_le_bytes());
    d.extend(v.flags.to_le_bytes());
    d.extend(v.color_index.to_le_bytes());
    d
}
fn shared_line(value: &LineFormat) -> format::Line {
    format::Line {
        color: value.color,
        pattern: value.pattern,
        weight: value.weight,
        flags: value.flags,
        color_index: value.color_index,
    }
}
#[cfg(test)]
fn shared_line_bytes(value: &format::Line) -> Vec<u8> {
    write_line(&LineFormat {
        color: value.color,
        pattern: value.pattern,
        weight: value.weight,
        flags: value.flags,
        color_index: value.color_index,
    })
}
fn parse_area_format(data: &[u8]) -> XlsResult<AreaFormat> {
    exact(data, 16, AREA_FORMAT)?;
    Ok(AreaFormat {
        foreground: array_at(data, 0)?,
        background: array_at(data, 4)?,
        pattern: u16_at(data, 8)?,
        flags: u16_at(data, 10)?,
        foreground_index: u16_at(data, 12)?,
        background_index: u16_at(data, 14)?,
    })
}
#[cfg(test)]
fn write_area(v: &AreaFormat) -> Vec<u8> {
    let mut d = v.foreground.to_vec();
    d.extend(v.background);
    d.extend(v.pattern.to_le_bytes());
    d.extend(v.flags.to_le_bytes());
    d.extend(v.foreground_index.to_le_bytes());
    d.extend(v.background_index.to_le_bytes());
    d
}
fn shared_area(value: &AreaFormat) -> format::Area {
    format::Area {
        foreground: value.foreground,
        background: value.background,
        pattern: value.pattern,
        flags: value.flags,
        foreground_index: value.foreground_index,
        background_index: value.background_index,
    }
}
#[cfg(test)]
fn shared_area_bytes(value: &format::Area) -> Vec<u8> {
    write_area(&AreaFormat {
        foreground: value.foreground,
        background: value.background,
        pattern: value.pattern,
        flags: value.flags,
        foreground_index: value.foreground_index,
        background_index: value.background_index,
    })
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
#[cfg(test)]
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
#[cfg(test)]
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
fn chart_scan_limits(limits: Limits) -> GraphLimits {
    GraphLimits {
        max_workbook_bytes: limits.max_workbook_bytes,
        max_charts: limits.max_charts,
        max_chart_records: limits.max_records_per_chart,
        max_series: limits.max_series,
        max_groups: limits.max_groups,
        max_axes: limits.max_axes,
        max_formula_bytes: limits.max_formula_bytes,
        max_cached_values: limits.max_cached_values,
        max_unknown_bytes: limits.max_unknown_bytes,
        biff: BiffLimits {
            max_records: limits.max_records_per_chart,
            max_input_bytes: limits.max_workbook_bytes,
            max_output_bytes: limits.max_workbook_bytes,
            ..BiffLimits::default()
        },
        ..GraphLimits::default()
    }
}
#[cfg(test)]
fn chart_encoder(limits: Limits) -> XlsResult<GraphEncoder> {
    GraphEncoder::with_limits(BiffLimits {
        max_records: limits.max_records_per_chart,
        max_output_bytes: limits.max_workbook_bytes,
        ..BiffLimits::default()
    })
    .map_err(|error| frame_error(CHART, error))
}
fn push_record(out: &mut GraphEncoder, kind: u16, data: &[u8]) -> XlsResult<()> {
    out.push(RecordKind::from_wire(kind), data)
        .map_err(|error| frame_error(kind, error))
}
fn record(kind: u16, data: &[u8]) -> XlsResult<Vec<u8>> {
    let output_bytes = data
        .len()
        .checked_add(4)
        .ok_or_else(|| XlsError::InvalidData("BIFF record size overflow".into()))?;
    let limits = BiffLimits {
        max_records: 1,
        max_output_bytes: output_bytes.max(1),
        ..BiffLimits::default()
    };
    let mut out = GraphEncoder::with_limits(limits).map_err(|error| frame_error(kind, error))?;
    push_record(&mut out, kind, data)?;
    Ok(out.finish())
}
#[cfg(test)]
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
            | BLANK
            | NUMBER
            | LABEL
            | BEGIN
            | END
            | SHT_PROPS
            | AXES_USED
            | AXIS_PARENT
    )
}
fn validate_limits(v: Limits) -> XlsResult<()> {
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

fn validate_sheet_properties(flags: u32) -> XlsResult<()> {
    let blank = (flags >> 16) & 0xff;
    let always_auto = flags & (1 << 4) != 0;
    let manual_plot = flags & (1 << 3) != 0;
    if flags & 0xff00_ffe0 != 0 || blank > 2 || (always_auto && !manual_plot) {
        return invalid(
            SHT_PROPS,
            "ShtProps reserved bits, blank mode, or plot-area flags are invalid",
        );
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
fn graph_error(kind: u16, error: litchi_ograph::Error) -> XlsError {
    invalid_error(kind, error.to_string())
}
fn frame_error(kind: u16, error: litchi_biff::Error) -> XlsError {
    invalid_error(kind, error.to_string())
}
fn invalid<T>(kind: u16, message: impl Into<String>) -> XlsResult<T> {
    Err(invalid_error(kind, message))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn abbreviated_test_fixture_exercises_chart_parser() {
        let mut chart = Chart {
            title: Some("Sales".into()),
            ..Default::default()
        };
        chart.series.push(Series {
            category_count: 2,
            value_count: 2,
            links: vec![DataLink {
                role: Role::Values,
                source: Source::Cells,
                unlinked_number_format: false,
                number_format: 0,
                formula_tokens: vec![0x3b, 0, 0, 1, 0, 2, 0, 1, 0, 1, 0],
                references: vec![CellRef {
                    extern_sheet_index: 0,
                    first_row: 1,
                    last_row: 2,
                    first_column: 1,
                    last_column: 1,
                }],
            }],
            ..Default::default()
        });
        chart.cached_values.push(Cache {
            cache_index: 0,
            row: 0,
            column: 0,
            format: 0,
            value: Value::Number(42.0),
        });
        let bytes = build_workbook_fixture(chart, Limits::default()).unwrap();
        let editor = Editor::open(bytes, Limits::default()).unwrap();
        let mut charts = editor.charts();
        assert_eq!(charts.len(), 1);
        let chart = charts.next().unwrap();
        assert_eq!(
            chart.location,
            Location::Embedded {
                sheet_index: 0,
                object_id: GENERATED_CHART_OBJECT_ID
            }
        );
        assert_eq!(chart.chart.title.as_deref(), Some("Sales"));
        assert_eq!(chart.chart.series.len(), 1);
        assert!(
            chart
                .chart
                .cached_values
                .iter()
                .any(|v| v.value == Value::Number(42.0))
        );
    }

    #[test]
    fn public_authoring_refuses_atomically_with_typed_error() {
        fn assert_unsupported(error: XlsError) {
            assert!(matches!(
                error,
                XlsError::Graph(litchi_ograph::Error::UnsupportedAuthoring { .. })
            ));
        }

        let original = build_workbook_fixture(Chart::default(), Limits::default()).unwrap();
        let location = Location::Embedded {
            sheet_index: 0,
            object_id: GENERATED_CHART_OBJECT_ID,
        };

        let mut editor = Editor::open(original.clone(), Limits::default()).unwrap();
        assert_unsupported(
            editor
                .replace_at(&location, Chart::default())
                .expect_err("replacement must be refused"),
        );
        assert_eq!(editor.finish().unwrap(), original);

        assert_unsupported(
            build_workbook(Chart::default(), Limits::default())
                .expect_err("fresh workbook authoring must be refused"),
        );
    }

    #[test]
    fn embedded_identity_reorder_is_exact_and_removal_is_atomic_refusal() {
        let original = build_workbook_fixture(Chart::default(), Limits::default()).unwrap();

        let mut reordered = Editor::open(original.clone(), Limits::default()).unwrap();
        reordered.reorder("Sheet1", &[0]).unwrap();
        assert_eq!(reordered.finish().unwrap(), original);

        let mut removed = Editor::open(original.clone(), Limits::default()).unwrap();
        assert!(matches!(
            removed
                .remove(Selector::Embedded {
                    sheet: "Sheet1",
                    index: 0,
                })
                .expect_err("embedded removal must be refused"),
            XlsError::Graph(litchi_ograph::Error::UnsupportedMutation { .. })
        ));
        assert_eq!(removed.finish().unwrap(), original);
    }

    #[test]
    fn editor_and_package_share_the_workbook_stream_allocation() {
        let bytes = build_workbook_fixture(Chart::default(), Limits::default()).unwrap();
        let editor = Editor::open(bytes, Limits::default()).unwrap();
        let captured = editor.package.stream_shared(&editor.workbook_path).unwrap();
        assert!(Arc::ptr_eq(&captured, &editor.workbook));
    }

    #[test]
    fn generated_combo_round_trips_and_refuses_unplaced_opaque_records() {
        let mut chart = Chart::default();
        chart.groups.push(Group {
            order: 1,
            vary_colors: true,
            kind: GroupKind::Pie {
                rotation: 45,
                hole_size: 50,
                flags: 0,
            },
            lines: Vec::new(),
            drop_bars: Vec::new(),
        });
        chart.series.push(Series {
            chart_group: 1,
            ..Default::default()
        });
        let bytes = serialize_chart(&chart, Limits::default()).unwrap();
        let parsed = parse_chart(&bytes, Limits::default()).unwrap();
        assert_eq!(parsed.kind(), Kind::Combo);

        chart.unknown_records.push(Raw {
            record_type: 0x7777,
            data: b"opaque".to_vec(),
        });
        assert!(matches!(
            serialize_chart(&chart, Limits::default()),
            Err(XlsError::UnsafeEdit(_))
        ));
    }

    #[test]
    fn axis_line_format_and_blank_cache_round_trip_once() {
        let format = LineFormat {
            color: [1, 2, 3, 4],
            pattern: 1,
            weight: 2,
            flags: 0,
            color_index: 8,
        };
        let mut chart = Chart::default();
        chart.axes.push(Axis {
            kind: AxisKind::CategoryOrHorizontal,
            scale: None,
            tick: None,
            lines: vec![AxisLine {
                kind: AxisLineKind::Axis,
                format: format.clone(),
            }],
        });
        chart.cached_values.push(Cache {
            cache_index: 3,
            row: 4,
            column: 5,
            format: 9,
            value: Value::Blank,
        });

        let bytes = serialize_chart(&chart, Limits::default()).unwrap();
        let parsed = parse_chart(&bytes, Limits::default()).unwrap();
        assert_eq!(parsed.axes[0].lines[0].format, format);
        assert!(
            parsed
                .formatting
                .iter()
                .all(|value| !matches!(value, Format::Line(_)))
        );
        assert!(parsed.cached_values.iter().any(|value| {
            value.cache_index == 3
                && value.row == 4
                && value.column == 5
                && value.format == 9
                && value.value == Value::Blank
        }));
    }

    #[test]
    fn group_lines_and_drop_bars_round_trip_without_collapsing_kinds() {
        let line_format = format::Line {
            color: [1, 2, 3, 4],
            pattern: 1,
            weight: 2,
            flags: 0,
            color_index: 8,
        };
        let area_format = format::Area {
            foreground: [5, 6, 7, 8],
            background: [9, 10, 11, 12],
            pattern: 1,
            flags: 0,
            foreground_index: 9,
            background_index: 10,
        };
        let mut chart = Chart::default();
        chart.groups[0].lines.extend([
            group::Line {
                kind: line::Kind::HighLow,
                format: line_format,
            },
            group::Line {
                kind: line::Kind::Series,
                format: line_format,
            },
        ]);
        chart.groups[0].drop_bars.push(group::DropBar {
            gap: group::Gap::new(257).expect("valid DropBar gap"),
            line: line_format,
            area: area_format,
        });

        let bytes = serialize_chart(&chart, Limits::default()).unwrap();
        let kinds = ranges(&bytes)
            .expect("framed chart")
            .into_iter()
            .map(|record| record.kind)
            .collect::<Vec<_>>();
        for (index, kind) in kinds.iter().enumerate() {
            if *kind == CRT_LINE {
                assert_eq!(kinds.get(index + 1), Some(&LINE_FORMAT));
            }
        }
        assert!(
            kinds
                .windows(5)
                .any(|window| { window == [DROP_BAR, BEGIN, LINE_FORMAT, AREA_FORMAT, END] })
        );
        let parsed = parse_chart(&bytes, Limits::default()).unwrap();
        assert_eq!(parsed.kind(), Kind::Stock);
        assert_eq!(parsed.groups[0].lines, chart.groups[0].lines);
        assert_eq!(parsed.groups[0].drop_bars, chart.groups[0].drop_bars);
    }

    #[test]
    fn sheet_properties_and_group_numeric_domains_are_strict() {
        let limits = Limits::default();
        let mut chart = Chart::default();
        for blank in 0..=2 {
            chart.sheet_properties = blank << 16;
            chart.validate(limits).expect("valid blank mode");
        }
        chart.sheet_properties = 3 << 16;
        assert!(chart.validate(limits).is_err());
        chart.sheet_properties = 1 << 2;
        chart
            .validate(limits)
            .expect("fNotSizeWith is a defined ShtProps bit");
        chart.sheet_properties = 1 << 4;
        assert!(chart.validate(limits).is_err());
        chart.sheet_properties = (1 << 4) | (1 << 3);
        chart.validate(limits).expect("paired plot-area flags");

        chart.groups[0].kind = GroupKind::Bar {
            overlap: 101,
            gap: 150,
            flags: 0,
        };
        assert!(chart.validate(limits).is_err());
        chart.groups[0].kind = GroupKind::Scatter {
            bubble_size_percent: 301,
            bubble_size_type: 1,
            flags: 0,
        };
        assert!(chart.validate(limits).is_err());
        chart.groups[0].kind = GroupKind::Scatter {
            bubble_size_percent: 100,
            bubble_size_type: 3,
            flags: 0,
        };
        assert!(chart.validate(limits).is_err());
    }

    #[test]
    fn malformed_nesting_formula_and_axis_are_rejected() {
        let mut bytes = record(BOF, &chart_bof()).unwrap();
        bytes.extend(record(CHART, &[0; 16]).unwrap());
        bytes.extend(record(END, &[]).unwrap());
        bytes.extend(record(EOF, &[]).unwrap());
        assert!(parse_chart(&bytes, Limits::default()).is_err());
        let mut chart = Chart::default();
        chart.series.push(Series {
            links: vec![DataLink {
                role: Role::Values,
                source: Source::Cells,
                unlinked_number_format: false,
                number_format: 0,
                formula_tokens: vec![0x3b, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0],
                references: vec![CellRef {
                    extern_sheet_index: 2,
                    first_row: 0,
                    last_row: 0,
                    first_column: 300,
                    last_column: 300,
                }],
            }],
            ..Default::default()
        });
        assert!(chart.validate(Limits::default()).is_err());
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
