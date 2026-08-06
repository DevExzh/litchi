use std::collections::HashSet;

use litchi_biff::MAX_RECORD_BYTES;
use litchi_ograph::chart::group;
use litchi_ograph::record::{line, pie};

use super::codec::validate_link;
use super::wire::*;
use crate::{Error, Result};

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

/// A bounded, inert chart formula.
///
/// The XLS owner intentionally exposes only the canonical single-cell and
/// rectangular `PtgRef3d`/`PtgArea3d` forms here. Other ChartParsedFormula
/// operands remain readable as raw link bytes, but are not accepted by this
/// mutation facade because their graph effects cannot yet be proven.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Formula {
    tokens: Vec<u8>,
    references: Vec<CellRef>,
}

impl Formula {
    /// Constructs an absolute single-cell reference.
    pub fn cell(reference: CellRef) -> Result<Self> {
        if reference.first_row != reference.last_row
            || reference.first_column != reference.last_column
        {
            return invalid(
                BRAI,
                "single-cell chart formula has a rectangular reference",
            );
        }
        Self::from_reference(reference, false)
    }

    /// Constructs an absolute rectangular cell-range reference.
    pub fn range(reference: CellRef) -> Result<Self> {
        Self::from_reference(reference, true)
    }

    /// Parses one canonical chart formula under the configured formula bound.
    ///
    /// Empty formulas are retained for inspection, but a data-link mutation
    /// requires a non-empty supported cell reference.
    pub fn parse(tokens: impl Into<Vec<u8>>, limits: Limits) -> Result<Self> {
        let tokens = tokens.into();
        if tokens.len() > limits.max_formula_bytes {
            return invalid(BRAI, "chart formula exceeds the configured limit");
        }
        let references = super::codec::parse_chart_references(&tokens)?;
        match tokens.len() {
            0 => {},
            7 if tokens[0] == 0x3a && u16_at(&tokens, 5)? & 0xc000 == 0 => {},
            11 if tokens[0] == 0x3b
                && u16_at(&tokens, 7)? & 0xc000 == 0
                && u16_at(&tokens, 9)? & 0xc000 == 0 => {},
            _ => {
                return invalid(
                    BRAI,
                    "chart formula uses an unsupported or opaque operand sequence",
                );
            },
        }
        if !tokens.is_empty() && references.is_empty() {
            return invalid(BRAI, "chart formula has no supported cell reference");
        }
        let formula = Self { tokens, references };
        for reference in &formula.references {
            validate_reference(reference)?;
        }
        Ok(formula)
    }

    /// Original inert formula tokens.
    #[must_use]
    pub fn tokens(&self) -> &[u8] {
        &self.tokens
    }

    /// Checked cell references represented by this formula.
    #[must_use]
    pub fn references(&self) -> &[CellRef] {
        &self.references
    }

    /// Number of cells covered by the represented references.
    pub fn cell_count(&self) -> Result<usize> {
        self.references.iter().try_fold(0usize, |count, reference| {
            let rows = usize::from(
                reference
                    .last_row
                    .checked_sub(reference.first_row)
                    .ok_or_else(|| {
                        Error::InvalidData("chart formula row order is invalid".into())
                    })?,
            )
            .checked_add(1)
            .ok_or_else(|| Error::InvalidData("chart formula row count overflow".into()))?;
            let columns = usize::from(
                reference
                    .last_column
                    .checked_sub(reference.first_column)
                    .ok_or_else(|| {
                        Error::InvalidData("chart formula column order is invalid".into())
                    })?,
            )
            .checked_add(1)
            .ok_or_else(|| Error::InvalidData("chart formula column count overflow".into()))?;
            count
                .checked_add(rows.checked_mul(columns).ok_or_else(|| {
                    Error::InvalidData("chart formula cell count overflow".into())
                })?)
                .ok_or_else(|| Error::InvalidData("chart formula cell count overflow".into()))
        })
    }

    fn from_reference(reference: CellRef, range: bool) -> Result<Self> {
        validate_reference(&reference)?;
        let mut tokens = Vec::with_capacity(if range { 11 } else { 7 });
        tokens.push(if range { 0x3b } else { 0x3a });
        tokens.extend(reference.extern_sheet_index.to_le_bytes());
        tokens.extend(reference.first_row.to_le_bytes());
        if range {
            tokens.extend(reference.last_row.to_le_bytes());
        }
        tokens.extend(reference.first_column.to_le_bytes());
        if range {
            tokens.extend(reference.last_column.to_le_bytes());
        }
        Ok(Self {
            tokens,
            references: vec![reference],
        })
    }

    fn into_parts(self) -> (Vec<u8>, Vec<CellRef>) {
        (self.tokens, self.references)
    }
}

fn validate_reference(reference: &CellRef) -> Result<()> {
    if reference.first_row > reference.last_row
        || reference.first_column > reference.last_column
        || reference.last_column > 255
    {
        return invalid(
            BRAI,
            "chart formula reference is outside the BIFF8 grid or is reversed",
        );
    }
    Ok(())
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

impl Role {
    /// The four roles used by a regular Excel series, in BIFF order.
    pub const ALL: [Self; 4] = [Self::Name, Self::Values, Self::Categories, Self::Bubbles];
}

/// Data-cache section selected by an `SIIndex` record.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[repr(u16)]
pub enum CacheKind {
    /// Series values or vertical values for scatter/bubble charts.
    Values = 1,
    /// Category labels or horizontal values for scatter/bubble charts.
    Categories = 2,
    /// Bubble-size values.
    Bubbles = 3,
}

impl CacheKind {
    /// Decodes the only three legal BIFF `SIIndex` values.
    pub const fn from_wire(value: u16) -> Option<Self> {
        match value {
            1 => Some(Self::Values),
            2 => Some(Self::Categories),
            3 => Some(Self::Bubbles),
            _ => None,
        }
    }

    /// Wire value written to `SIIndex`.
    pub const fn wire(self) -> u16 {
        self as u16
    }

    /// Series role represented by this cache section.
    pub const fn role(self) -> Role {
        match self {
            Self::Values => Role::Values,
            Self::Categories => Role::Categories,
            Self::Bubbles => Role::Bubbles,
        }
    }
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

impl DataLink {
    /// Creates an automatically generated link with no formula payload.
    pub const fn automatic(role: Role) -> Self {
        Self {
            role,
            source: Source::Automatic,
            unlinked_number_format: false,
            number_format: 0,
            formula_tokens: Vec::new(),
            references: Vec::new(),
        }
    }

    /// Creates a worksheet-cell link from a checked inert formula.
    pub fn cells(role: Role, formula: Formula) -> Self {
        let (formula_tokens, references) = formula.into_parts();
        Self {
            role,
            source: Source::Cells,
            unlinked_number_format: false,
            number_format: 0,
            formula_tokens,
            references,
        }
    }

    /// Returns the formula as a checked typed view.
    pub fn formula(&self, limits: Limits) -> Result<Formula> {
        Formula::parse(self.formula_tokens.clone(), limits)
    }
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
    /// `SIIndex` data section.
    pub kind: CacheKind,
    /// Zero-based point index encoded by the cache cell.
    pub point: u16,
    /// Zero-based series index encoded by the cache cell.
    pub series: u8,
    /// BIFF number-format index stored with the cached cell.
    pub format: u16,
    /// Cached cell content.
    pub value: Value,
}

impl Cache {
    /// Creates one typed cached chart value.
    pub fn new(kind: CacheKind, series: u8, point: u16, format: u16, value: Value) -> Result<Self> {
        validate_cache_value(&value)?;
        Ok(Self {
            kind,
            point,
            series,
            format,
            value,
        })
    }
}

fn validate_cache_value(value: &Value) -> Result<()> {
    match value {
        Value::Number(number) if !number.is_finite() => {
            invalid(NUMBER, "cached chart number must be finite")
        },
        Value::Text(text) if text.encode_utf16().count() > 255 => {
            invalid(LABEL, "cached chart label exceeds 255 UTF-16 code units")
        },
        Value::Number(_) | Value::Text(_) | Value::Blank => Ok(()),
    }
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

impl Series {
    /// Returns the first data link for a semantic series role.
    ///
    /// Parsed charts can retain incomplete or duplicated links for
    /// compatibility; callers that require the MS-XLS four-link invariant
    /// should use [`Chart::validate_semantics`](Chart::validate_semantics)
    /// before relying on this lookup as a complete view.
    #[must_use]
    pub fn link(&self, role: Role) -> Option<&DataLink> {
        self.links.iter().find(|link| link.role == role)
    }

    /// Replaces the formula of one existing worksheet-cell link.
    ///
    /// The link role, source kind, series shape, and declared cell count stay
    /// fixed. Adding, removing, duplicating, or retargeting a link is outside
    /// this narrow safe mutation boundary.
    pub fn set_formula(&mut self, role: Role, formula: Formula) -> Result<()> {
        let mut link_index = None;
        for (index, link) in self.links.iter().enumerate() {
            if link.role == role {
                if link_index.is_some() {
                    return Err(Error::UnsafeEdit(
                        "series formula link is duplicated".into(),
                    ));
                }
                link_index = Some(index);
            }
        }
        if role != Role::Name {
            let expected = match role {
                Role::Values => self.value_count,
                Role::Categories => self.category_count,
                Role::Bubbles => self.bubble_count,
                Role::Name => 0,
            };
            let actual = u16::try_from(formula.cell_count()?)
                .map_err(|_| Error::InvalidData("chart formula cell count exceeds u16".into()))?;
            if actual != expected {
                return invalid(
                    BRAI,
                    "retargeted chart formula changes the declared series cardinality",
                );
            }
        }
        let link_index =
            link_index.ok_or_else(|| Error::UnsafeEdit("series formula link is missing".into()))?;
        let link = self
            .links
            .get(link_index)
            .ok_or_else(|| Error::InvalidData("series formula link disappeared".into()))?;
        if link.source != Source::Cells {
            return Err(Error::UnsafeEdit(
                "only worksheet-cell chart links can be retargeted".into(),
            ));
        }
        if link.references.len() != formula.references().len()
            || link
                .references
                .iter()
                .zip(formula.references())
                .any(|(old, new)| {
                    old.last_row.checked_sub(old.first_row)
                        != new.last_row.checked_sub(new.first_row)
                        || old.last_column.checked_sub(old.first_column)
                            != new.last_column.checked_sub(new.first_column)
                })
        {
            return Err(Error::UnsafeEdit(
                "retargeted chart formula changes the reference shape".into(),
            ));
        }
        let link = self
            .links
            .get_mut(link_index)
            .ok_or_else(|| Error::InvalidData("series formula link disappeared".into()))?;
        let (formula_tokens, references) = formula.into_parts();
        link.formula_tokens = formula_tokens;
        link.references = references;
        Ok(())
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
    /// Replaces one existing series formula without changing chart topology.
    pub fn set_formula(&mut self, series: usize, role: Role, formula: Formula) -> Result<()> {
        self.set_formula_with(series, role, formula, Limits::default())
    }

    pub(crate) fn set_formula_with(
        &mut self,
        series: usize,
        role: Role,
        formula: Formula,
        limits: Limits,
    ) -> Result<()> {
        let value = self
            .series
            .get_mut(series)
            .ok_or_else(|| Error::InvalidData("chart series index is out of range".into()))?;
        let previous = value.clone();
        value.set_formula(role, formula)?;
        if let Err(error) = self.validate_semantics(limits) {
            *self
                .series
                .get_mut(series)
                .ok_or_else(|| Error::InvalidData("chart series disappeared".into()))? = previous;
            return Err(error);
        }
        Ok(())
    }

    /// Replaces one existing cached cell while preserving its section and
    /// coordinate. Cache insertion, removal, and reclassification are refused.
    pub fn set_cache(
        &mut self,
        kind: CacheKind,
        series: u8,
        point: u16,
        format: u16,
        value: Value,
    ) -> Result<()> {
        self.set_cache_with(kind, series, point, format, value, Limits::default())
    }

    pub(crate) fn set_cache_with(
        &mut self,
        kind: CacheKind,
        series: u8,
        point: u16,
        format: u16,
        value: Value,
        limits: Limits,
    ) -> Result<()> {
        let mut matches = self.cached_values.iter().enumerate().filter(|(_, cache)| {
            cache.kind == kind && cache.series == series && cache.point == point
        });
        let index = matches
            .next()
            .map(|(index, _)| index)
            .ok_or_else(|| Error::UnsafeEdit("chart cache cell is missing".into()))?;
        if matches.next().is_some() {
            return Err(Error::UnsafeEdit("chart cache cell is duplicated".into()));
        }
        validate_cache_value(&value)?;
        let previous = self.cached_values[index].clone();
        self.cached_values[index].format = format;
        self.cached_values[index].value = value;
        if let Err(error) = self.validate_semantics(limits) {
            self.cached_values[index] = previous;
            return Err(error);
        }
        Ok(())
    }

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
    pub fn validate(&self, limits: Limits) -> Result<()> {
        validate_limits(limits)?;
        validate_sheet_properties(self.sheet_properties)?;
        if self.series.len() > limits.max_series
            || self.groups.len() > limits.max_groups
            || self.axes.len() > limits.max_axes
            || self.cached_values.len() > limits.max_cached_values
            || self.unknown_records.len() > limits.max_records_per_chart
        {
            return invalid(CHART, "chart resource limit exceeded");
        }
        let data_link_count = self
            .series
            .iter()
            .try_fold(0usize, |count, series| {
                count.checked_add(series.links.len())
            })
            .ok_or_else(|| Error::InvalidData("chart data-link count overflow".into()))?;
        if data_link_count > limits.max_records_per_chart {
            return invalid(BRAI, "chart data-link count exceeds the record limit");
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
        for cache in &self.cached_values {
            validate_cache_value(&cache.value)?;
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
            .ok_or_else(|| Error::InvalidData("chart unknown-record size overflow".into()))?;
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
