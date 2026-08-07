//! Package-neutral sparkline values and invariants.

use std::num::FpCategory;

use bitflags::bitflags;
use litchi_sheet::sparkline::{AxisType, EmptyCells, SparklineType};
use thiserror::Error;

const WIRE_FORMULA_MAX: usize = 16_384;
const ROWS: u32 = litchi_sheet::ROWS;
const COLUMNS: u32 = litchi_sheet::COLUMNS;

/// Result of reading, validating, or writing an XLSB sparkline block.
pub type Result<T> = std::result::Result<T, Error>;

/// A strict XLSB sparkline failure.
#[derive(Debug, Error)]
#[non_exhaustive]
pub enum Error {
    /// The BIFF12 record framing was malformed or exceeded its payload budget.
    #[error("invalid sparkline BIFF12 framing: {0}")]
    Wire(#[from] crate::raw::Error),
    /// A configured limit was zero or outside the wire-supported domain.
    #[error("invalid sparkline limit {resource}={value}: {reason}")]
    InvalidLimit {
        /// Name of the rejected resource.
        resource: &'static str,
        /// Rejected value.
        value: usize,
        /// Required domain.
        reason: &'static str,
    },
    /// A finite resource budget was exceeded.
    #[error("sparkline {resource} count/size {actual} exceeds limit {maximum}")]
    Limit {
        /// Bounded resource.
        resource: &'static str,
        /// Observed or requested count/size.
        actual: usize,
        /// Configured maximum.
        maximum: usize,
    },
    /// A required allocation could not be reserved.
    #[error("unable to reserve memory for sparkline {resource}")]
    Allocation {
        /// Allocation purpose.
        resource: &'static str,
    },
    /// The record grammar did not match the sparkline block ABNF.
    #[error("invalid sparkline record sequence: expected {expected}, found {found}")]
    Record {
        /// Required next record.
        expected: &'static str,
        /// Actual record kind or end-of-input.
        found: String,
    },
    /// A begin/end delimiter carried a forbidden payload.
    #[error("{record} delimiter payload must be empty, found {length} bytes")]
    Delimiter {
        /// Delimiter name.
        record: &'static str,
        /// Unexpected payload length.
        length: usize,
    },
    /// A typed record payload violated a normative field constraint.
    #[error("invalid {record}: {reason}")]
    Value {
        /// Record or structure name.
        record: &'static str,
        /// Rejected condition.
        reason: String,
    },
    /// A group or groups collection was empty where the grammar requires 1+.
    #[error("{collection} must contain at least one item")]
    Empty {
        /// Empty collection name.
        collection: &'static str,
    },
}

/// Finite resource limits for one XLSB sparkline block.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    groups: usize,
    per_group: usize,
    total: usize,
    formula_tokens: usize,
    formula_ancillary: usize,
    record_bytes: usize,
    block_bytes: usize,
    worksheet_bytes: usize,
}

impl Limits {
    /// Safe defaults for worksheet-scale processing.
    pub const DEFAULT: Self = Self {
        groups: 230,
        per_group: 230,
        total: 52_900,
        formula_tokens: WIRE_FORMULA_MAX,
        formula_ancillary: 64 * 1024,
        record_bytes: 1024 * 1024,
        block_bytes: 8 * 1024 * 1024,
        worksheet_bytes: 512 * 1024 * 1024,
    };

    /// Construct an explicitly bounded policy.
    pub fn new(
        groups: usize,
        per_group: usize,
        total: usize,
        formula_tokens: usize,
        formula_ancillary: usize,
        record_bytes: usize,
        block_bytes: usize,
        worksheet_bytes: usize,
    ) -> Result<Self> {
        let value = Self {
            groups,
            per_group,
            total,
            formula_tokens,
            formula_ancillary,
            record_bytes,
            block_bytes,
            worksheet_bytes,
        };
        value.validate()?;
        Ok(value)
    }

    /// Maximum groups in the optional worksheet block.
    #[must_use]
    pub const fn groups(self) -> usize {
        self.groups
    }

    /// Maximum sparklines in one group.
    #[must_use]
    pub const fn per_group(self) -> usize {
        self.per_group
    }

    /// Maximum sparklines across all groups.
    #[must_use]
    pub const fn total(self) -> usize {
        self.total
    }

    /// Maximum `rgce` bytes in one formula.
    #[must_use]
    pub const fn formula_tokens(self) -> usize {
        self.formula_tokens
    }

    /// Maximum retained `rgcb` bytes in one formula.
    #[must_use]
    pub const fn formula_ancillary(self) -> usize {
        self.formula_ancillary
    }

    /// Maximum payload bytes in one BIFF12 sparkline record.
    #[must_use]
    pub const fn record_bytes(self) -> usize {
        self.record_bytes
    }

    /// Maximum encoded bytes in one complete sparkline block.
    #[must_use]
    pub const fn block_bytes(self) -> usize {
        self.block_bytes
    }

    /// Maximum retained bytes in the complete worksheet source stream that
    /// contains the sparkline block.
    #[must_use]
    pub const fn worksheet_bytes(self) -> usize {
        self.worksheet_bytes
    }

    /// Replace the group-count budget.
    pub fn with_groups(mut self, maximum: usize) -> Result<Self> {
        self.groups = maximum;
        self.validate()?;
        Ok(self)
    }

    /// Replace the per-group sparkline budget.
    pub fn with_per_group(mut self, maximum: usize) -> Result<Self> {
        self.per_group = maximum;
        self.validate()?;
        Ok(self)
    }

    /// Replace the aggregate sparkline budget.
    pub fn with_total(mut self, maximum: usize) -> Result<Self> {
        self.total = maximum;
        self.validate()?;
        Ok(self)
    }

    /// Replace the per-formula token budget.
    pub fn with_formula_tokens(mut self, maximum: usize) -> Result<Self> {
        self.formula_tokens = maximum;
        self.validate()?;
        Ok(self)
    }

    /// Replace the per-formula ancillary budget.
    pub fn with_formula_ancillary(mut self, maximum: usize) -> Result<Self> {
        self.formula_ancillary = maximum;
        self.validate()?;
        Ok(self)
    }

    /// Replace the per-record payload budget.
    pub fn with_record_bytes(mut self, maximum: usize) -> Result<Self> {
        self.record_bytes = maximum;
        self.validate()?;
        Ok(self)
    }

    /// Replace the complete-block byte budget.
    pub fn with_block_bytes(mut self, maximum: usize) -> Result<Self> {
        self.block_bytes = maximum;
        self.validate()?;
        Ok(self)
    }

    /// Replace the complete worksheet-source byte budget.
    pub fn with_worksheet_bytes(mut self, maximum: usize) -> Result<Self> {
        self.worksheet_bytes = maximum;
        self.validate()?;
        Ok(self)
    }

    pub(crate) fn validate(self) -> Result<()> {
        for (resource, value) in [
            ("groups", self.groups),
            ("sparklines per group", self.per_group),
            ("total sparklines", self.total),
            ("formula token bytes", self.formula_tokens),
            ("formula ancillary bytes", self.formula_ancillary),
            ("record payload bytes", self.record_bytes),
            ("block bytes", self.block_bytes),
            ("worksheet source bytes", self.worksheet_bytes),
        ] {
            if value == 0 {
                return Err(Error::InvalidLimit {
                    resource,
                    value,
                    reason: "must be nonzero",
                });
            }
        }
        if self.formula_tokens > WIRE_FORMULA_MAX {
            return Err(Error::InvalidLimit {
                resource: "formula token bytes",
                value: self.formula_tokens,
                reason: "must not exceed the FRTParsedFormula wire maximum 16384",
            });
        }
        if self.block_bytes > self.worksheet_bytes {
            return Err(Error::InvalidLimit {
                resource: "block bytes",
                value: self.block_bytes,
                reason: "must not exceed the complete worksheet-source byte limit",
            });
        }
        Ok(())
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// Allowed single-token formula kinds from `[MS-XLSB]` sections 2.4.228 and
/// 2.4.806.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FormulaKind {
    /// A workbook defined name (`PtgName`).
    Name,
    /// An external defined name (`PtgNameX`).
    ExternalName,
    /// A single-cell 3-D reference (`PtgRef3d`).
    Reference3d,
    /// A one-dimensional 3-D area (`PtgArea3d`).
    Area3d,
}

/// Structurally decoded coordinates of a `PtgRef3d` token.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Reference {
    row: u32,
    column: u32,
    row_relative: bool,
    column_relative: bool,
}

impl Reference {
    /// Zero-based source row.
    #[must_use]
    pub const fn row(self) -> u32 {
        self.row
    }

    /// Zero-based source column.
    #[must_use]
    pub const fn column(self) -> u32 {
        self.column
    }

    /// Whether the row coordinate carries `fRwRel`.
    #[must_use]
    pub const fn row_relative(self) -> bool {
        self.row_relative
    }

    /// Whether the column coordinate carries `fColRel`.
    #[must_use]
    pub const fn column_relative(self) -> bool {
        self.column_relative
    }
}

/// Structurally decoded coordinates of a `PtgArea3d` token.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Area {
    row_first: u32,
    row_last: u32,
    column_first: u32,
    column_last: u32,
    row_first_relative: bool,
    row_last_relative: bool,
    column_first_relative: bool,
    column_last_relative: bool,
}

impl Area {
    /// First zero-based row.
    #[must_use]
    pub const fn row_first(self) -> u32 {
        self.row_first
    }

    /// Last zero-based row.
    #[must_use]
    pub const fn row_last(self) -> u32 {
        self.row_last
    }

    /// First zero-based column.
    #[must_use]
    pub const fn column_first(self) -> u32 {
        self.column_first
    }

    /// Last zero-based column.
    #[must_use]
    pub const fn column_last(self) -> u32 {
        self.column_last
    }

    /// First coordinate `fRwRel` state.
    #[must_use]
    pub const fn row_first_relative(self) -> bool {
        self.row_first_relative
    }

    /// Last coordinate `fRwRel` state.
    #[must_use]
    pub const fn row_last_relative(self) -> bool {
        self.row_last_relative
    }

    /// First coordinate `fColRel` state.
    #[must_use]
    pub const fn column_first_relative(self) -> bool {
        self.column_first_relative
    }

    /// Last coordinate `fColRel` state.
    #[must_use]
    pub const fn column_last_relative(self) -> bool {
        self.column_last_relative
    }
}

/// One inert, wire-exact FRT formula.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Formula {
    kind: FormulaKind,
    rgce: Box<[u8]>,
    rgcb: Box<[u8]>,
}

impl Formula {
    /// Validate and retain one formula with default resource limits.
    pub fn new(rgce: Vec<u8>, rgcb: Vec<u8>) -> Result<Self> {
        Self::with_limits(rgce, rgcb, Limits::DEFAULT)
    }

    /// Validate and retain one formula with explicit resource limits.
    pub fn with_limits(rgce: Vec<u8>, rgcb: Vec<u8>, limits: Limits) -> Result<Self> {
        limits.validate()?;
        let kind = validate_formula(&rgce, &rgcb, limits)?;
        Ok(Self {
            kind,
            rgce: rgce.into_boxed_slice(),
            rgcb: rgcb.into_boxed_slice(),
        })
    }

    /// Construct a workbook-name token without resolving the name table.
    pub fn name(index: u32) -> Result<Self> {
        if index == 0 {
            return Err(value("FRTFormula", "PtgName index must be one-based"));
        }
        let mut rgce = Vec::new();
        rgce.try_reserve_exact(5)
            .map_err(|_| allocation("formula token"))?;
        rgce.push(0x23);
        rgce.extend_from_slice(&index.to_le_bytes());
        Self::new(rgce, Vec::new())
    }

    /// Return the structurally validated token family.
    #[must_use]
    pub const fn kind(&self) -> FormulaKind {
        self.kind
    }

    /// Lend the exact `rgce` token bytes.
    #[must_use]
    pub fn tokens(&self) -> &[u8] {
        &self.rgce
    }

    /// Lend the exact bounded `rgcb` ancillary bytes.
    #[must_use]
    pub fn ancillary(&self) -> &[u8] {
        &self.rgcb
    }

    /// Return the one-based name index for `PtgName` or `PtgNameX` without
    /// resolving the workbook/external name table.
    #[must_use]
    pub fn name_index(&self) -> Option<u32> {
        match self.kind {
            FormulaKind::Name => Some(u32::from_le_bytes([
                self.rgce[1],
                self.rgce[2],
                self.rgce[3],
                self.rgce[4],
            ])),
            FormulaKind::ExternalName => Some(u32::from_le_bytes([
                self.rgce[3],
                self.rgce[4],
                self.rgce[5],
                self.rgce[6],
            ])),
            FormulaKind::Reference3d | FormulaKind::Area3d => None,
        }
    }

    /// Return the unresolved `XtiIndex` for 3-D and external-name tokens.
    #[must_use]
    pub fn ixti(&self) -> Option<u16> {
        match self.kind {
            FormulaKind::ExternalName | FormulaKind::Reference3d | FormulaKind::Area3d => {
                Some(u16::from_le_bytes([self.rgce[1], self.rgce[2]]))
            },
            FormulaKind::Name => None,
        }
    }

    /// Decode a `PtgRef3d` coordinate without resolving its `XtiIndex`.
    #[must_use]
    pub fn reference(&self) -> Option<Reference> {
        if self.kind != FormulaKind::Reference3d {
            return None;
        }
        let row = u32::from_le_bytes([self.rgce[3], self.rgce[4], self.rgce[5], self.rgce[6]]);
        let column = u16::from_le_bytes([self.rgce[7], self.rgce[8]]);
        Some(Reference {
            row,
            column: u32::from(column & 0x3fff),
            row_relative: column & 0x4000 != 0,
            column_relative: column & 0x8000 != 0,
        })
    }

    /// Decode a `PtgArea3d` range without resolving its `XtiIndex`.
    #[must_use]
    pub fn area(&self) -> Option<Area> {
        if self.kind != FormulaKind::Area3d {
            return None;
        }
        let row_first =
            u32::from_le_bytes([self.rgce[3], self.rgce[4], self.rgce[5], self.rgce[6]]);
        let row_last =
            u32::from_le_bytes([self.rgce[7], self.rgce[8], self.rgce[9], self.rgce[10]]);
        let first = u16::from_le_bytes([self.rgce[11], self.rgce[12]]);
        let last = u16::from_le_bytes([self.rgce[13], self.rgce[14]]);
        Some(Area {
            row_first,
            row_last,
            column_first: u32::from(first & 0x3fff),
            column_last: u32::from(last & 0x3fff),
            row_first_relative: first & 0x4000 != 0,
            row_last_relative: last & 0x4000 != 0,
            column_first_relative: first & 0x8000 != 0,
            column_last_relative: last & 0x8000 != 0,
        })
    }

    pub(crate) fn from_slices(rgce: &[u8], rgcb: &[u8], limits: Limits) -> Result<Self> {
        let kind = validate_formula(rgce, rgcb, limits)?;
        Ok(Self {
            kind,
            rgce: copy_boxed(rgce, "formula token")?,
            rgcb: copy_boxed(rgcb, "formula ancillary data")?,
        })
    }
}

/// A non-automatic BIFF12 color family.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ColorType {
    /// Indexed palette color.
    Palette,
    /// Explicit red-green-blue-alpha color.
    Rgb,
    /// Theme color.
    Theme,
}

/// A validated, wire-exact `BrtColor` value.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Color {
    raw: [u8; 8],
}

impl Color {
    /// Validate exact `BrtColor` bytes. Automatic colors are forbidden for all
    /// eight sparkline slots.
    pub fn from_raw(raw: [u8; 8]) -> Result<Self> {
        validate_color(raw)?;
        Ok(Self { raw })
    }

    /// Construct an indexed palette color.
    pub fn palette(index: u8, tint: i16) -> Result<Self> {
        let mut raw = [0; 8];
        raw[0] = 0x02;
        raw[1] = index;
        raw[2..4].copy_from_slice(&tint.to_le_bytes());
        Self::from_raw(raw)
    }

    /// Construct an explicit RGBA color.
    pub fn rgb(red: u8, green: u8, blue: u8, alpha: u8, tint: i16) -> Self {
        let mut raw = [0; 8];
        raw[0] = 0x05;
        raw[2..4].copy_from_slice(&tint.to_le_bytes());
        raw[4..8].copy_from_slice(&[red, green, blue, alpha]);
        Self { raw }
    }

    /// Construct a theme color using the `clrScheme` index domain 0..=11.
    pub fn theme(index: u8, tint: i16) -> Result<Self> {
        let mut raw = [0; 8];
        raw[0] = 0x06;
        raw[1] = index;
        raw[2..4].copy_from_slice(&tint.to_le_bytes());
        Self::from_raw(raw)
    }

    /// Return the color family.
    #[must_use]
    pub const fn color_type(self) -> ColorType {
        match self.raw[0] >> 1 {
            1 => ColorType::Palette,
            2 => ColorType::Rgb,
            3 => ColorType::Theme,
            _ => unreachable!(),
        }
    }

    /// Return the stored index byte, including its ignored value for RGB.
    #[must_use]
    pub const fn index(self) -> u8 {
        self.raw[1]
    }

    /// Return the exact tint/shade integer.
    #[must_use]
    pub const fn tint(self) -> i16 {
        i16::from_le_bytes([self.raw[2], self.raw[3]])
    }

    /// Return the stored RGBA bytes, including ignored bytes for non-RGB
    /// colors.
    #[must_use]
    pub const fn rgba(self) -> [u8; 4] {
        [self.raw[4], self.raw[5], self.raw[6], self.raw[7]]
    }

    /// Return the exact validated wire representation.
    #[must_use]
    pub const fn raw(self) -> [u8; 8] {
        self.raw
    }
}

/// The eight colors in normative XLSB wire order.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Colors {
    series: Color,
    negative: Color,
    axis: Color,
    markers: Color,
    first: Color,
    last: Color,
    high: Color,
    low: Color,
}

impl Colors {
    /// Construct all eight ordered color slots.
    #[allow(clippy::too_many_arguments)]
    #[must_use]
    pub const fn new(
        series: Color,
        negative: Color,
        axis: Color,
        markers: Color,
        first: Color,
        last: Color,
        high: Color,
        low: Color,
    ) -> Self {
        Self {
            series,
            negative,
            axis,
            markers,
            first,
            last,
            high,
            low,
        }
    }

    /// Use one color for every slot.
    #[must_use]
    pub const fn uniform(color: Color) -> Self {
        Self::new(color, color, color, color, color, color, color, color)
    }

    /// Main series color.
    #[must_use]
    pub const fn series(self) -> Color {
        self.series
    }

    /// Negative-point color.
    #[must_use]
    pub const fn negative(self) -> Color {
        self.negative
    }

    /// Horizontal-axis color.
    #[must_use]
    pub const fn axis(self) -> Color {
        self.axis
    }

    /// Marker color.
    #[must_use]
    pub const fn markers(self) -> Color {
        self.markers
    }

    /// First-point color.
    #[must_use]
    pub const fn first(self) -> Color {
        self.first
    }

    /// Last-point color.
    #[must_use]
    pub const fn last(self) -> Color {
        self.last
    }

    /// High-point color.
    #[must_use]
    pub const fn high(self) -> Color {
        self.high
    }

    /// Low-point color.
    #[must_use]
    pub const fn low(self) -> Color {
        self.low
    }

    pub(crate) const fn ordered(self) -> [Color; 8] {
        [
            self.series,
            self.negative,
            self.axis,
            self.markers,
            self.first,
            self.last,
            self.high,
            self.low,
        ]
    }
}

bitflags! {
    /// Independent sparkline display switches.
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
    pub struct Options: u16 {
        /// Display a marker for each point.
        const MARKERS = 0x0001;
        /// Highlight the highest point.
        const HIGH = 0x0002;
        /// Highlight the lowest point.
        const LOW = 0x0004;
        /// Highlight the first point.
        const FIRST = 0x0008;
        /// Highlight the last point.
        const LAST = 0x0010;
        /// Highlight negative points.
        const NEGATIVE = 0x0020;
        /// Display the horizontal axis.
        const AXIS = 0x0040;
        /// Include hidden cells in the plot.
        const DISPLAY_HIDDEN = 0x0080;
        /// Display points right-to-left.
        const RIGHT_TO_LEFT = 0x0100;
    }
}

/// One validated vertical-axis bound.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Axis {
    kind: AxisType,
    manual: f64,
}

impl Axis {
    /// Construct an axis. Automatic individual/group axes require a stored
    /// manual value of positive zero, exactly as required on the XLSB wire.
    pub fn new(kind: AxisType, manual: f64) -> Result<Self> {
        validate_xnum(manual, "sparkline axis bound")?;
        if kind != AxisType::Custom && manual.to_bits() != 0 {
            return Err(value(
                "BrtBeginSparklineGroup",
                "automatic axis bounds require dManualMin/dManualMax to be positive zero",
            ));
        }
        Ok(Self { kind, manual })
    }

    /// An individually scaled automatic axis.
    #[must_use]
    pub const fn individual() -> Self {
        Self {
            kind: AxisType::Individual,
            manual: 0.0,
        }
    }

    /// A group-scaled automatic axis.
    #[must_use]
    pub const fn group() -> Self {
        Self {
            kind: AxisType::Group,
            manual: 0.0,
        }
    }

    /// A custom manual axis value.
    pub fn custom(value: f64) -> Result<Self> {
        Self::new(AxisType::Custom, value)
    }

    /// Axis scaling mode.
    #[must_use]
    pub const fn kind(self) -> AxisType {
        self.kind
    }

    /// Exact stored manual value.
    #[must_use]
    pub const fn manual(self) -> f64 {
        self.manual
    }
}

/// FRT range-adjustment state retained for exact rewrite.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub struct FrtState {
    adjusted_deleted: bool,
    adjusted_changed: bool,
    edited: bool,
    unused: bool,
}

impl FrtState {
    /// Construct all retained state bits. `fDoAdjust` is always emitted as one.
    #[must_use]
    pub const fn new(
        adjusted_deleted: bool,
        adjusted_changed: bool,
        edited: bool,
        unused: bool,
    ) -> Self {
        Self {
            adjusted_deleted,
            adjusted_changed,
            edited,
            unused,
        }
    }

    /// Whether a future-record consumer marked the range deleted.
    #[must_use]
    pub const fn adjusted_deleted(self) -> bool {
        self.adjusted_deleted
    }

    /// Whether a future-record consumer adjusted the range.
    #[must_use]
    pub const fn adjusted_changed(self) -> bool {
        self.adjusted_changed
    }

    /// Whether a cell in the range was edited.
    #[must_use]
    pub const fn edited(self) -> bool {
        self.edited
    }

    /// Undefined bit 16 retained without interpretation.
    #[must_use]
    pub const fn unused(self) -> bool {
        self.unused
    }

    pub(crate) const fn from_wire(flags: u32) -> Self {
        Self::new(
            flags & 0x01 != 0,
            flags & 0x04 != 0,
            flags & 0x08 != 0,
            flags & 0x0001_0000 != 0,
        )
    }

    pub(crate) const fn wire(self) -> u32 {
        0x02 | (self.adjusted_deleted as u32)
            | ((self.adjusted_changed as u32) << 2)
            | ((self.edited as u32) << 3)
            | ((self.unused as u32) << 16)
    }
}

/// A checked single-cell sparkline destination.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Location {
    row: u32,
    column: u32,
    state: FrtState,
}

impl Location {
    /// Construct a destination with pristine FRT state.
    pub fn new(row: u32, column: u32) -> Result<Self> {
        Self::with_state(row, column, FrtState::default())
    }

    /// Construct a destination with explicit retained FRT state.
    pub fn with_state(row: u32, column: u32, state: FrtState) -> Result<Self> {
        if row >= ROWS || column >= COLUMNS {
            return Err(value(
                "BrtSparkline",
                format!("destination ({row}, {column}) is outside the worksheet grid"),
            ));
        }
        Ok(Self { row, column, state })
    }

    /// Zero-based row.
    #[must_use]
    pub const fn row(self) -> u32 {
        self.row
    }

    /// Zero-based column.
    #[must_use]
    pub const fn column(self) -> u32 {
        self.column
    }

    /// Retained FRT range-adjustment state.
    #[must_use]
    pub const fn state(self) -> FrtState {
        self.state
    }
}

/// One sparkline and its destination cell.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Sparkline {
    formula: Option<Formula>,
    location: Location,
}

impl Sparkline {
    /// Construct one sparkline. `None` preserves the allowed formula-absent
    /// FRT form without inventing a source range.
    #[must_use]
    pub const fn new(location: Location, formula: Option<Formula>) -> Self {
        Self { formula, location }
    }

    /// Source formula, retained inertly.
    #[must_use]
    pub const fn formula(&self) -> Option<&Formula> {
        self.formula.as_ref()
    }

    /// Single destination cell.
    #[must_use]
    pub const fn location(&self) -> Location {
        self.location
    }
}

/// One XLSB sparkline group.
#[derive(Debug, Clone, PartialEq)]
pub struct Group {
    kind: SparklineType,
    empty_cells: EmptyCells,
    options: Options,
    colors: Colors,
    minimum: Axis,
    maximum: Axis,
    line_weight: f64,
    date_axis: bool,
    date_formula: Option<Formula>,
    items: Vec<Sparkline>,
}

impl Group {
    /// Construct a group using default limits and ordinary line-chart defaults.
    pub fn new(kind: SparklineType, colors: Colors, items: Vec<Sparkline>) -> Result<Self> {
        Self::with_limits(kind, colors, items, Limits::DEFAULT)
    }

    /// Construct a group using an explicit resource policy.
    pub fn with_limits(
        kind: SparklineType,
        colors: Colors,
        items: Vec<Sparkline>,
        limits: Limits,
    ) -> Result<Self> {
        let value = Self {
            kind,
            empty_cells: EmptyCells::Zero,
            options: Options::empty(),
            colors,
            minimum: Axis::individual(),
            maximum: Axis::individual(),
            line_weight: 1.0,
            date_axis: false,
            date_formula: None,
            items,
        };
        validate_group(&value, limits)?;
        Ok(value)
    }

    /// Set empty-cell plotting behavior.
    #[must_use]
    pub fn with_empty_cells(mut self, value: EmptyCells) -> Self {
        self.empty_cells = value;
        self
    }

    /// Set independent display switches.
    #[must_use]
    pub fn with_options(mut self, value: Options) -> Result<Self> {
        validate_options(value)?;
        self.options = value;
        Ok(self)
    }

    /// Set the minimum and maximum vertical-axis modes.
    #[must_use]
    pub fn with_axes(mut self, minimum: Axis, maximum: Axis) -> Self {
        self.minimum = minimum;
        self.maximum = maximum;
        self
    }

    /// Set a validated line weight in points.
    pub fn with_line_weight(mut self, value: f64) -> Result<Self> {
        validate_line_weight(value)?;
        self.line_weight = value;
        Ok(self)
    }

    /// Enable a date axis using one inert source formula.
    #[must_use]
    pub fn with_date_axis(mut self, formula: Formula) -> Self {
        self.date_axis = true;
        self.date_formula = Some(formula);
        self
    }

    /// Set the date-axis flag independently of formula presence. The XLSB
    /// record does not normatively require these two fields to agree.
    #[must_use]
    pub fn with_date_axis_enabled(mut self, enabled: bool) -> Self {
        self.date_axis = enabled;
        self
    }

    /// Set or clear the inert date-range formula independently of the flag.
    #[must_use]
    pub fn with_date_formula(mut self, formula: Option<Formula>) -> Self {
        self.date_formula = formula;
        self
    }

    /// Sparkline visual type.
    #[must_use]
    pub const fn kind(&self) -> SparklineType {
        self.kind
    }

    /// Empty-cell plotting behavior.
    #[must_use]
    pub const fn empty_cells(&self) -> EmptyCells {
        self.empty_cells
    }

    /// Independent display switches.
    #[must_use]
    pub const fn options(&self) -> Options {
        self.options
    }

    /// Eight ordered colors.
    #[must_use]
    pub const fn colors(&self) -> Colors {
        self.colors
    }

    /// Minimum vertical-axis mode and wire value.
    #[must_use]
    pub const fn minimum(&self) -> Axis {
        self.minimum
    }

    /// Maximum vertical-axis mode and wire value.
    #[must_use]
    pub const fn maximum(&self) -> Axis {
        self.maximum
    }

    /// Line weight in points.
    #[must_use]
    pub const fn line_weight(&self) -> f64 {
        self.line_weight
    }

    /// Whether the group uses its retained date range as an axis.
    #[must_use]
    pub const fn date_axis(&self) -> bool {
        self.date_axis
    }

    /// Optional inert date-range formula.
    #[must_use]
    pub const fn date_formula(&self) -> Option<&Formula> {
        self.date_formula.as_ref()
    }

    /// Ordered sparkline slice.
    #[must_use]
    pub fn sparklines(&self) -> &[Sparkline] {
        &self.items
    }

    pub(crate) fn from_wire(
        kind: SparklineType,
        empty_cells: EmptyCells,
        options: Options,
        colors: Colors,
        minimum: Axis,
        maximum: Axis,
        line_weight: f64,
        date_axis: bool,
        date_formula: Option<Formula>,
        items: Vec<Sparkline>,
        limits: Limits,
    ) -> Result<Self> {
        let value = Self {
            kind,
            empty_cells,
            options,
            colors,
            minimum,
            maximum,
            line_weight,
            date_axis,
            date_formula,
            items,
        };
        validate_group(&value, limits)?;
        Ok(value)
    }
}

/// A nonempty, ordered worksheet sparkline collection.
///
/// `[MS-XLSB]` does not require destination cells to be unique. Duplicate
/// locations are retained in stream order; consumers must define how they
/// present or edit multiple records targeting the same cell.
#[derive(Debug, Clone, PartialEq)]
pub struct Groups {
    groups: Vec<Group>,
}

impl Groups {
    /// Construct and validate using safe defaults.
    pub fn new(groups: Vec<Group>) -> Result<Self> {
        Self::with_limits(groups, Limits::DEFAULT)
    }

    /// Construct and validate with an explicit policy.
    pub fn with_limits(groups: Vec<Group>, limits: Limits) -> Result<Self> {
        let value = Self { groups };
        value.validate(limits)?;
        Ok(value)
    }

    /// Ordered group slice.
    #[must_use]
    pub fn as_slice(&self) -> &[Group] {
        &self.groups
    }

    /// Number of groups.
    #[must_use]
    pub fn len(&self) -> usize {
        self.groups.len()
    }

    /// This is always false for a valid instance.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.groups.is_empty()
    }

    /// Iterate in worksheet stream order.
    pub fn iter(&self) -> impl ExactSizeIterator<Item = &Group> {
        self.groups.iter()
    }

    pub(crate) fn validate(&self, limits: Limits) -> Result<()> {
        limits.validate()?;
        if self.groups.is_empty() {
            return Err(Error::Empty {
                collection: "sparkline groups",
            });
        }
        check_limit("groups", self.groups.len(), limits.groups)?;

        let expected_total = self.groups.iter().try_fold(0usize, |total, group| {
            total.checked_add(group.items.len()).ok_or(Error::Limit {
                resource: "total sparklines",
                actual: usize::MAX,
                maximum: limits.total,
            })
        })?;
        check_limit("total sparklines", expected_total, limits.total)?;

        let mut total = 0usize;
        for group in &self.groups {
            validate_group(group, limits)?;
            total = total.checked_add(group.items.len()).ok_or(Error::Limit {
                resource: "total sparklines",
                actual: usize::MAX,
                maximum: limits.total,
            })?;
            check_limit("total sparklines", total, limits.total)?;
            for item in &group.items {
                if let Some(formula) = &item.formula {
                    validate_formula(&formula.rgce, &formula.rgcb, limits)?;
                }
            }
        }
        Ok(())
    }
}

impl<'a> IntoIterator for &'a Groups {
    type Item = &'a Group;
    type IntoIter = std::slice::Iter<'a, Group>;

    fn into_iter(self) -> Self::IntoIter {
        self.groups.iter()
    }
}

pub(crate) fn validate_formula(rgce: &[u8], rgcb: &[u8], limits: Limits) -> Result<FormulaKind> {
    check_limit("formula token bytes", rgce.len(), limits.formula_tokens)?;
    check_limit(
        "formula ancillary bytes",
        rgcb.len(),
        limits.formula_ancillary,
    )?;
    let Some(&token) = rgce.first() else {
        return Err(value("FRTFormula", "cce must be in 1..=16384"));
    };
    if token & 0x80 != 0 {
        return Err(value("FRTFormula", "Ptg reserved bit must be zero"));
    }
    let class = (token >> 5) & 0x03;
    if !(1..=3).contains(&class) {
        return Err(value("FRTFormula", "PtgDataType must be 1, 2, or 3"));
    }
    match token & 0x1f {
        0x03 => {
            require_token_len(rgce, 5, "PtgName")?;
            if u32::from_le_bytes(rgce[1..5].try_into().expect("checked token")) == 0 {
                return Err(value("FRTFormula", "PtgName index must be one-based"));
            }
            Ok(FormulaKind::Name)
        },
        0x19 => {
            require_token_len(rgce, 7, "PtgNameX")?;
            if u32::from_le_bytes(rgce[3..7].try_into().expect("checked token")) == 0 {
                return Err(value("FRTFormula", "PtgNameX name index must be one-based"));
            }
            Ok(FormulaKind::ExternalName)
        },
        0x1a => {
            require_token_len(rgce, 9, "PtgRef3d")?;
            let row = u32::from_le_bytes(rgce[3..7].try_into().expect("checked token"));
            let column = u16::from_le_bytes(rgce[7..9].try_into().expect("checked token"));
            if row >= ROWS || u32::from(column & 0x3fff) >= COLUMNS {
                return Err(value(
                    "FRTFormula",
                    "PtgRef3d coordinate is outside the worksheet grid",
                ));
            }
            Ok(FormulaKind::Reference3d)
        },
        0x1b => {
            require_token_len(rgce, 15, "PtgArea3d")?;
            let row_first = u32::from_le_bytes(rgce[3..7].try_into().expect("checked token"));
            let row_last = u32::from_le_bytes(rgce[7..11].try_into().expect("checked token"));
            let first = u16::from_le_bytes(rgce[11..13].try_into().expect("checked token"));
            let last = u16::from_le_bytes(rgce[13..15].try_into().expect("checked token"));
            let col_first = u32::from(first & 0x3fff);
            let col_last = u32::from(last & 0x3fff);
            if row_first > row_last
                || row_last >= ROWS
                || col_first > col_last
                || col_last >= COLUMNS
            {
                return Err(value(
                    "FRTFormula",
                    "PtgArea3d range is outside the worksheet grid",
                ));
            }
            let row_vector = row_first == row_last && (first & 0x4000) == (last & 0x4000);
            let column_vector = col_first == col_last && (first & 0x8000) == (last & 0x8000);
            if !row_vector && !column_vector {
                return Err(value(
                    "FRTFormula",
                    "PtgArea3d sparkline source must be one-dimensional",
                ));
            }
            Ok(FormulaKind::Area3d)
        },
        _ => Err(value(
            "FRTFormula",
            "single token must be PtgName, PtgNameX, PtgRef3d, or PtgArea3d",
        )),
    }
}

fn validate_group(group: &Group, limits: Limits) -> Result<()> {
    limits.validate()?;
    if group.items.is_empty() {
        return Err(Error::Empty {
            collection: "sparklines in a group",
        });
    }
    check_limit("sparklines per group", group.items.len(), limits.per_group)?;
    validate_line_weight(group.line_weight)?;
    validate_options(group.options)?;
    Axis::new(group.minimum.kind, group.minimum.manual)?;
    Axis::new(group.maximum.kind, group.maximum.manual)?;
    if let Some(formula) = &group.date_formula {
        validate_formula(&formula.rgce, &formula.rgcb, limits)?;
    }
    for color in group.colors.ordered() {
        validate_color(color.raw)?;
    }
    Ok(())
}

fn validate_options(options: Options) -> Result<()> {
    let unknown = options.bits() & !Options::all().bits();
    if unknown != 0 {
        return Err(value(
            "BrtBeginSparklineGroup",
            format!("unknown semantic option bits 0x{unknown:04X}"),
        ));
    }
    Ok(())
}

fn validate_line_weight(weight: f64) -> Result<()> {
    validate_xnum(weight, "sparkline line weight")?;
    if !(0.0..=1584.0).contains(&weight) {
        return Err(value(
            "BrtBeginSparklineGroup",
            format!("dLineWeight {weight} is outside 0..=1584"),
        ));
    }
    Ok(())
}

fn validate_xnum(number: f64, field: &'static str) -> Result<()> {
    if matches!(
        number.classify(),
        FpCategory::Nan | FpCategory::Infinite | FpCategory::Subnormal
    ) || (number == 0.0 && number.is_sign_negative())
    {
        return Err(value(
            "Xnum",
            format!(
                "{field} has forbidden IEEE-754 bits {:#018x}",
                number.to_bits()
            ),
        ));
    }
    Ok(())
}

fn validate_color(raw: [u8; 8]) -> Result<()> {
    let valid_rgb = raw[0] & 1 != 0;
    match raw[0] >> 1 {
        1 if raw[1] <= 0x51 => Ok(()),
        1 => Err(value(
            "BrtColor",
            format!("palette index {} is outside the Icv domain", raw[1]),
        )),
        2 if valid_rgb => Ok(()),
        2 => Err(value(
            "BrtColor",
            "fValidRGB must be one for an explicit RGB color",
        )),
        3 if raw[1] <= 0x0b => Ok(()),
        3 => Err(value(
            "BrtColor",
            format!("theme index {} is outside 0..=11", raw[1]),
        )),
        0 => Err(value(
            "BrtColor",
            "automatic colors are forbidden in sparkline color slots",
        )),
        kind => Err(value(
            "BrtColor",
            format!("xColorType {kind} is outside 0..=3"),
        )),
    }
}

fn require_token_len(rgce: &[u8], expected: usize, kind: &'static str) -> Result<()> {
    if rgce.len() != expected {
        return Err(value(
            "FRTFormula",
            format!(
                "{kind} must occupy exactly {expected} rgce bytes, found {}",
                rgce.len()
            ),
        ));
    }
    Ok(())
}

pub(crate) fn check_limit(resource: &'static str, actual: usize, maximum: usize) -> Result<()> {
    if actual > maximum {
        Err(Error::Limit {
            resource,
            actual,
            maximum,
        })
    } else {
        Ok(())
    }
}

pub(crate) fn allocation(resource: &'static str) -> Error {
    Error::Allocation { resource }
}

pub(crate) fn value(record: &'static str, reason: impl Into<String>) -> Error {
    Error::Value {
        record,
        reason: reason.into(),
    }
}

fn copy_boxed(bytes: &[u8], resource: &'static str) -> Result<Box<[u8]>> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(bytes.len())
        .map_err(|_| allocation(resource))?;
    output.extend_from_slice(bytes);
    Ok(output.into_boxed_slice())
}
