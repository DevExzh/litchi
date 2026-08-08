#![cfg_attr(not(test), deny(clippy::indexing_slicing))]

//! RTF table support.
//!
//! This module provides basic table parsing for RTF documents.
//! RTF tables use a complex row-based model with cell boundaries.

use crate::TextDirection;
use std::borrow::Cow;
use std::collections::{BTreeSet, HashSet};

pub const MAX_TABLE_DISTANCE_TWIPS: i32 = 31_680;
pub const MAX_TABLE_GEOMETRY_TWIPS: i32 = 31_680;
pub const MAX_TABLE_WIDTH_PERCENT: i32 = 5_000;
pub const MAX_FLOATING_TABLE_DISTANCE_TWIPS: i32 = 31_680;
pub const MAX_TABLE_NESTING_DEPTH: usize = 32;
pub const MAX_TABLE_CELLS_PER_ROW: usize = 4_096;
pub const MAX_TABLE_ROW_INDEX: u16 = u16::MAX;

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum TablePreferredWidthUnit {
    #[default]
    Null,
    Auto,
    Percent,
    Twips,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TablePreferredWidth {
    unit: TablePreferredWidthUnit,
    value: Option<u16>,
}
impl TablePreferredWidth {
    pub fn new(unit: TablePreferredWidthUnit, value: Option<u16>) -> crate::RtfResult<Self> {
        let width = Self { unit, value };
        width.validate()?;
        Ok(width)
    }
    pub const fn unit(self) -> TablePreferredWidthUnit {
        self.unit
    }
    pub const fn value(self) -> Option<u16> {
        self.value
    }
    pub(crate) fn validate(self) -> crate::RtfResult<()> {
        match (self.unit, self.value) {
            (TablePreferredWidthUnit::Null | TablePreferredWidthUnit::Auto, None) => Ok(()),
            (TablePreferredWidthUnit::Percent, Some(value))
                if i32::from(value) <= MAX_TABLE_WIDTH_PERCENT =>
            {
                Ok(())
            },
            (TablePreferredWidthUnit::Twips, Some(value))
                if i32::from(value) <= MAX_TABLE_GEOMETRY_TWIPS =>
            {
                Ok(())
            },
            (TablePreferredWidthUnit::Null | TablePreferredWidthUnit::Auto, Some(_)) => {
                Err(crate::RtfError::MalformedDocument(
                    "RTF null or auto preferred width cannot carry a value".to_string(),
                ))
            },
            (TablePreferredWidthUnit::Percent | TablePreferredWidthUnit::Twips, None) => {
                Err(crate::RtfError::MalformedDocument(
                    "RTF percentage or twip preferred width requires a value".to_string(),
                ))
            },
            _ => Err(crate::RtfError::MalformedDocument(
                "RTF preferred table width is out of range".to_string(),
            )),
        }
    }
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum TableRowHeight {
    #[default]
    Automatic,
    Minimum(u16),
    Exact(u16),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableIndentUnit {
    Auto,
    Twips,
    Nil,
    Percent,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TableIndent {
    unit: TableIndentUnit,
    value: i32,
}
impl TableIndent {
    pub fn new(unit: TableIndentUnit, value: i32) -> crate::RtfResult<Self> {
        let indent = Self { unit, value };
        indent.validate()?;
        Ok(indent)
    }
    pub const fn unit(self) -> TableIndentUnit {
        self.unit
    }
    pub const fn value(self) -> i32 {
        self.value
    }
    pub(crate) fn validate(self) -> crate::RtfResult<()> {
        let cap = match self.unit {
            TableIndentUnit::Twips => MAX_TABLE_GEOMETRY_TWIPS,
            TableIndentUnit::Percent => MAX_TABLE_WIDTH_PERCENT,
            TableIndentUnit::Auto | TableIndentUnit::Nil => {
                return if self.value == 0 {
                    Ok(())
                } else {
                    Err(crate::RtfError::MalformedDocument(
                        "RTF auto or nil table indent requires zero".to_string(),
                    ))
                };
            },
        };
        if self.value.unsigned_abs() <= cap as u32 {
            Ok(())
        } else {
            Err(crate::RtfError::MalformedDocument(
                "RTF table indent is out of range".to_string(),
            ))
        }
    }
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct TableRowGeometry {
    half_gap_twips: Option<u16>,
    left_edge_twips: Option<i32>,
    height: TableRowHeight,
    preferred_width: Option<TablePreferredWidth>,
    leading_invisible_width: Option<TablePreferredWidth>,
    trailing_invisible_width: Option<TablePreferredWidth>,
    auto_fit: bool,
    indent: Option<TableIndent>,
}
impl TableRowGeometry {
    pub const fn half_gap_twips(self) -> Option<u16> {
        self.half_gap_twips
    }
    pub fn set_half_gap_twips(&mut self, value: Option<u16>) {
        self.half_gap_twips = value
    }
    pub const fn left_edge_twips(self) -> Option<i32> {
        self.left_edge_twips
    }
    pub fn set_left_edge_twips(&mut self, value: Option<i32>) {
        self.left_edge_twips = value
    }
    pub const fn height(self) -> TableRowHeight {
        self.height
    }
    pub fn set_height(&mut self, value: TableRowHeight) {
        self.height = value
    }
    pub const fn preferred_width(self) -> Option<TablePreferredWidth> {
        self.preferred_width
    }
    pub fn set_preferred_width(&mut self, value: Option<TablePreferredWidth>) {
        self.preferred_width = value
    }
    /// Width of the invisible cell at the logical beginning of the row (`trftsWidthB`/`trwWidthB`).
    /// This is the left side for an LTR row and the right side for an RTL row.
    pub const fn leading_invisible_width(self) -> Option<TablePreferredWidth> {
        self.leading_invisible_width
    }
    pub fn set_leading_invisible_width(&mut self, value: Option<TablePreferredWidth>) {
        self.leading_invisible_width = value
    }
    /// Width of the invisible cell at the logical end of the row (`trftsWidthA`/`trwWidthA`).
    /// This is the right side for an LTR row and the left side for an RTL row.
    pub const fn trailing_invisible_width(self) -> Option<TablePreferredWidth> {
        self.trailing_invisible_width
    }
    pub fn set_trailing_invisible_width(&mut self, value: Option<TablePreferredWidth>) {
        self.trailing_invisible_width = value
    }
    pub const fn auto_fit(self) -> bool {
        self.auto_fit
    }
    pub fn set_auto_fit(&mut self, value: bool) {
        self.auto_fit = value
    }
    pub const fn indent(self) -> Option<TableIndent> {
        self.indent
    }
    pub fn set_indent(&mut self, value: Option<TableIndent>) {
        self.indent = value
    }
    pub(crate) fn validate(self) -> crate::RtfResult<()> {
        if self
            .half_gap_twips
            .is_some_and(|value| i32::from(value) > MAX_TABLE_GEOMETRY_TWIPS)
            || self
                .left_edge_twips
                .is_some_and(|value| value.unsigned_abs() > MAX_TABLE_GEOMETRY_TWIPS as u32)
            || matches!(self.height,TableRowHeight::Minimum(value)|TableRowHeight::Exact(value)if i32::from(value)>MAX_TABLE_GEOMETRY_TWIPS)
        {
            return Err(crate::RtfError::MalformedDocument(
                "RTF table row geometry is out of range".to_string(),
            ));
        }
        for width in [
            self.preferred_width,
            self.leading_invisible_width,
            self.trailing_invisible_width,
        ]
        .into_iter()
        .flatten()
        {
            width.validate()?
        }
        if let Some(indent) = self.indent {
            indent.validate()?
        }
        Ok(())
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableCellMergeAxis {
    Horizontal,
    Vertical,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableCellMergeRole {
    First,
    Continuation,
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct TableCellMergeState {
    pub horizontal: Option<TableCellMergeRole>,
    pub vertical: Option<TableCellMergeRole>,
}

/// Kind of tracked-change revision attached to one table cell.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CellRevisionKind {
    /// The cell was inserted as a revision (`\clins`).
    Inserted,
    /// The cell was deleted as a revision (`\cldel`).
    Deleted,
    /// The cell was removed by a merge revision (`\clmrgd`).
    MergeDeleted,
}

impl CellRevisionKind {
    /// The RTF control word that marks this revision kind.
    pub const fn control_word(self) -> &'static str {
        match self {
            Self::Inserted => "clins",
            Self::Deleted => "cldel",
            Self::MergeDeleted => "clmrgd",
        }
    }

    /// The RTF control word carrying this revision's author index.
    pub const fn author_control_word(self) -> &'static str {
        match self {
            Self::Inserted => "clinsauth",
            Self::Deleted => "cldelauth",
            Self::MergeDeleted => "clmrgdauth",
        }
    }

    /// The RTF control word carrying this revision's packed DTTM timestamp.
    pub const fn date_control_word(self) -> &'static str {
        match self {
            Self::Inserted => "clinsdttm",
            Self::Deleted => "cldeldttm",
            Self::MergeDeleted => "clmrgddttm",
        }
    }
}

/// Tracked-change revision attached to one table cell.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CellRevision {
    pub kind: CellRevisionKind,
    /// Author/date metadata from the matching `\cl*authN`/`\cl*dttmN` pair.
    pub metadata: crate::RevisionMetadata,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableHorizontalReference {
    Column,
    Margin,
    Page,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableVerticalReference {
    Margin,
    Paragraph,
    Page,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableHorizontalPosition {
    Offset(i32),
    NegativeOffset(i32),
    Center,
    Inside,
    Left,
    Outside,
    Right,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableVerticalPosition {
    Offset(i32),
    NegativeOffset(i32),
    Bottom,
    Center,
    Inline,
    Inside,
    Outside,
    Top,
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct TableWrapDistances {
    pub left: Option<u16>,
    pub right: Option<u16>,
    pub top: Option<u16>,
    pub bottom: Option<u16>,
}
impl TableWrapDistances {
    pub(crate) fn side_mut(&mut self, edge: TableEdge) -> &mut Option<u16> {
        match edge {
            TableEdge::Left => &mut self.left,
            TableEdge::Right => &mut self.right,
            TableEdge::Top => &mut self.top,
            TableEdge::Bottom => &mut self.bottom,
        }
    }
    pub fn validate(&self) -> crate::RtfResult<()> {
        if [self.left, self.right, self.top, self.bottom]
            .into_iter()
            .flatten()
            .any(|value| i32::from(value) > MAX_FLOATING_TABLE_DISTANCE_TWIPS)
        {
            return Err(crate::RtfError::MalformedDocument(
                "RTF floating-table wrap distance is out of range".to_string(),
            ));
        }
        Ok(())
    }
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct FloatingTablePosition {
    pub horizontal_reference: Option<TableHorizontalReference>,
    pub horizontal_position: Option<TableHorizontalPosition>,
    pub vertical_reference: Option<TableVerticalReference>,
    pub vertical_position: Option<TableVerticalPosition>,
    pub no_overlap: bool,
    pub wrap_distances: TableWrapDistances,
}
impl FloatingTablePosition {
    pub fn is_empty(&self) -> bool {
        self.horizontal_reference.is_none()
            && self.horizontal_position.is_none()
            && self.vertical_reference.is_none()
            && self.vertical_position.is_none()
            && !self.no_overlap
            && self.wrap_distances == TableWrapDistances::default()
    }
    pub fn validate(&self) -> crate::RtfResult<()> {
        let offset = |value| (0..=MAX_FLOATING_TABLE_DISTANCE_TWIPS).contains(&value);
        let negative = |value| (-MAX_FLOATING_TABLE_DISTANCE_TWIPS..=-1).contains(&value);
        if matches!(self.horizontal_position,Some(TableHorizontalPosition::Offset(value))if !offset(value))
            || matches!(self.vertical_position,Some(TableVerticalPosition::Offset(value))if !offset(value))
            || matches!(self.horizontal_position,Some(TableHorizontalPosition::NegativeOffset(value))if !negative(value))
            || matches!(self.vertical_position,Some(TableVerticalPosition::NegativeOffset(value))if !negative(value))
        {
            return Err(crate::RtfError::MalformedDocument(
                "RTF floating-table position is out of range".to_string(),
            ));
        }
        self.wrap_distances.validate()
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableDistanceUnit {
    Null,
    Twips,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableDistanceScope {
    Row,
    Cell,
}
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableDistanceKind {
    Padding,
    Spacing,
}
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableEdge {
    Left,
    Right,
    Top,
    Bottom,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableRowBorderSide {
    Top,
    Left,
    Bottom,
    Right,
    Horizontal,
    Vertical,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableCellBorderSide {
    Top,
    Left,
    Bottom,
    Right,
    UpperLeftToLowerRight,
    UpperRightToLowerLeft,
}

/// Sides of the table-style default borders (`\tsbrdr*`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableStyleBorderSide {
    Top,
    Left,
    Bottom,
    Right,
    HorizontalInside,
    VerticalInside,
    DiagonalUpperLeftToLowerRight,
    DiagonalUpperRightToLowerLeft,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum TableBorderTarget {
    Row(TableRowBorderSide),
    Cell(TableCellBorderSide),
    StyleDefault(TableStyleBorderSide),
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct TableRowBorders {
    pub top: Option<crate::Border>,
    pub left: Option<crate::Border>,
    pub bottom: Option<crate::Border>,
    pub right: Option<crate::Border>,
    pub horizontal: Option<crate::Border>,
    pub vertical: Option<crate::Border>,
}
impl TableRowBorders {
    pub(crate) fn side_mut(&mut self, side: TableRowBorderSide) -> &mut Option<crate::Border> {
        match side {
            TableRowBorderSide::Top => &mut self.top,
            TableRowBorderSide::Left => &mut self.left,
            TableRowBorderSide::Bottom => &mut self.bottom,
            TableRowBorderSide::Right => &mut self.right,
            TableRowBorderSide::Horizontal => &mut self.horizontal,
            TableRowBorderSide::Vertical => &mut self.vertical,
        }
    }
    pub fn validate(&self) -> crate::RtfResult<()> {
        for border in [
            self.top,
            self.left,
            self.bottom,
            self.right,
            self.horizontal,
            self.vertical,
        ]
        .into_iter()
        .flatten()
        {
            border.validate_table()?;
        }
        Ok(())
    }
}

/// Table-style default borders declared once per row (`\tsbrdr*`).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct TableStyleDefaultBorders {
    pub top: Option<crate::Border>,
    pub left: Option<crate::Border>,
    pub bottom: Option<crate::Border>,
    pub right: Option<crate::Border>,
    pub horizontal_inside: Option<crate::Border>,
    pub vertical_inside: Option<crate::Border>,
    pub diagonal_upper_left_to_lower_right: Option<crate::Border>,
    pub diagonal_upper_right_to_lower_left: Option<crate::Border>,
}
impl TableStyleDefaultBorders {
    pub(crate) fn side_mut(&mut self, side: TableStyleBorderSide) -> &mut Option<crate::Border> {
        match side {
            TableStyleBorderSide::Top => &mut self.top,
            TableStyleBorderSide::Left => &mut self.left,
            TableStyleBorderSide::Bottom => &mut self.bottom,
            TableStyleBorderSide::Right => &mut self.right,
            TableStyleBorderSide::HorizontalInside => &mut self.horizontal_inside,
            TableStyleBorderSide::VerticalInside => &mut self.vertical_inside,
            TableStyleBorderSide::DiagonalUpperLeftToLowerRight => {
                &mut self.diagonal_upper_left_to_lower_right
            },
            TableStyleBorderSide::DiagonalUpperRightToLowerLeft => {
                &mut self.diagonal_upper_right_to_lower_left
            },
        }
    }
    /// Whether no default border was explicitly retained.
    pub fn is_empty(&self) -> bool {
        [
            self.top,
            self.left,
            self.bottom,
            self.right,
            self.horizontal_inside,
            self.vertical_inside,
            self.diagonal_upper_left_to_lower_right,
            self.diagonal_upper_right_to_lower_left,
        ]
        .into_iter()
        .all(|border| border.is_none())
    }
    pub fn validate(&self) -> crate::RtfResult<()> {
        for border in [
            self.top,
            self.left,
            self.bottom,
            self.right,
            self.horizontal_inside,
            self.vertical_inside,
            self.diagonal_upper_left_to_lower_right,
            self.diagonal_upper_right_to_lower_left,
        ]
        .into_iter()
        .flatten()
        {
            border.validate_table()?;
        }
        Ok(())
    }
}

/// Row-scoped default formatting applied to every cell in the row
/// (`\tsbrdr*`, `\tscellpadd*`, `\tscellspc*`, and `\tscellwidth*`).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct TableRowCellDefaults {
    pub borders: TableStyleDefaultBorders,
    pub padding: TableEdgeDistances,
    pub spacing: TableEdgeDistances,
    pub preferred_cell_width: Option<TablePreferredWidth>,
}
impl TableRowCellDefaults {
    /// Whether no default formatting was explicitly retained.
    pub fn is_empty(&self) -> bool {
        self.borders.is_empty()
            && self.padding == TableEdgeDistances::default()
            && self.spacing == TableEdgeDistances::default()
            && self.preferred_cell_width.is_none()
    }
    pub fn validate(&self) -> crate::RtfResult<()> {
        self.borders.validate()?;
        self.padding.validate()?;
        self.spacing.validate()?;
        if let Some(width) = self.preferred_cell_width {
            width.validate()?;
        }
        Ok(())
    }
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct TableCellBorders {
    pub top: Option<crate::Border>,
    pub left: Option<crate::Border>,
    pub bottom: Option<crate::Border>,
    pub right: Option<crate::Border>,
    pub upper_left_to_lower_right: Option<crate::Border>,
    pub upper_right_to_lower_left: Option<crate::Border>,
}
impl TableCellBorders {
    pub(crate) fn side_mut(&mut self, side: TableCellBorderSide) -> &mut Option<crate::Border> {
        match side {
            TableCellBorderSide::Top => &mut self.top,
            TableCellBorderSide::Left => &mut self.left,
            TableCellBorderSide::Bottom => &mut self.bottom,
            TableCellBorderSide::Right => &mut self.right,
            TableCellBorderSide::UpperLeftToLowerRight => &mut self.upper_left_to_lower_right,
            TableCellBorderSide::UpperRightToLowerLeft => &mut self.upper_right_to_lower_left,
        }
    }
    pub fn validate(&self) -> crate::RtfResult<()> {
        for border in [
            self.top,
            self.left,
            self.bottom,
            self.right,
            self.upper_left_to_lower_right,
            self.upper_right_to_lower_left,
        ]
        .into_iter()
        .flatten()
        {
            border.validate_table()?;
        }
        Ok(())
    }
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct TableShading {
    pub amount: Option<u16>,
    /// Raw cell shading amount (`\\clshdngraw`).
    pub raw_amount: Option<u16>,
    /// Whether cell shading is explicitly raw-nil (`\\clshdrawnil`).
    pub raw_nil: bool,
    pub foreground_color: Option<crate::ColorRef>,
    pub background_color: Option<crate::ColorRef>,
    pub pattern: Option<crate::ShadingPattern>,
    pub pattern_index: Option<u16>,
}
impl TableShading {
    pub fn validate(&self) -> crate::RtfResult<()> {
        if self
            .amount
            .into_iter()
            .chain(self.raw_amount)
            .any(|value| value > 10_000)
        {
            return Err(crate::RtfError::MalformedDocument(
                "RTF table shading must be in 0..=10000".to_string(),
            ));
        }
        if self.pattern.is_some() && self.pattern_index.is_some() {
            return Err(crate::RtfError::MalformedDocument(
                "RTF table shading has conflicting pattern controls".to_string(),
            ));
        }
        Ok(())
    }
}
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TableDistanceTarget {
    pub scope: TableDistanceScope,
    pub kind: TableDistanceKind,
    pub edge: TableEdge,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableRowAlignment {
    Left,
    Center,
    Right,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableAutoformatFlag {
    Border,
    Shading,
    Font,
    Color,
    BestFit,
    HeaderRows,
    LastRow,
    HeaderColumns,
    LastColumn,
    NoRowBanding,
    NoColumnBanding,
}

impl TableAutoformatFlag {
    const fn bit(self) -> u16 {
        match self {
            Self::Border => 1 << 0,
            Self::Shading => 1 << 1,
            Self::Font => 1 << 2,
            Self::Color => 1 << 3,
            Self::BestFit => 1 << 4,
            Self::HeaderRows => 1 << 5,
            Self::LastRow => 1 << 6,
            Self::HeaderColumns => 1 << 7,
            Self::LastColumn => 1 << 8,
            Self::NoRowBanding => 1 << 9,
            Self::NoColumnBanding => 1 << 10,
        }
    }
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct TableAutoformatFlags {
    bits: u16,
}
impl TableAutoformatFlags {
    pub const fn contains(self, flag: TableAutoformatFlag) -> bool {
        self.bits & flag.bit() != 0
    }
    pub fn set(&mut self, flag: TableAutoformatFlag, enabled: bool) {
        if enabled {
            self.bits |= flag.bit()
        } else {
            self.bits &= !flag.bit()
        }
    }
    pub(crate) fn insert(&mut self, flag: TableAutoformatFlag) -> bool {
        let was_present = self.contains(flag);
        self.set(flag, true);
        !was_present
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableRowBandIndex {
    Header,
    Row(u16),
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct TableRowBanding {
    pub row_index: Option<u16>,
    pub band_index: Option<TableRowBandIndex>,
    pub last_row: bool,
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct TableRowLayout {
    pub header: bool,
    pub keep_together: bool,
    pub keep_with_following: bool,
    pub alignment: Option<TableRowAlignment>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableCellVerticalAlignment {
    Top,
    Center,
    Bottom,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableCellTextFlow {
    LeftToRightTopToBottom,
    RightToLeftTopToBottom,
    LeftToRightBottomToTop,
    LeftToRightTopToBottomVertical,
    TopToBottomRightToLeftVertical,
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct TableCellLayout {
    pub vertical_alignment: Option<TableCellVerticalAlignment>,
    pub text_flow: Option<TableCellTextFlow>,
    pub fit_text: bool,
    pub no_wrap: bool,
    pub hide_mark: bool,
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct TableSideDistance {
    pub value: Option<u16>,
    pub unit: Option<TableDistanceUnit>,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct TableEdgeDistances {
    pub left: TableSideDistance,
    pub right: TableSideDistance,
    pub top: TableSideDistance,
    pub bottom: TableSideDistance,
}
impl TableEdgeDistances {
    pub(crate) fn side_mut(&mut self, edge: TableEdge) -> &mut TableSideDistance {
        match edge {
            TableEdge::Left => &mut self.left,
            TableEdge::Right => &mut self.right,
            TableEdge::Top => &mut self.top,
            TableEdge::Bottom => &mut self.bottom,
        }
    }
    pub fn validate(&self) -> crate::RtfResult<()> {
        for side in [&self.left, &self.right, &self.top, &self.bottom] {
            if side
                .value
                .is_some_and(|value| i32::from(value) > MAX_TABLE_DISTANCE_TWIPS)
            {
                return Err(crate::RtfError::MalformedDocument(
                    "RTF table distance is out of range".to_string(),
                ));
            }
        }
        Ok(())
    }
}

/// A table in an RTF document.
#[derive(Debug, Clone)]
pub struct Table<'a> {
    /// Table rows
    rows: Vec<Row<'a>>,
    /// Explicit table direction from `\taprtl` or `\taprtl0`.
    direction: Option<TextDirection>,
}

impl<'a> Table<'a> {
    /// Create a new table.
    pub fn new() -> Self {
        Self {
            rows: Vec::new(),
            direction: None,
        }
    }

    /// Add a row to the table.
    pub fn add_row(&mut self, row: Row<'a>) {
        self.rows.push(row);
    }

    /// Get the number of rows.
    pub fn row_count(&self) -> usize {
        self.rows.len()
    }

    /// Get all rows.
    pub fn rows(&self) -> &[Row<'a>] {
        &self.rows
    }

    /// Return the explicit table direction.
    pub fn direction(&self) -> Option<TextDirection> {
        self.direction
    }

    /// Set or clear the explicit table direction.
    pub fn set_direction(&mut self, direction: Option<TextDirection>) {
        self.direction = direction;
    }
}

impl<'a> Default for Table<'a> {
    fn default() -> Self {
        Self::new()
    }
}

/// A table row.
#[derive(Debug, Clone)]
pub struct Row<'a> {
    /// Row cells
    cells: Vec<Cell<'a>>,
    /// Optional table-style handle referenced by this row.
    table_style: Option<u16>,
    /// RSID attached to the row (`\tblrsidN`).
    table_rsid: Option<u32>,
    /// Explicit row direction.
    direction: Option<TextDirection>,
    layout: TableRowLayout,
    padding: TableEdgeDistances,
    spacing: TableEdgeDistances,
    cell_defaults: TableRowCellDefaults,
    positioning: FloatingTablePosition,
    borders: TableRowBorders,
    shading: TableShading,
    geometry: TableRowGeometry,
    autoformat_flags: TableAutoformatFlags,
    banding: TableRowBanding,
    /// Author/date metadata for the revision that changed this row
    /// (`\trauthN`, `\trdateN`).
    revision: crate::RevisionMetadata,
}

impl<'a> Row<'a> {
    /// Create a new row.
    pub fn new() -> Self {
        Self {
            cells: Vec::new(),
            table_style: None,
            table_rsid: None,
            direction: None,
            layout: TableRowLayout::default(),
            padding: TableEdgeDistances::default(),
            spacing: TableEdgeDistances::default(),
            cell_defaults: TableRowCellDefaults::default(),
            positioning: FloatingTablePosition::default(),
            borders: Default::default(),
            shading: Default::default(),
            geometry: Default::default(),
            autoformat_flags: Default::default(),
            banding: Default::default(),
            revision: crate::RevisionMetadata::default(),
        }
    }

    /// Add a cell to the row.
    pub fn add_cell(&mut self, cell: Cell<'a>) {
        self.cells.push(cell);
    }

    /// Get the number of cells.
    pub fn cell_count(&self) -> usize {
        self.cells.len()
    }

    /// Get all cells.
    pub fn cells(&self) -> &[Cell<'a>] {
        &self.cells
    }

    /// Return the table-style handle referenced by this row.
    pub fn table_style(&self) -> Option<u16> {
        self.table_style
    }

    /// Set or clear the table-style handle referenced by this row.
    pub fn set_table_style(&mut self, table_style: Option<u16>) {
        self.table_style = table_style;
    }

    /// RSID attached to this row, when present.
    pub fn table_rsid(&self) -> Option<u32> {
        self.table_rsid
    }

    /// Set or clear the row RSID.
    pub fn set_table_rsid(&mut self, table_rsid: Option<u32>) {
        self.table_rsid = table_rsid;
    }

    /// Return the explicit row direction.
    pub fn direction(&self) -> Option<TextDirection> {
        self.direction
    }

    /// Set or clear the explicit row direction.
    pub fn set_direction(&mut self, direction: Option<TextDirection>) {
        self.direction = direction;
    }
    pub fn layout(&self) -> &TableRowLayout {
        &self.layout
    }
    pub fn set_layout(&mut self, value: TableRowLayout) {
        self.layout = value
    }
    pub fn padding(&self) -> &TableEdgeDistances {
        &self.padding
    }
    pub fn spacing(&self) -> &TableEdgeDistances {
        &self.spacing
    }
    pub fn set_padding(&mut self, value: TableEdgeDistances) {
        self.padding = value
    }
    pub fn set_spacing(&mut self, value: TableEdgeDistances) {
        self.spacing = value
    }
    /// Row-scoped default cell formatting (`\tsbrdr*`, `\tscellpadd*`,
    /// `\tscellspc*`, and `\tscellwidth*`).
    pub fn cell_defaults(&self) -> &TableRowCellDefaults {
        &self.cell_defaults
    }
    pub fn set_cell_defaults(&mut self, value: TableRowCellDefaults) {
        self.cell_defaults = value
    }
    pub fn positioning(&self) -> &FloatingTablePosition {
        &self.positioning
    }
    pub fn set_positioning(&mut self, value: FloatingTablePosition) {
        self.positioning = value
    }
    pub fn borders(&self) -> &TableRowBorders {
        &self.borders
    }
    pub fn shading(&self) -> TableShading {
        self.shading
    }
    pub fn set_borders(&mut self, value: TableRowBorders) {
        self.borders = value
    }
    pub fn set_shading(&mut self, value: TableShading) {
        self.shading = value
    }
    pub fn geometry(&self) -> TableRowGeometry {
        self.geometry
    }
    pub fn set_geometry(&mut self, value: TableRowGeometry) {
        self.geometry = value
    }
    pub const fn autoformat_flags(&self) -> TableAutoformatFlags {
        self.autoformat_flags
    }
    pub fn set_autoformat_flags(&mut self, value: TableAutoformatFlags) {
        self.autoformat_flags = value
    }
    pub const fn banding(&self) -> TableRowBanding {
        self.banding
    }
    pub fn set_banding(&mut self, value: TableRowBanding) {
        self.banding = value
    }
    /// Author/date metadata for the revision that changed this row.
    pub const fn revision(&self) -> crate::RevisionMetadata {
        self.revision
    }
    pub fn set_revision(&mut self, value: crate::RevisionMetadata) {
        self.revision = value
    }
}

impl<'a> Default for Row<'a> {
    fn default() -> Self {
        Self::new()
    }
}

/// A table cell.
#[derive(Debug, Clone)]
pub struct Cell<'a> {
    /// Cell text content
    text: Cow<'a, str>,
    padding: TableEdgeDistances,
    spacing: TableEdgeDistances,
    layout: TableCellLayout,
    borders: TableCellBorders,
    shading: TableShading,
    merge: TableCellMergeState,
    right_boundary: Option<i32>,
    preferred_width: Option<TablePreferredWidth>,
    /// Tracked-change revision attached to this cell, if any.
    revision: Option<CellRevision>,
    nested_tables: Vec<CellNestedTable<'a>>,
    shapes: Vec<crate::Shape<'a>>,
    shape_groups: Vec<crate::ShapeGroup<'a>>,
    drawing_order: Vec<crate::StoryDrawing>,
    story_events: Vec<CellStoryEvent>,
}

/// A nested table inserted between UTF-8 text bytes in its containing cell.
#[derive(Debug, Clone)]
pub struct CellNestedTable<'a> {
    pub text_offset: usize,
    pub table: Table<'a>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CellStoryReference {
    pub index: usize,
    pub position: usize,
}

/// A row/cell coordinate inside a document table or nested table.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TableCellCoordinate {
    pub table_index: usize,
    pub row_index: usize,
    pub cell_index: usize,
}

/// A stable route from a document table to an outer or nested cell.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TableCellPath {
    pub root: TableCellCoordinate,
    pub nested: Vec<TableCellCoordinate>,
}

impl TableCellPath {
    pub fn outer(table_index: usize, row_index: usize, cell_index: usize) -> Self {
        Self {
            root: TableCellCoordinate {
                table_index,
                row_index,
                cell_index,
            },
            nested: Vec::new(),
        }
    }

    pub fn with_nested(mut self, coordinate: TableCellCoordinate) -> Self {
        self.nested.push(coordinate);
        self
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CellStoryEvent {
    NestedTable(usize),
    Drawing(crate::StoryDrawing),
    Field(crate::StoryField),
    PageBreak(crate::PageBreak),
    ColumnBreak(crate::ColumnBreak),
    NavigationEntry(CellStoryReference),
    RevisionStart(CellStoryReference),
    RevisionEnd(CellStoryReference),
    RevisionDeletion(CellStoryReference),
}

impl<'a> Cell<'a> {
    /// Create a new cell.
    pub fn new(text: Cow<'a, str>) -> Self {
        Self {
            text,
            padding: TableEdgeDistances::default(),
            spacing: TableEdgeDistances::default(),
            layout: TableCellLayout::default(),
            borders: Default::default(),
            shading: Default::default(),
            merge: Default::default(),
            right_boundary: None,
            preferred_width: None,
            revision: None,
            nested_tables: Vec::new(),
            shapes: Vec::new(),
            shape_groups: Vec::new(),
            drawing_order: Vec::new(),
            story_events: Vec::new(),
        }
    }
    pub fn with_distances(
        text: Cow<'a, str>,
        padding: TableEdgeDistances,
        spacing: TableEdgeDistances,
    ) -> Self {
        Self {
            text,
            padding,
            spacing,
            layout: TableCellLayout::default(),
            borders: Default::default(),
            shading: Default::default(),
            merge: Default::default(),
            right_boundary: None,
            preferred_width: None,
            revision: None,
            nested_tables: Vec::new(),
            shapes: Vec::new(),
            shape_groups: Vec::new(),
            drawing_order: Vec::new(),
            story_events: Vec::new(),
        }
    }

    /// Get the cell text.
    pub fn text(&self) -> &str {
        &self.text
    }
    pub fn padding(&self) -> &TableEdgeDistances {
        &self.padding
    }
    pub fn spacing(&self) -> &TableEdgeDistances {
        &self.spacing
    }
    pub fn set_padding(&mut self, value: TableEdgeDistances) {
        self.padding = value
    }
    pub fn set_spacing(&mut self, value: TableEdgeDistances) {
        self.spacing = value
    }
    pub fn layout(&self) -> &TableCellLayout {
        &self.layout
    }
    pub fn set_layout(&mut self, value: TableCellLayout) {
        self.layout = value
    }
    pub fn borders(&self) -> &TableCellBorders {
        &self.borders
    }
    pub fn shading(&self) -> TableShading {
        self.shading
    }
    pub fn set_borders(&mut self, value: TableCellBorders) {
        self.borders = value
    }
    pub fn set_shading(&mut self, value: TableShading) {
        self.shading = value
    }
    pub fn merge(&self) -> TableCellMergeState {
        self.merge
    }
    pub fn set_merge(&mut self, value: TableCellMergeState) {
        self.merge = value
    }
    pub fn right_boundary(&self) -> Option<i32> {
        self.right_boundary
    }
    pub fn set_right_boundary(&mut self, value: Option<i32>) {
        self.right_boundary = value
    }
    pub fn preferred_width(&self) -> Option<TablePreferredWidth> {
        self.preferred_width
    }
    pub fn set_preferred_width(&mut self, value: Option<TablePreferredWidth>) {
        self.preferred_width = value
    }
    /// Tracked-change revision attached to this cell, if any.
    pub const fn revision(&self) -> Option<CellRevision> {
        self.revision
    }
    pub fn set_revision(&mut self, value: Option<CellRevision>) {
        self.revision = value
    }
    pub fn nested_tables(&self) -> &[CellNestedTable<'a>] {
        &self.nested_tables
    }
    pub fn add_nested_table(
        &mut self,
        text_offset: usize,
        table: Table<'a>,
    ) -> crate::RtfResult<()> {
        let last_story_position = self.last_story_position()?;
        if text_offset > self.text.len()
            || !self.text.is_char_boundary(text_offset)
            || self
                .nested_tables
                .last()
                .is_some_and(|entry| entry.text_offset > text_offset)
            || last_story_position.is_some_and(|position| position > text_offset)
        {
            return Err(crate::RtfError::MalformedDocument(
                "invalid nested-table text insertion offset".to_string(),
            ));
        }
        crate::error::try_reserve_one(&mut self.nested_tables, "table-cell nested tables")?;
        crate::error::try_reserve_one(&mut self.story_events, "table-cell story events")?;
        let index = self.nested_tables.len();
        self.nested_tables
            .push(CellNestedTable { text_offset, table });
        self.story_events.push(CellStoryEvent::NestedTable(index));
        Ok(())
    }
    pub fn nested_tables_mut(&mut self) -> &mut Vec<CellNestedTable<'a>> {
        &mut self.nested_tables
    }
    pub fn shapes(&self) -> &[crate::Shape<'a>] {
        &self.shapes
    }
    pub fn shape_groups(&self) -> &[crate::ShapeGroup<'a>] {
        &self.shape_groups
    }
    pub fn drawing_order(&self) -> &[crate::StoryDrawing] {
        &self.drawing_order
    }
    pub fn story_events(&self) -> &[CellStoryEvent] {
        &self.story_events
    }
    fn event_position(&self, event: CellStoryEvent) -> crate::RtfResult<usize> {
        Self::event_position_in(&self.nested_tables, &self.shapes, &self.shape_groups, event)
    }

    fn event_position_in(
        nested_tables: &[CellNestedTable<'_>],
        shapes: &[crate::Shape<'_>],
        shape_groups: &[crate::ShapeGroup<'_>],
        event: CellStoryEvent,
    ) -> crate::RtfResult<usize> {
        match event {
            CellStoryEvent::NestedTable(index) => {
                nested_tables.get(index).map(|nested| nested.text_offset)
            },
            CellStoryEvent::Drawing(crate::StoryDrawing::Shape(index)) => {
                shapes.get(index).map(|shape| shape.position)
            },
            CellStoryEvent::Drawing(crate::StoryDrawing::ShapeGroup(index)) => {
                shape_groups.get(index).map(|group| group.position)
            },
            CellStoryEvent::Field(field) => Some(field.position),
            CellStoryEvent::PageBreak(page_break) => Some(page_break.position),
            CellStoryEvent::ColumnBreak(column_break) => Some(column_break.position),
            CellStoryEvent::NavigationEntry(reference)
            | CellStoryEvent::RevisionStart(reference)
            | CellStoryEvent::RevisionEnd(reference)
            | CellStoryEvent::RevisionDeletion(reference) => Some(reference.position),
        }
        .ok_or_else(|| {
            crate::RtfError::MalformedDocument(
                "RTF table-cell story event references missing metadata".to_string(),
            )
        })
    }

    fn last_story_position(&self) -> crate::RtfResult<Option<usize>> {
        self.story_events
            .last()
            .copied()
            .map(|event| self.event_position(event))
            .transpose()
    }

    fn validate_story_content(
        text: &str,
        nested_tables: &[CellNestedTable<'_>],
        shapes: &[crate::Shape<'_>],
        shape_groups: &[crate::ShapeGroup<'_>],
        drawing_order: &[crate::StoryDrawing],
        story_events: &[CellStoryEvent],
    ) -> crate::RtfResult<()> {
        crate::shape::validate_story_drawings(
            text,
            shapes,
            shape_groups,
            drawing_order,
            "table cell",
        )?;

        let mut saw_nested = Vec::new();
        crate::error::try_reserve_additional(
            &mut saw_nested,
            nested_tables.len(),
            "table-cell nested-table references",
        )?;
        saw_nested.resize(nested_tables.len(), false);

        let field_count = story_events
            .iter()
            .filter(|event| matches!(event, CellStoryEvent::Field(_)))
            .count();
        let mut saw_fields = HashSet::new();
        saw_fields
            .try_reserve(field_count)
            .map_err(|_| crate::RtfError::AllocationFailed {
                resource: "table-cell field references",
                requested: field_count.saturating_mul(std::mem::size_of::<usize>()),
            })?;

        let mut drawings = drawing_order.iter().copied();
        let mut previous = None;
        for event in story_events {
            let position = match *event {
                CellStoryEvent::NestedTable(index) => {
                    let seen = saw_nested.get_mut(index).ok_or_else(|| {
                        crate::RtfError::MalformedDocument(
                            "RTF table-cell story order has an invalid nested-table reference"
                                .to_string(),
                        )
                    })?;
                    if std::mem::replace(seen, true) {
                        return Err(crate::RtfError::MalformedDocument(
                            "RTF table-cell story order has a duplicate nested-table reference"
                                .to_string(),
                        ));
                    }
                    Self::event_position_in(nested_tables, shapes, shape_groups, *event)?
                },
                CellStoryEvent::Drawing(drawing) => {
                    if drawings.next() != Some(drawing) {
                        return Err(crate::RtfError::MalformedDocument(
                            "RTF table-cell story order changes drawing order".to_string(),
                        ));
                    }
                    Self::event_position_in(nested_tables, shapes, shape_groups, *event)?
                },
                CellStoryEvent::Field(field) => {
                    if !saw_fields.insert(field.field_index) {
                        return Err(crate::RtfError::MalformedDocument(
                            "RTF table-cell story order has a duplicate field reference"
                                .to_string(),
                        ));
                    }
                    field.position
                },
                CellStoryEvent::PageBreak(page_break) => page_break.position,
                CellStoryEvent::ColumnBreak(column_break) => column_break.position,
                CellStoryEvent::NavigationEntry(reference)
                | CellStoryEvent::RevisionStart(reference)
                | CellStoryEvent::RevisionEnd(reference)
                | CellStoryEvent::RevisionDeletion(reference) => reference.position,
            };
            if text.get(position..position).is_none() {
                return Err(crate::RtfError::MalformedDocument(
                    "RTF table-cell story event is outside a UTF-8 boundary".to_string(),
                ));
            }
            if previous.is_some_and(|value| value > position) {
                return Err(crate::RtfError::MalformedDocument(
                    "RTF table-cell story order moves backwards".to_string(),
                ));
            }
            previous = Some(position);
        }
        if saw_nested.iter().any(|seen| !*seen) || drawings.next().is_some() {
            return Err(crate::RtfError::MalformedDocument(
                "RTF table-cell story order is incomplete or changes drawing order".to_string(),
            ));
        }
        Ok(())
    }

    pub fn validate_drawings(&self) -> crate::RtfResult<()> {
        Self::validate_story_content(
            self.text.as_ref(),
            &self.nested_tables,
            &self.shapes,
            &self.shape_groups,
            &self.drawing_order,
            &self.story_events,
        )
    }
    pub fn push_shape(&mut self, shape: crate::Shape<'a>) -> crate::RtfResult<()> {
        let position = shape.position;
        if self.shapes.len() >= crate::shape::MAX_SHAPES_PER_GROUP {
            return Err(crate::RtfError::MalformedDocument(
                "RTF table-cell shape count exceeds the safety limit".to_string(),
            ));
        }
        crate::shape::validate_story_drawings(
            self.text.as_ref(),
            &self.shapes,
            &self.shape_groups,
            &self.drawing_order,
            "table cell",
        )?;
        shape.validate()?;
        let last_drawing_position = self
            .drawing_order
            .last()
            .copied()
            .map(|drawing| self.event_position(CellStoryEvent::Drawing(drawing)))
            .transpose()?;
        let last_story_position = self.last_story_position()?;
        if shape.is_background
            || self.text.get(position..position).is_none()
            || self
                .shapes
                .last()
                .is_some_and(|previous| previous.position > position)
            || last_drawing_position.is_some_and(|previous| previous > position)
            || last_story_position.is_some_and(|previous| previous > position)
        {
            return Err(crate::RtfError::MalformedDocument(
                "RTF table-cell shape is outside or out of story order".to_string(),
            ));
        }
        let drawing = crate::StoryDrawing::Shape(self.shapes.len());
        crate::error::try_reserve_one(&mut self.shapes, "table-cell shapes")?;
        crate::error::try_reserve_one(&mut self.drawing_order, "table-cell drawing order")?;
        crate::error::try_reserve_one(&mut self.story_events, "table-cell story events")?;
        self.shapes.push(shape);
        self.drawing_order.push(drawing);
        self.story_events.push(CellStoryEvent::Drawing(drawing));
        Ok(())
    }
    pub fn push_shape_group(&mut self, group: crate::ShapeGroup<'a>) -> crate::RtfResult<()> {
        let position = group.position;
        if self.shape_groups.len() >= crate::shape::MAX_GROUPS_PER_GROUP {
            return Err(crate::RtfError::MalformedDocument(
                "RTF table-cell shape-group count exceeds the safety limit".to_string(),
            ));
        }
        crate::shape::validate_story_drawings(
            self.text.as_ref(),
            &self.shapes,
            &self.shape_groups,
            &self.drawing_order,
            "table cell",
        )?;
        group.validate()?;
        let last_drawing_position = self
            .drawing_order
            .last()
            .copied()
            .map(|drawing| self.event_position(CellStoryEvent::Drawing(drawing)))
            .transpose()?;
        let last_story_position = self.last_story_position()?;
        if self.text.get(position..position).is_none()
            || self
                .shape_groups
                .last()
                .is_some_and(|previous| previous.position > position)
            || last_drawing_position.is_some_and(|previous| previous > position)
            || last_story_position.is_some_and(|previous| previous > position)
        {
            return Err(crate::RtfError::MalformedDocument(
                "RTF table-cell shape group is outside or out of story order".to_string(),
            ));
        }
        let drawing = crate::StoryDrawing::ShapeGroup(self.shape_groups.len());
        crate::error::try_reserve_one(&mut self.shape_groups, "table-cell shape groups")?;
        crate::error::try_reserve_one(&mut self.drawing_order, "table-cell drawing order")?;
        crate::error::try_reserve_one(&mut self.story_events, "table-cell story events")?;
        self.shape_groups.push(group);
        self.drawing_order.push(drawing);
        self.story_events.push(CellStoryEvent::Drawing(drawing));
        Ok(())
    }
    pub fn clear_drawings(&mut self) {
        self.shapes.clear();
        self.shape_groups.clear();
        self.drawing_order.clear();
        self.story_events
            .retain(|event| !matches!(event, CellStoryEvent::Drawing(_)));
    }
    pub fn page_breaks(&self) -> impl Iterator<Item = &crate::PageBreak> {
        self.story_events.iter().filter_map(|event| match event {
            CellStoryEvent::PageBreak(page_break) => Some(page_break),
            _ => None,
        })
    }
    pub fn push_page_break(&mut self, position: usize) -> crate::RtfResult<()> {
        self.push_ordered_story_events(
            [CellStoryEvent::PageBreak(crate::PageBreak::new(position))],
            "invalid table-cell page-break insertion offset",
        )
    }
    pub fn clear_page_breaks(&mut self) {
        self.story_events
            .retain(|event| !matches!(event, CellStoryEvent::PageBreak(_)));
    }
    pub fn column_breaks(&self) -> impl Iterator<Item = &crate::ColumnBreak> {
        self.story_events.iter().filter_map(|event| match event {
            CellStoryEvent::ColumnBreak(column_break) => Some(column_break),
            _ => None,
        })
    }
    pub fn push_column_break(&mut self, position: usize) -> crate::RtfResult<()> {
        self.push_ordered_story_events(
            [CellStoryEvent::ColumnBreak(crate::ColumnBreak::new(
                position,
            ))],
            "invalid table-cell column-break insertion offset",
        )
    }
    pub fn clear_column_breaks(&mut self) {
        self.story_events
            .retain(|event| !matches!(event, CellStoryEvent::ColumnBreak(_)));
    }
    pub fn navigation_entry_references(&self) -> impl Iterator<Item = &CellStoryReference> {
        self.story_events.iter().filter_map(|event| match event {
            CellStoryEvent::NavigationEntry(reference) => Some(reference),
            _ => None,
        })
    }
    pub fn revision_events(&self) -> impl Iterator<Item = &CellStoryEvent> {
        self.story_events.iter().filter(|event| {
            matches!(
                event,
                CellStoryEvent::RevisionStart(_)
                    | CellStoryEvent::RevisionEnd(_)
                    | CellStoryEvent::RevisionDeletion(_)
            )
        })
    }
    pub fn push_navigation_entry_reference(
        &mut self,
        index: usize,
        position: usize,
    ) -> crate::RtfResult<()> {
        self.push_positional_event(CellStoryEvent::NavigationEntry(CellStoryReference {
            index,
            position,
        }))
    }
    pub fn push_insertion_revision_reference(
        &mut self,
        index: usize,
        position: usize,
        range_end: usize,
    ) -> crate::RtfResult<()> {
        if range_end <= position || self.text.get(position..range_end).is_none() {
            return Err(crate::RtfError::MalformedDocument(
                "invalid table-cell insertion revision range".to_string(),
            ));
        }
        self.push_ordered_story_events(
            [
                CellStoryEvent::RevisionStart(CellStoryReference { index, position }),
                CellStoryEvent::RevisionEnd(CellStoryReference {
                    index,
                    position: range_end,
                }),
            ],
            "invalid or out-of-order table-cell insertion revision",
        )
    }
    pub fn push_deletion_revision_reference(
        &mut self,
        index: usize,
        position: usize,
    ) -> crate::RtfResult<()> {
        self.push_positional_event(CellStoryEvent::RevisionDeletion(CellStoryReference {
            index,
            position,
        }))
    }
    fn push_positional_event(&mut self, event: CellStoryEvent) -> crate::RtfResult<()> {
        self.push_ordered_story_events([event], "invalid or out-of-order table-cell story event")
    }
    fn push_ordered_story_events<const N: usize>(
        &mut self,
        events: [CellStoryEvent; N],
        error_message: &'static str,
    ) -> crate::RtfResult<()> {
        let mut previous_position = self.last_story_position()?;
        for event in &events {
            let position = self.event_position(*event)?;
            if self.text.get(position..position).is_none()
                || previous_position.is_some_and(|previous| previous > position)
            {
                return Err(crate::RtfError::MalformedDocument(
                    error_message.to_string(),
                ));
            }
            previous_position = Some(position);
        }
        crate::error::try_reserve_additional(&mut self.story_events, N, "table-cell story events")?;
        self.story_events.extend(events);
        Ok(())
    }
    pub fn set_text(&mut self, text: Cow<'a, str>) -> crate::RtfResult<()> {
        Self::validate_story_content(
            text.as_ref(),
            &self.nested_tables,
            &self.shapes,
            &self.shape_groups,
            &self.drawing_order,
            &self.story_events,
        )?;
        self.text = text;
        Ok(())
    }
    pub fn clear_navigation_entry_references(&mut self) {
        self.story_events
            .retain(|event| !matches!(event, CellStoryEvent::NavigationEntry(_)));
        for nested in &mut self.nested_tables {
            nested.table.clear_navigation_entry_references();
        }
    }
    pub fn clear_revision_references(&mut self) {
        self.story_events.retain(|event| {
            !matches!(
                event,
                CellStoryEvent::RevisionStart(_)
                    | CellStoryEvent::RevisionEnd(_)
                    | CellStoryEvent::RevisionDeletion(_)
            )
        });
        for nested in &mut self.nested_tables {
            nested.table.clear_revision_references();
        }
    }
    pub(crate) fn set_story_content(
        &mut self,
        shapes: Vec<crate::Shape<'a>>,
        shape_groups: Vec<crate::ShapeGroup<'a>>,
        drawing_order: Vec<crate::StoryDrawing>,
        story_events: Vec<CellStoryEvent>,
    ) -> crate::RtfResult<()> {
        Self::validate_story_content(
            self.text.as_ref(),
            &self.nested_tables,
            &shapes,
            &shape_groups,
            &drawing_order,
            &story_events,
        )?;
        self.shapes = shapes;
        self.shape_groups = shape_groups;
        self.drawing_order = drawing_order;
        self.story_events = story_events;
        Ok(())
    }
}

impl<'a> Row<'a> {
    pub fn cells_mut(&mut self) -> &mut [Cell<'a>] {
        &mut self.cells
    }
}

impl Table<'_> {
    /// Detach this table from the source buffer, cloning any borrowed text.
    ///
    /// Every row, cell, nested table, and anchored drawing is carried over, so
    /// the result is structurally identical to the parsed table rather than a
    /// text-only skeleton. Cells whose text is already owned move without
    /// copying.
    pub fn into_owned(self) -> Table<'static> {
        Table {
            rows: self.rows.into_iter().map(Row::into_owned).collect(),
            direction: self.direction,
        }
    }
}

impl Row<'_> {
    /// Detach this row from the source buffer, cloning any borrowed text.
    ///
    /// See [`Table::into_owned`]; all row-level geometry, borders, shading, and
    /// autoformat state is preserved.
    pub fn into_owned(self) -> Row<'static> {
        Row {
            cells: self.cells.into_iter().map(Cell::into_owned).collect(),
            table_style: self.table_style,
            table_rsid: self.table_rsid,
            direction: self.direction,
            layout: self.layout,
            padding: self.padding,
            spacing: self.spacing,
            cell_defaults: self.cell_defaults,
            positioning: self.positioning,
            borders: self.borders,
            shading: self.shading,
            geometry: self.geometry,
            autoformat_flags: self.autoformat_flags,
            banding: self.banding,
            revision: self.revision,
        }
    }
}

impl Cell<'_> {
    /// Detach this cell from the source buffer, cloning any borrowed text.
    ///
    /// See [`Table::into_owned`]; merge roles, borders, shading, layout,
    /// boundaries, preferred width, revision, nested tables, shapes, and the
    /// story event ordering are all preserved.
    pub fn into_owned(self) -> Cell<'static> {
        Cell {
            text: Cow::Owned(self.text.into_owned()),
            padding: self.padding,
            spacing: self.spacing,
            layout: self.layout,
            borders: self.borders,
            shading: self.shading,
            merge: self.merge,
            right_boundary: self.right_boundary,
            preferred_width: self.preferred_width,
            revision: self.revision,
            nested_tables: self
                .nested_tables
                .into_iter()
                .map(CellNestedTable::into_owned)
                .collect(),
            shapes: self
                .shapes
                .into_iter()
                .map(crate::Shape::into_owned)
                .collect(),
            shape_groups: self
                .shape_groups
                .into_iter()
                .map(crate::ShapeGroup::into_owned)
                .collect(),
            drawing_order: self.drawing_order,
            story_events: self.story_events,
        }
    }
}

impl CellNestedTable<'_> {
    /// Detach this nested table from the source buffer, keeping its anchor.
    pub fn into_owned(self) -> CellNestedTable<'static> {
        CellNestedTable {
            text_offset: self.text_offset,
            table: self.table.into_owned(),
        }
    }
}

impl<'a> Table<'a> {
    pub fn rows_mut(&mut self) -> &mut [Row<'a>] {
        &mut self.rows
    }
    pub(crate) fn clear_navigation_entry_references(&mut self) {
        for row in &mut self.rows {
            for cell in &mut row.cells {
                cell.clear_navigation_entry_references();
            }
        }
    }
    pub(crate) fn clear_revision_references(&mut self) {
        for row in &mut self.rows {
            for cell in &mut row.cells {
                cell.clear_revision_references();
            }
        }
    }
}

impl Table<'_> {
    pub(crate) fn validate_merges(&self) -> Result<(), String> {
        let mut active_vertical: BTreeSet<(i32, i32)> = BTreeSet::new();
        for row in self.rows() {
            let mut horizontal_open = false;
            let mut next_vertical: BTreeSet<(i32, i32)> = BTreeSet::new();
            let mut left = 0i32;
            for (index, cell) in row.cells().iter().enumerate() {
                let right = cell.right_boundary().unwrap_or(2880 * ((index + 1) as i32));
                match cell.merge().horizontal {
                    None => horizontal_open = false,
                    Some(TableCellMergeRole::First) => horizontal_open = true,
                    Some(TableCellMergeRole::Continuation) => {
                        if !horizontal_open {
                            return Err(
                                "RTF horizontal merge continuation does not follow a first cell"
                                    .to_string(),
                            );
                        }
                    },
                }
                let span = (left, right);
                match cell.merge().vertical {
                    None => {},
                    Some(TableCellMergeRole::First) => {
                        next_vertical.insert(span);
                    },
                    Some(TableCellMergeRole::Continuation) => {
                        if !active_vertical.contains(&span) {
                            return Err("RTF vertical merge continuation does not match a cell boundary in the preceding row".to_string());
                        }
                        next_vertical.insert(span);
                    },
                }
                left = right;
            }
            active_vertical = next_vertical;
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn malformed_story_references_fail_without_mutating_the_cell() {
        let mut cell = Cell::new(Cow::Borrowed("abc"));
        cell.add_nested_table(1, Table::new()).unwrap();
        cell.nested_tables_mut().clear();
        let events_before = cell.story_events().to_vec();

        assert!(cell.validate_drawings().is_err());
        assert!(cell.push_page_break(1).is_err());
        assert!(cell.add_nested_table(2, Table::new()).is_err());
        assert_eq!(cell.story_events(), events_before);
        assert!(cell.nested_tables().is_empty());

        let text_before = cell.text().to_string();
        assert!(cell.set_text(Cow::Borrowed("replacement")).is_err());
        assert_eq!(cell.text(), text_before);

        let mut cell = Cell::new(Cow::Borrowed("abc"));
        let mut shape = crate::Shape::new(crate::ShapeType::Rectangle);
        shape.position = 1;
        cell.shapes.push(shape);
        cell.drawing_order.push(crate::StoryDrawing::Shape(0));
        cell.story_events
            .push(CellStoryEvent::Drawing(crate::StoryDrawing::Shape(
                usize::MAX,
            )));
        let events_before = cell.story_events.clone();
        assert!(cell.validate_drawings().is_err());
        assert!(cell.push_column_break(1).is_err());
        assert_eq!(cell.story_events, events_before);

        let mut cell = Cell::new(Cow::Borrowed("abc"));
        let mut group = crate::ShapeGroup::new();
        group.position = 1;
        cell.shape_groups.push(group);
        cell.drawing_order.push(crate::StoryDrawing::ShapeGroup(0));
        cell.story_events
            .push(CellStoryEvent::Drawing(crate::StoryDrawing::ShapeGroup(
                usize::MAX,
            )));
        let events_before = cell.story_events.clone();
        assert!(cell.validate_drawings().is_err());
        assert!(cell.push_navigation_entry_reference(0, 1).is_err());
        assert_eq!(cell.story_events, events_before);
    }

    #[test]
    fn story_content_replacement_and_text_edits_are_atomic() {
        let mut cell = Cell::new(Cow::Borrowed("ab"));
        let mut shape = crate::Shape::new(crate::ShapeType::Rectangle);
        shape.position = 1;
        cell.push_shape(shape).unwrap();
        let shapes_before = cell.shapes.clone();
        let order_before = cell.drawing_order.clone();
        let events_before = cell.story_events.clone();

        assert!(
            cell.set_story_content(
                Vec::new(),
                Vec::new(),
                Vec::new(),
                vec![CellStoryEvent::Drawing(crate::StoryDrawing::Shape(
                    usize::MAX,
                ))],
            )
            .is_err()
        );
        assert_eq!(cell.shapes, shapes_before);
        assert_eq!(cell.drawing_order, order_before);
        assert_eq!(cell.story_events, events_before);

        let mut cell = Cell::new(Cow::Borrowed("ab"));
        cell.add_nested_table(1, Table::new()).unwrap();
        assert!(cell.set_text(Cow::Borrowed("你")).is_err());
        assert_eq!(cell.text(), "ab");
        cell.validate_drawings().unwrap();
    }

    #[test]
    fn drawing_append_preserves_existing_owned_payload_allocations() {
        let mut cell = Cell::new(Cow::Borrowed("abc"));
        let mut first = crate::Shape::new(crate::ShapeType::Rectangle);
        first.position = 1;
        first.name = Cow::Owned("first shape".repeat(1_024));
        cell.push_shape(first).unwrap();
        let name_pointer = cell.shapes().first().unwrap().name.as_ptr();

        let mut second = crate::Shape::new(crate::ShapeType::Ellipse);
        second.position = 1;
        cell.push_shape(second).unwrap();

        assert_eq!(cell.shapes().first().unwrap().name.as_ptr(), name_pointer);
        assert_eq!(
            cell.drawing_order(),
            &[crate::StoryDrawing::Shape(0), crate::StoryDrawing::Shape(1)]
        );
        cell.validate_drawings().unwrap();
    }
}
