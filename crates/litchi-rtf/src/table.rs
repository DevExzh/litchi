//! RTF table support.
//!
//! This module provides basic table parsing for RTF documents.
//! RTF tables use a complex row-based model with cell boundaries.

use crate::TextDirection;
use std::borrow::Cow;
use std::collections::BTreeSet;

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

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum TableBorderTarget {
    Row(TableRowBorderSide),
    Cell(TableCellBorderSide),
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
    pub foreground_color: Option<crate::ColorRef>,
    pub background_color: Option<crate::ColorRef>,
    pub pattern: Option<crate::ShadingPattern>,
    pub pattern_index: Option<u16>,
}
impl TableShading {
    pub fn validate(&self) -> crate::RtfResult<()> {
        if self.amount.is_some_and(|value| value > 10_000) {
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
    /// Explicit row direction.
    direction: Option<TextDirection>,
    layout: TableRowLayout,
    padding: TableEdgeDistances,
    spacing: TableEdgeDistances,
    positioning: FloatingTablePosition,
    borders: TableRowBorders,
    shading: TableShading,
    geometry: TableRowGeometry,
    autoformat_flags: TableAutoformatFlags,
    banding: TableRowBanding,
}

impl<'a> Row<'a> {
    /// Create a new row.
    pub fn new() -> Self {
        Self {
            cells: Vec::new(),
            table_style: None,
            direction: None,
            layout: TableRowLayout::default(),
            padding: TableEdgeDistances::default(),
            spacing: TableEdgeDistances::default(),
            positioning: FloatingTablePosition::default(),
            borders: Default::default(),
            shading: Default::default(),
            geometry: Default::default(),
            autoformat_flags: Default::default(),
            banding: Default::default(),
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
    pub fn nested_tables(&self) -> &[CellNestedTable<'a>] {
        &self.nested_tables
    }
    pub fn add_nested_table(
        &mut self,
        text_offset: usize,
        table: Table<'a>,
    ) -> crate::RtfResult<()> {
        if text_offset > self.text.len()
            || !self.text.is_char_boundary(text_offset)
            || self
                .nested_tables
                .last()
                .is_some_and(|entry| entry.text_offset > text_offset)
            || self
                .story_events
                .last()
                .is_some_and(|event| self.event_position(*event) > text_offset)
        {
            return Err(crate::RtfError::MalformedDocument(
                "invalid nested-table text insertion offset".to_string(),
            ));
        }
        self.story_events
            .push(CellStoryEvent::NestedTable(self.nested_tables.len()));
        self.nested_tables
            .push(CellNestedTable { text_offset, table });
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
    fn event_position(&self, event: CellStoryEvent) -> usize {
        match event {
            CellStoryEvent::NestedTable(index) => self.nested_tables[index].text_offset,
            CellStoryEvent::Drawing(crate::StoryDrawing::Shape(index)) => {
                self.shapes[index].position
            },
            CellStoryEvent::Drawing(crate::StoryDrawing::ShapeGroup(index)) => {
                self.shape_groups[index].position
            },
            CellStoryEvent::Field(field) => field.position,
            CellStoryEvent::PageBreak(page_break) => page_break.position,
            CellStoryEvent::ColumnBreak(column_break) => column_break.position,
            CellStoryEvent::NavigationEntry(reference)
            | CellStoryEvent::RevisionStart(reference)
            | CellStoryEvent::RevisionEnd(reference)
            | CellStoryEvent::RevisionDeletion(reference) => reference.position,
        }
    }
    pub fn validate_drawings(&self) -> crate::RtfResult<()> {
        crate::shape::validate_story_drawings(
            self.text.as_ref(),
            &self.shapes,
            &self.shape_groups,
            &self.drawing_order,
            "table cell",
        )?;
        let mut saw_nested = vec![false; self.nested_tables.len()];
        let mut saw_fields = std::collections::BTreeSet::new();
        let mut drawings = Vec::with_capacity(self.drawing_order.len());
        let mut previous = None;
        for event in &self.story_events {
            let position = match *event {
                CellStoryEvent::NestedTable(index)
                    if index < self.nested_tables.len() && !saw_nested[index] =>
                {
                    saw_nested[index] = true;
                    self.nested_tables[index].text_offset
                },
                CellStoryEvent::Drawing(drawing) => {
                    drawings.push(drawing);
                    match drawing {
                        crate::StoryDrawing::Shape(index) if index < self.shapes.len() => {
                            self.shapes[index].position
                        },
                        crate::StoryDrawing::ShapeGroup(index)
                            if index < self.shape_groups.len() =>
                        {
                            self.shape_groups[index].position
                        },
                        _ => {
                            return Err(crate::RtfError::MalformedDocument(
                                "RTF table-cell story order has an invalid drawing reference"
                                    .to_string(),
                            ));
                        },
                    }
                },
                CellStoryEvent::Field(field)
                    if saw_fields.insert(field.field_index)
                        && self.text.get(field.position..field.position).is_some() =>
                {
                    field.position
                },
                CellStoryEvent::PageBreak(page_break)
                    if self
                        .text
                        .get(page_break.position..page_break.position)
                        .is_some() =>
                {
                    page_break.position
                },
                CellStoryEvent::ColumnBreak(column_break)
                    if self
                        .text
                        .get(column_break.position..column_break.position)
                        .is_some() =>
                {
                    column_break.position
                },
                CellStoryEvent::NavigationEntry(reference)
                | CellStoryEvent::RevisionStart(reference)
                | CellStoryEvent::RevisionEnd(reference)
                | CellStoryEvent::RevisionDeletion(reference)
                    if self
                        .text
                        .get(reference.position..reference.position)
                        .is_some() =>
                {
                    reference.position
                },
                _ => {
                    return Err(crate::RtfError::MalformedDocument(
                        "RTF table-cell story order has an invalid or duplicate reference"
                            .to_string(),
                    ));
                },
            };
            if previous.is_some_and(|value| value > position) {
                return Err(crate::RtfError::MalformedDocument(
                    "RTF table-cell story order moves backwards".to_string(),
                ));
            }
            previous = Some(position);
        }
        if saw_nested.iter().any(|seen| !*seen) || drawings != self.drawing_order {
            return Err(crate::RtfError::MalformedDocument(
                "RTF table-cell story order is incomplete or changes drawing order".to_string(),
            ));
        }
        Ok(())
    }
    pub fn push_shape(&mut self, shape: crate::Shape<'a>) -> crate::RtfResult<()> {
        let drawing = crate::StoryDrawing::Shape(self.shapes.len());
        let mut shapes = self.shapes.clone();
        let mut order = self.drawing_order.clone();
        order.push(drawing);
        shapes.push(shape);
        crate::shape::validate_story_drawings(
            self.text.as_ref(),
            &shapes,
            &self.shape_groups,
            &order,
            "table cell",
        )?;
        let position = shapes.last().unwrap().position;
        if self
            .story_events
            .last()
            .is_some_and(|event| self.event_position(*event) > position)
        {
            return Err(crate::RtfError::MalformedDocument(
                "RTF table-cell story order moves backwards".to_string(),
            ));
        }
        self.shapes = shapes;
        self.drawing_order = order;
        self.story_events.push(CellStoryEvent::Drawing(drawing));
        Ok(())
    }
    pub fn push_shape_group(&mut self, group: crate::ShapeGroup<'a>) -> crate::RtfResult<()> {
        let drawing = crate::StoryDrawing::ShapeGroup(self.shape_groups.len());
        let mut groups = self.shape_groups.clone();
        let mut order = self.drawing_order.clone();
        order.push(drawing);
        groups.push(group);
        crate::shape::validate_story_drawings(
            self.text.as_ref(),
            &self.shapes,
            &groups,
            &order,
            "table cell",
        )?;
        let position = groups.last().unwrap().position;
        if self
            .story_events
            .last()
            .is_some_and(|event| self.event_position(*event) > position)
        {
            return Err(crate::RtfError::MalformedDocument(
                "RTF table-cell story order moves backwards".to_string(),
            ));
        }
        self.shape_groups = groups;
        self.drawing_order = order;
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
        if self.text.get(position..position).is_none()
            || self
                .story_events
                .last()
                .is_some_and(|event| self.event_position(*event) > position)
        {
            return Err(crate::RtfError::MalformedDocument(
                "invalid table-cell page-break insertion offset".to_string(),
            ));
        }
        self.story_events
            .push(CellStoryEvent::PageBreak(crate::PageBreak::new(position)));
        Ok(())
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
        if self.text.get(position..position).is_none()
            || self
                .story_events
                .last()
                .is_some_and(|event| self.event_position(*event) > position)
        {
            return Err(crate::RtfError::MalformedDocument(
                "invalid table-cell column-break insertion offset".to_string(),
            ));
        }
        self.story_events
            .push(CellStoryEvent::ColumnBreak(crate::ColumnBreak::new(position)));
        Ok(())
    }
    pub fn clear_column_breaks(&mut self) {
        self.story_events
            .retain(|event| !matches!(event, CellStoryEvent::ColumnBreak(_)));
    }
    pub fn navigation_entry_references(
        &self,
    ) -> impl Iterator<Item = &CellStoryReference> {
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
        self.push_positional_event(CellStoryEvent::RevisionStart(CellStoryReference {
            index,
            position,
        }))?;
        self.push_positional_event(CellStoryEvent::RevisionEnd(CellStoryReference {
            index,
            position: range_end,
        }))
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
        let position = self.event_position(event);
        if self.text.get(position..position).is_none()
            || self
                .story_events
                .last()
                .is_some_and(|previous| self.event_position(*previous) > position)
        {
            return Err(crate::RtfError::MalformedDocument(
                "invalid or out-of-order table-cell story event".to_string(),
            ));
        }
        self.story_events.push(event);
        Ok(())
    }
    pub fn set_text(&mut self, text: Cow<'a, str>) -> crate::RtfResult<()> {
        let previous = std::mem::replace(&mut self.text, text);
        if let Err(error) = self.validate_drawings() {
            self.text = previous;
            return Err(error);
        }
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
        self.shapes = shapes;
        self.shape_groups = shape_groups;
        self.drawing_order = drawing_order;
        self.story_events = story_events;
        if let Err(error) = self.validate_drawings() {
            self.shapes = Vec::new();
            self.shape_groups = Vec::new();
            self.drawing_order = Vec::new();
            self.story_events = self
                .nested_tables
                .iter()
                .enumerate()
                .map(|(index, _)| CellStoryEvent::NestedTable(index))
                .collect();
            return Err(error);
        }
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
            direction: self.direction,
            layout: self.layout,
            padding: self.padding,
            spacing: self.spacing,
            positioning: self.positioning,
            borders: self.borders,
            shading: self.shading,
            geometry: self.geometry,
            autoformat_flags: self.autoformat_flags,
            banding: self.banding,
        }
    }
}

impl Cell<'_> {
    /// Detach this cell from the source buffer, cloning any borrowed text.
    ///
    /// See [`Table::into_owned`]; merge roles, borders, shading, layout,
    /// boundaries, preferred width, nested tables, shapes, and the story event
    /// ordering are all preserved.
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
