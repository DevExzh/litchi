//! RTF table support.
//!
//! This module provides basic table parsing for RTF documents.
//! RTF tables use a complex row-based model with cell boundaries.

use crate::TextDirection;
use std::borrow::Cow;

pub const MAX_TABLE_DISTANCE_TWIPS: i32 = 31_680;
pub const MAX_FLOATING_TABLE_DISTANCE_TWIPS: i32 = 31_680;
pub const MAX_TABLE_NESTING_DEPTH: usize = 32;
pub const MAX_TABLE_CELLS_PER_ROW: usize = 4_096;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableHorizontalReference { Column, Margin, Page }

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableVerticalReference { Margin, Paragraph, Page }

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableHorizontalPosition { Offset(i32), NegativeOffset(i32), Center, Inside, Left, Outside, Right }

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableVerticalPosition { Offset(i32), NegativeOffset(i32), Bottom, Center, Inline, Inside, Outside, Top }

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct TableWrapDistances { pub left:Option<u16>, pub right:Option<u16>, pub top:Option<u16>, pub bottom:Option<u16> }
impl TableWrapDistances {
    pub(crate) fn side_mut(&mut self,edge:TableEdge)->&mut Option<u16>{match edge{TableEdge::Left=>&mut self.left,TableEdge::Right=>&mut self.right,TableEdge::Top=>&mut self.top,TableEdge::Bottom=>&mut self.bottom}}
    pub fn validate(&self)->crate::RtfResult<()>{if [self.left,self.right,self.top,self.bottom].into_iter().flatten().any(|value|i32::from(value)>MAX_FLOATING_TABLE_DISTANCE_TWIPS){return Err(crate::RtfError::MalformedDocument("RTF floating-table wrap distance is out of range".to_string()))}Ok(())}
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct FloatingTablePosition { pub horizontal_reference:Option<TableHorizontalReference>, pub horizontal_position:Option<TableHorizontalPosition>, pub vertical_reference:Option<TableVerticalReference>, pub vertical_position:Option<TableVerticalPosition>, pub no_overlap:bool, pub wrap_distances:TableWrapDistances }
impl FloatingTablePosition {
    pub fn is_empty(&self)->bool{self.horizontal_reference.is_none()&&self.horizontal_position.is_none()&&self.vertical_reference.is_none()&&self.vertical_position.is_none()&&!self.no_overlap&&self.wrap_distances==TableWrapDistances::default()}
    pub fn validate(&self)->crate::RtfResult<()>{let offset=|value|(0..=MAX_FLOATING_TABLE_DISTANCE_TWIPS).contains(&value);let negative=|value|(-MAX_FLOATING_TABLE_DISTANCE_TWIPS..=-1).contains(&value);if matches!(self.horizontal_position,Some(TableHorizontalPosition::Offset(value))if !offset(value))||matches!(self.vertical_position,Some(TableVerticalPosition::Offset(value))if !offset(value))||matches!(self.horizontal_position,Some(TableHorizontalPosition::NegativeOffset(value))if !negative(value))||matches!(self.vertical_position,Some(TableVerticalPosition::NegativeOffset(value))if !negative(value)){return Err(crate::RtfError::MalformedDocument("RTF floating-table position is out of range".to_string()))}self.wrap_distances.validate()}
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableDistanceUnit { Null, Twips }

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableDistanceScope { Row, Cell }
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableDistanceKind { Padding, Spacing }
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableEdge { Left, Right, Top, Bottom }
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TableDistanceTarget { pub scope: TableDistanceScope, pub kind: TableDistanceKind, pub edge: TableEdge }

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct TableSideDistance { pub value: Option<u16>, pub unit: Option<TableDistanceUnit> }

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct TableEdgeDistances { pub left: TableSideDistance, pub right: TableSideDistance, pub top: TableSideDistance, pub bottom: TableSideDistance }
impl TableEdgeDistances {
    pub(crate) fn side_mut(&mut self, edge: TableEdge) -> &mut TableSideDistance { match edge { TableEdge::Left=>&mut self.left,TableEdge::Right=>&mut self.right,TableEdge::Top=>&mut self.top,TableEdge::Bottom=>&mut self.bottom } }
    pub fn validate(&self) -> crate::RtfResult<()> { for side in [&self.left,&self.right,&self.top,&self.bottom] { if side.value.is_some_and(|value|i32::from(value)>MAX_TABLE_DISTANCE_TWIPS){return Err(crate::RtfError::MalformedDocument("RTF table distance is out of range".to_string()))} } Ok(()) }
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
    /// Explicit row direction.
    direction: Option<TextDirection>,
    padding: TableEdgeDistances,
    spacing: TableEdgeDistances,
    positioning: FloatingTablePosition,
}

impl<'a> Row<'a> {
    /// Create a new row.
    pub fn new() -> Self {
        Self {
            cells: Vec::new(),
            direction: None,
            padding: TableEdgeDistances::default(),
            spacing: TableEdgeDistances::default(),
            positioning: FloatingTablePosition::default(),
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

    /// Return the explicit row direction.
    pub fn direction(&self) -> Option<TextDirection> {
        self.direction
    }

    /// Set or clear the explicit row direction.
    pub fn set_direction(&mut self, direction: Option<TextDirection>) {
        self.direction = direction;
    }
    pub fn padding(&self)->&TableEdgeDistances{&self.padding}
    pub fn spacing(&self)->&TableEdgeDistances{&self.spacing}
    pub fn set_padding(&mut self,value:TableEdgeDistances){self.padding=value}
    pub fn set_spacing(&mut self,value:TableEdgeDistances){self.spacing=value}
    pub fn positioning(&self)->&FloatingTablePosition{&self.positioning}
    pub fn set_positioning(&mut self,value:FloatingTablePosition){self.positioning=value}
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
    nested_tables: Vec<CellNestedTable<'a>>,
}

/// A nested table inserted between UTF-8 text bytes in its containing cell.
#[derive(Debug, Clone)]
pub struct CellNestedTable<'a> { pub text_offset: usize, pub table: Table<'a> }

impl<'a> Cell<'a> {
    /// Create a new cell.
    pub fn new(text: Cow<'a, str>) -> Self {
        Self { text, padding:TableEdgeDistances::default(), spacing:TableEdgeDistances::default(), nested_tables:Vec::new() }
    }
    pub fn with_distances(text:Cow<'a,str>,padding:TableEdgeDistances,spacing:TableEdgeDistances)->Self{Self{text,padding,spacing,nested_tables:Vec::new()}}

    /// Get the cell text.
    pub fn text(&self) -> &str {
        &self.text
    }
    pub fn padding(&self)->&TableEdgeDistances{&self.padding}
    pub fn spacing(&self)->&TableEdgeDistances{&self.spacing}
    pub fn set_padding(&mut self,value:TableEdgeDistances){self.padding=value}
    pub fn set_spacing(&mut self,value:TableEdgeDistances){self.spacing=value}
    pub fn nested_tables(&self)->&[CellNestedTable<'a>]{&self.nested_tables}
    pub fn add_nested_table(&mut self,text_offset:usize,table:Table<'a>)->crate::RtfResult<()>{if text_offset>self.text.len()||!self.text.is_char_boundary(text_offset)||self.nested_tables.last().is_some_and(|entry|entry.text_offset>text_offset){return Err(crate::RtfError::MalformedDocument("invalid nested-table text insertion offset".to_string()))}self.nested_tables.push(CellNestedTable{text_offset,table});Ok(())}
    pub(crate) fn nested_tables_mut(&mut self)->&mut Vec<CellNestedTable<'a>>{&mut self.nested_tables}
}

impl<'a> Row<'a>{pub(crate) fn cells_mut(&mut self)->&mut [Cell<'a>]{&mut self.cells}}
