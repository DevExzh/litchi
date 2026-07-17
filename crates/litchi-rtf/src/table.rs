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

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum TablePreferredWidthUnit { #[default] Null, Auto, Percent, Twips }

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TablePreferredWidth { unit:TablePreferredWidthUnit, value:Option<u16> }
impl TablePreferredWidth {
    pub fn new(unit:TablePreferredWidthUnit,value:Option<u16>)->crate::RtfResult<Self>{let width=Self{unit,value};width.validate()?;Ok(width)}
    pub const fn unit(self)->TablePreferredWidthUnit{self.unit}
    pub const fn value(self)->Option<u16>{self.value}
    pub(crate) fn validate(self)->crate::RtfResult<()>{match(self.unit,self.value){(TablePreferredWidthUnit::Null|TablePreferredWidthUnit::Auto,None)=>Ok(()),(TablePreferredWidthUnit::Percent,Some(value))if i32::from(value)<=MAX_TABLE_WIDTH_PERCENT=>Ok(()),(TablePreferredWidthUnit::Twips,Some(value))if i32::from(value)<=MAX_TABLE_GEOMETRY_TWIPS=>Ok(()),(TablePreferredWidthUnit::Null|TablePreferredWidthUnit::Auto,Some(_))=>Err(crate::RtfError::MalformedDocument("RTF null or auto preferred width cannot carry a value".to_string())),(TablePreferredWidthUnit::Percent|TablePreferredWidthUnit::Twips,None)=>Err(crate::RtfError::MalformedDocument("RTF percentage or twip preferred width requires a value".to_string())),_=>Err(crate::RtfError::MalformedDocument("RTF preferred table width is out of range".to_string()))}}
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum TableRowHeight { #[default] Automatic, Minimum(u16), Exact(u16) }

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableIndentUnit { Auto, Twips, Nil, Percent }

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TableIndent { unit:TableIndentUnit, value:i32 }
impl TableIndent {
    pub fn new(unit:TableIndentUnit,value:i32)->crate::RtfResult<Self>{let indent=Self{unit,value};indent.validate()?;Ok(indent)}
    pub const fn unit(self)->TableIndentUnit{self.unit}
    pub const fn value(self)->i32{self.value}
    pub(crate) fn validate(self)->crate::RtfResult<()>{let cap=match self.unit{TableIndentUnit::Twips=>MAX_TABLE_GEOMETRY_TWIPS,TableIndentUnit::Percent=>MAX_TABLE_WIDTH_PERCENT,TableIndentUnit::Auto|TableIndentUnit::Nil=>return if self.value==0{Ok(())}else{Err(crate::RtfError::MalformedDocument("RTF auto or nil table indent requires zero".to_string()))}};if self.value.unsigned_abs()<=cap as u32{Ok(())}else{Err(crate::RtfError::MalformedDocument("RTF table indent is out of range".to_string()))}}
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct TableRowGeometry { half_gap_twips:Option<u16>, left_edge_twips:Option<i32>, height:TableRowHeight, preferred_width:Option<TablePreferredWidth>, leading_invisible_width:Option<TablePreferredWidth>, trailing_invisible_width:Option<TablePreferredWidth>, auto_fit:bool, indent:Option<TableIndent> }
impl TableRowGeometry {
    pub const fn half_gap_twips(self)->Option<u16>{self.half_gap_twips} pub fn set_half_gap_twips(&mut self,value:Option<u16>){self.half_gap_twips=value}
    pub const fn left_edge_twips(self)->Option<i32>{self.left_edge_twips} pub fn set_left_edge_twips(&mut self,value:Option<i32>){self.left_edge_twips=value}
    pub const fn height(self)->TableRowHeight{self.height} pub fn set_height(&mut self,value:TableRowHeight){self.height=value}
    pub const fn preferred_width(self)->Option<TablePreferredWidth>{self.preferred_width} pub fn set_preferred_width(&mut self,value:Option<TablePreferredWidth>){self.preferred_width=value}
    /// Width of the invisible cell at the logical beginning of the row (`trftsWidthB`/`trwWidthB`).
    /// This is the left side for an LTR row and the right side for an RTL row.
    pub const fn leading_invisible_width(self)->Option<TablePreferredWidth>{self.leading_invisible_width} pub fn set_leading_invisible_width(&mut self,value:Option<TablePreferredWidth>){self.leading_invisible_width=value}
    /// Width of the invisible cell at the logical end of the row (`trftsWidthA`/`trwWidthA`).
    /// This is the right side for an LTR row and the left side for an RTL row.
    pub const fn trailing_invisible_width(self)->Option<TablePreferredWidth>{self.trailing_invisible_width} pub fn set_trailing_invisible_width(&mut self,value:Option<TablePreferredWidth>){self.trailing_invisible_width=value}
    pub const fn auto_fit(self)->bool{self.auto_fit} pub fn set_auto_fit(&mut self,value:bool){self.auto_fit=value}
    pub const fn indent(self)->Option<TableIndent>{self.indent} pub fn set_indent(&mut self,value:Option<TableIndent>){self.indent=value}
    pub(crate) fn validate(self)->crate::RtfResult<()>{if self.half_gap_twips.is_some_and(|value|i32::from(value)>MAX_TABLE_GEOMETRY_TWIPS)||self.left_edge_twips.is_some_and(|value|value.unsigned_abs()>MAX_TABLE_GEOMETRY_TWIPS as u32)||matches!(self.height,TableRowHeight::Minimum(value)|TableRowHeight::Exact(value)if i32::from(value)>MAX_TABLE_GEOMETRY_TWIPS){return Err(crate::RtfError::MalformedDocument("RTF table row geometry is out of range".to_string()))}for width in [self.preferred_width,self.leading_invisible_width,self.trailing_invisible_width].into_iter().flatten(){width.validate()?}if let Some(indent)=self.indent{indent.validate()?}Ok(())}
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableCellMergeAxis { Horizontal, Vertical }

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableCellMergeRole { First, Continuation }

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct TableCellMergeState { pub horizontal:Option<TableCellMergeRole>, pub vertical:Option<TableCellMergeRole> }

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
pub enum TableRowBorderSide { Top, Left, Bottom, Right, Horizontal, Vertical }

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableCellBorderSide { Top, Left, Bottom, Right, UpperLeftToLowerRight, UpperRightToLowerLeft }

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum TableBorderTarget { Row(TableRowBorderSide), Cell(TableCellBorderSide) }

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct TableRowBorders { pub top:Option<crate::Border>, pub left:Option<crate::Border>, pub bottom:Option<crate::Border>, pub right:Option<crate::Border>, pub horizontal:Option<crate::Border>, pub vertical:Option<crate::Border> }
impl TableRowBorders {
    pub(crate) fn side_mut(&mut self,side:TableRowBorderSide)->&mut Option<crate::Border>{match side{TableRowBorderSide::Top=>&mut self.top,TableRowBorderSide::Left=>&mut self.left,TableRowBorderSide::Bottom=>&mut self.bottom,TableRowBorderSide::Right=>&mut self.right,TableRowBorderSide::Horizontal=>&mut self.horizontal,TableRowBorderSide::Vertical=>&mut self.vertical}}
    pub fn validate(&self)->crate::RtfResult<()>{for border in [self.top,self.left,self.bottom,self.right,self.horizontal,self.vertical].into_iter().flatten(){border.validate_table()?;}Ok(())}
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct TableCellBorders { pub top:Option<crate::Border>, pub left:Option<crate::Border>, pub bottom:Option<crate::Border>, pub right:Option<crate::Border>, pub upper_left_to_lower_right:Option<crate::Border>, pub upper_right_to_lower_left:Option<crate::Border> }
impl TableCellBorders {
    pub(crate) fn side_mut(&mut self,side:TableCellBorderSide)->&mut Option<crate::Border>{match side{TableCellBorderSide::Top=>&mut self.top,TableCellBorderSide::Left=>&mut self.left,TableCellBorderSide::Bottom=>&mut self.bottom,TableCellBorderSide::Right=>&mut self.right,TableCellBorderSide::UpperLeftToLowerRight=>&mut self.upper_left_to_lower_right,TableCellBorderSide::UpperRightToLowerLeft=>&mut self.upper_right_to_lower_left}}
    pub fn validate(&self)->crate::RtfResult<()>{for border in [self.top,self.left,self.bottom,self.right,self.upper_left_to_lower_right,self.upper_right_to_lower_left].into_iter().flatten(){border.validate_table()?;}Ok(())}
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct TableShading { pub amount:Option<u16>, pub foreground_color:Option<crate::ColorRef>, pub background_color:Option<crate::ColorRef>, pub pattern:Option<crate::ShadingPattern>, pub pattern_index:Option<u16> }
impl TableShading { pub fn validate(&self)->crate::RtfResult<()>{if self.amount.is_some_and(|value|value>10_000){return Err(crate::RtfError::MalformedDocument("RTF table shading must be in 0..=10000".to_string()))}if self.pattern.is_some()&&self.pattern_index.is_some(){return Err(crate::RtfError::MalformedDocument("RTF table shading has conflicting pattern controls".to_string()))}Ok(())} }
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TableDistanceTarget { pub scope: TableDistanceScope, pub kind: TableDistanceKind, pub edge: TableEdge }

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableRowAlignment { Left, Center, Right }

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct TableRowLayout { pub header: bool, pub keep_together: bool, pub keep_with_following: bool, pub alignment: Option<TableRowAlignment> }

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableCellVerticalAlignment { Top, Center, Bottom }

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableCellTextFlow { LeftToRightTopToBottom, RightToLeftTopToBottom, LeftToRightBottomToTop, LeftToRightTopToBottomVertical, TopToBottomRightToLeftVertical }

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct TableCellLayout { pub vertical_alignment: Option<TableCellVerticalAlignment>, pub text_flow: Option<TableCellTextFlow>, pub fit_text: bool, pub no_wrap: bool, pub hide_mark: bool }

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
    layout: TableRowLayout,
    padding: TableEdgeDistances,
    spacing: TableEdgeDistances,
    positioning: FloatingTablePosition,
    borders: TableRowBorders,
    shading: TableShading,
    geometry: TableRowGeometry,
}

impl<'a> Row<'a> {
    /// Create a new row.
    pub fn new() -> Self {
        Self {
            cells: Vec::new(),
            direction: None,
            layout: TableRowLayout::default(),
            padding: TableEdgeDistances::default(),
            spacing: TableEdgeDistances::default(),
            positioning: FloatingTablePosition::default(),
            borders: Default::default(), shading: Default::default(), geometry:Default::default(),
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
    pub fn layout(&self)->&TableRowLayout{&self.layout}
    pub fn set_layout(&mut self,value:TableRowLayout){self.layout=value}
    pub fn padding(&self)->&TableEdgeDistances{&self.padding}
    pub fn spacing(&self)->&TableEdgeDistances{&self.spacing}
    pub fn set_padding(&mut self,value:TableEdgeDistances){self.padding=value}
    pub fn set_spacing(&mut self,value:TableEdgeDistances){self.spacing=value}
    pub fn positioning(&self)->&FloatingTablePosition{&self.positioning}
    pub fn set_positioning(&mut self,value:FloatingTablePosition){self.positioning=value}
    pub fn borders(&self)->&TableRowBorders{&self.borders}
    pub fn shading(&self)->TableShading{self.shading}
    pub fn set_borders(&mut self,value:TableRowBorders){self.borders=value}
    pub fn set_shading(&mut self,value:TableShading){self.shading=value}
    pub fn geometry(&self)->TableRowGeometry{self.geometry}
    pub fn set_geometry(&mut self,value:TableRowGeometry){self.geometry=value}
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
}

/// A nested table inserted between UTF-8 text bytes in its containing cell.
#[derive(Debug, Clone)]
pub struct CellNestedTable<'a> { pub text_offset: usize, pub table: Table<'a> }

impl<'a> Cell<'a> {
    /// Create a new cell.
    pub fn new(text: Cow<'a, str>) -> Self {
        Self { text, padding:TableEdgeDistances::default(), spacing:TableEdgeDistances::default(), layout:TableCellLayout::default(), borders:Default::default(), shading:Default::default(), merge:Default::default(), right_boundary:None, preferred_width:None, nested_tables:Vec::new() }
    }
    pub fn with_distances(text:Cow<'a,str>,padding:TableEdgeDistances,spacing:TableEdgeDistances)->Self{Self{text,padding,spacing,layout:TableCellLayout::default(),borders:Default::default(),shading:Default::default(),merge:Default::default(),right_boundary:None,preferred_width:None,nested_tables:Vec::new()}}

    /// Get the cell text.
    pub fn text(&self) -> &str {
        &self.text
    }
    pub fn padding(&self)->&TableEdgeDistances{&self.padding}
    pub fn spacing(&self)->&TableEdgeDistances{&self.spacing}
    pub fn set_padding(&mut self,value:TableEdgeDistances){self.padding=value}
    pub fn set_spacing(&mut self,value:TableEdgeDistances){self.spacing=value}
    pub fn layout(&self)->&TableCellLayout{&self.layout}
    pub fn set_layout(&mut self,value:TableCellLayout){self.layout=value}
    pub fn borders(&self)->&TableCellBorders{&self.borders}
    pub fn shading(&self)->TableShading{self.shading}
    pub fn set_borders(&mut self,value:TableCellBorders){self.borders=value}
    pub fn set_shading(&mut self,value:TableShading){self.shading=value}
    pub fn merge(&self)->TableCellMergeState{self.merge}
    pub fn set_merge(&mut self,value:TableCellMergeState){self.merge=value}
    pub fn right_boundary(&self)->Option<i32>{self.right_boundary}
    pub fn set_right_boundary(&mut self,value:Option<i32>){self.right_boundary=value}
    pub fn preferred_width(&self)->Option<TablePreferredWidth>{self.preferred_width}
    pub fn set_preferred_width(&mut self,value:Option<TablePreferredWidth>){self.preferred_width=value}
    pub fn nested_tables(&self)->&[CellNestedTable<'a>]{&self.nested_tables}
    pub fn add_nested_table(&mut self,text_offset:usize,table:Table<'a>)->crate::RtfResult<()>{if text_offset>self.text.len()||!self.text.is_char_boundary(text_offset)||self.nested_tables.last().is_some_and(|entry|entry.text_offset>text_offset){return Err(crate::RtfError::MalformedDocument("invalid nested-table text insertion offset".to_string()))}self.nested_tables.push(CellNestedTable{text_offset,table});Ok(())}
    pub(crate) fn nested_tables_mut(&mut self)->&mut Vec<CellNestedTable<'a>>{&mut self.nested_tables}
}

impl<'a> Row<'a>{pub(crate) fn cells_mut(&mut self)->&mut [Cell<'a>]{&mut self.cells}}

impl Table<'_>{
    pub(crate) fn validate_merges(&self)->Result<(),String>{
        let mut active_vertical:BTreeSet<(i32,i32)>=BTreeSet::new();
        for row in self.rows(){
            let mut horizontal_open=false;
            let mut next_vertical:BTreeSet<(i32,i32)>=BTreeSet::new();
            let mut left=0i32;
            for(index,cell)in row.cells().iter().enumerate(){
                let right=cell.right_boundary().unwrap_or(2880*((index+1)as i32));
                match cell.merge().horizontal{
                    None=>horizontal_open=false,
                    Some(TableCellMergeRole::First)=>horizontal_open=true,
                    Some(TableCellMergeRole::Continuation)=>{if !horizontal_open{return Err("RTF horizontal merge continuation does not follow a first cell".to_string())}},
                }
                let span=(left,right);
                match cell.merge().vertical{
                    None=>{},
                    Some(TableCellMergeRole::First)=>{next_vertical.insert(span);},
                    Some(TableCellMergeRole::Continuation)=>{if !active_vertical.contains(&span){return Err("RTF vertical merge continuation does not match a cell boundary in the preceding row".to_string())}next_vertical.insert(span);},
                }
                left=right;
            }
            active_vertical=next_vertical;
        }
        Ok(())
    }
}
