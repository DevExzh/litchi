//! Typed worksheet-view models for SpreadsheetML.

use crate::error::Result;

use super::invalid;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ViewType {
    Normal,
    PageBreakPreview,
    PageLayout,
}
impl ViewType {
    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "normal" => Ok(Self::Normal),
            "pageBreakPreview" => Ok(Self::PageBreakPreview),
            "pageLayout" => Ok(Self::PageLayout),
            _ => Err(invalid(format!("invalid worksheet-view type '{value}'"))),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PanePosition {
    BottomRight,
    TopRight,
    BottomLeft,
    TopLeft,
}
impl PanePosition {
    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "bottomRight" => Ok(Self::BottomRight),
            "topRight" => Ok(Self::TopRight),
            "bottomLeft" => Ok(Self::BottomLeft),
            "topLeft" => Ok(Self::TopLeft),
            _ => Err(invalid(format!("invalid worksheet-view pane '{value}'"))),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PaneState {
    Split,
    Frozen,
    FrozenSplit,
}
impl PaneState {
    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "split" => Ok(Self::Split),
            "frozen" => Ok(Self::Frozen),
            "frozenSplit" => Ok(Self::FrozenSplit),
            _ => Err(invalid(format!(
                "invalid worksheet-view pane state '{value}'"
            ))),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PivotSelectionAxis {
    Row,
    Column,
    Page,
    Values,
}
impl PivotSelectionAxis {
    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "axisRow" => Ok(Self::Row),
            "axisCol" => Ok(Self::Column),
            "axisPage" => Ok(Self::Page),
            "axisValues" => Ok(Self::Values),
            _ => Err(invalid(format!("invalid pivot-selection axis '{value}'"))),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PivotAreaType {
    None,
    Normal,
    Data,
    All,
    Origin,
    Button,
    TopRight,
    TopEnd,
}
impl PivotAreaType {
    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "none" => Ok(Self::None),
            "normal" => Ok(Self::Normal),
            "data" => Ok(Self::Data),
            "all" => Ok(Self::All),
            "origin" => Ok(Self::Origin),
            "button" => Ok(Self::Button),
            "topRight" => Ok(Self::TopRight),
            "topEnd" => Ok(Self::TopEnd),
            _ => Err(invalid(format!("invalid pivot-area type '{value}'"))),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CellReference(pub(super) String);
impl CellReference {
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RangeReference(pub(super) String);
impl RangeReference {
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Sqref(pub(super) Vec<RangeReference>);
impl Sqref {
    pub fn ranges(&self) -> &[RangeReference] {
        &self.0
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Extension {
    pub(super) uri: String,
    pub(super) markup: Vec<u8>,
}
impl Extension {
    pub fn uri(&self) -> &str {
        &self.uri
    }
    /// MCE-processed extension markup. It is retained, not interpreted or executed.
    pub fn markup(&self) -> &[u8] {
        &self.markup
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Pane {
    pub(super) x_split: Option<f64>,
    pub(super) y_split: Option<f64>,
    pub(super) top_left_cell: Option<CellReference>,
    pub(super) active_pane: PanePosition,
    pub(super) state: PaneState,
}
impl Pane {
    pub fn x_split(&self) -> Option<f64> {
        self.x_split
    }
    pub fn y_split(&self) -> Option<f64> {
        self.y_split
    }
    pub fn top_left_cell(&self) -> Option<&CellReference> {
        self.top_left_cell.as_ref()
    }
    pub fn active_pane(&self) -> PanePosition {
        self.active_pane
    }
    pub fn state(&self) -> PaneState {
        self.state
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Selection {
    pub(super) pane: PanePosition,
    pub(super) active_cell: CellReference,
    pub(super) active_cell_id: u32,
    pub(super) sqref: Sqref,
}
impl Selection {
    pub fn pane(&self) -> PanePosition {
        self.pane
    }
    pub fn active_cell(&self) -> &CellReference {
        &self.active_cell
    }
    pub fn active_cell_id(&self) -> u32 {
        self.active_cell_id
    }
    pub fn sqref(&self) -> &Sqref {
        &self.sqref
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotArea {
    pub(super) field: Option<i32>,
    pub(super) area_type: PivotAreaType,
    pub(super) data_only: bool,
    pub(super) label_only: bool,
    pub(super) grand_row: bool,
    pub(super) grand_column: bool,
    pub(super) cache_index: bool,
    pub(super) outline: bool,
    pub(super) offset: Option<RangeReference>,
    pub(super) collapsed_levels_are_subtotals: bool,
    pub(super) axis: Option<PivotSelectionAxis>,
    pub(super) field_position: Option<u32>,
    pub(super) markup: Vec<u8>,
}
impl PivotArea {
    pub fn field(&self) -> Option<i32> {
        self.field
    }
    pub fn area_type(&self) -> PivotAreaType {
        self.area_type
    }
    pub fn data_only(&self) -> bool {
        self.data_only
    }
    pub fn label_only(&self) -> bool {
        self.label_only
    }
    pub fn grand_row(&self) -> bool {
        self.grand_row
    }
    pub fn grand_column(&self) -> bool {
        self.grand_column
    }
    pub fn cache_index(&self) -> bool {
        self.cache_index
    }
    pub fn outline(&self) -> bool {
        self.outline
    }
    pub fn offset(&self) -> Option<&RangeReference> {
        self.offset.as_ref()
    }
    pub fn collapsed_levels_are_subtotals(&self) -> bool {
        self.collapsed_levels_are_subtotals
    }
    pub fn axis(&self) -> Option<PivotSelectionAxis> {
        self.axis
    }
    pub fn field_position(&self) -> Option<u32> {
        self.field_position
    }
    /// Complete, bounded, MCE-processed `pivotArea` markup, including references and extensions.
    pub fn markup(&self) -> &[u8] {
        &self.markup
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotSelection {
    pub(super) pane: PanePosition,
    pub(super) show_header: bool,
    pub(super) label: bool,
    pub(super) data: bool,
    pub(super) extendable: bool,
    pub(super) count: u32,
    pub(super) axis: Option<PivotSelectionAxis>,
    pub(super) dimension: u32,
    pub(super) start: u32,
    pub(super) min: u32,
    pub(super) max: u32,
    pub(super) active_row: u32,
    pub(super) active_column: u32,
    pub(super) previous_row: u32,
    pub(super) previous_column: u32,
    pub(super) click: u32,
    pub(super) relationship_id: Option<String>,
    pub(super) area: PivotArea,
}
impl PivotSelection {
    pub fn pane(&self) -> PanePosition {
        self.pane
    }
    pub fn show_header(&self) -> bool {
        self.show_header
    }
    pub fn label(&self) -> bool {
        self.label
    }
    pub fn data(&self) -> bool {
        self.data
    }
    pub fn extendable(&self) -> bool {
        self.extendable
    }
    pub fn count(&self) -> u32 {
        self.count
    }
    pub fn axis(&self) -> Option<PivotSelectionAxis> {
        self.axis
    }
    pub fn dimension(&self) -> u32 {
        self.dimension
    }
    pub fn start(&self) -> u32 {
        self.start
    }
    pub fn min(&self) -> u32 {
        self.min
    }
    pub fn max(&self) -> u32 {
        self.max
    }
    pub fn active_row(&self) -> u32 {
        self.active_row
    }
    pub fn active_column(&self) -> u32 {
        self.active_column
    }
    pub fn previous_row(&self) -> u32 {
        self.previous_row
    }
    pub fn previous_column(&self) -> u32 {
        self.previous_column
    }
    pub fn click(&self) -> u32 {
        self.click
    }
    pub fn relationship_id(&self) -> Option<&str> {
        self.relationship_id.as_deref()
    }
    pub fn area(&self) -> &PivotArea {
        &self.area
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct View {
    pub(super) workbook_view_id: u32,
    pub(super) window_protection: bool,
    pub(super) show_formulas: bool,
    pub(super) show_grid_lines: bool,
    pub(super) show_row_col_headers: bool,
    pub(super) show_zeros: bool,
    pub(super) right_to_left: bool,
    pub(super) tab_selected: bool,
    pub(super) show_ruler: bool,
    pub(super) show_outline_symbols: bool,
    pub(super) default_grid_color: bool,
    pub(super) show_white_space: bool,
    pub(super) view_type: ViewType,
    pub(super) top_left_cell: Option<CellReference>,
    pub(super) color_id: u32,
    pub(super) zoom_scale: u16,
    pub(super) zoom_scale_normal: u16,
    pub(super) zoom_scale_sheet_layout_view: u16,
    pub(super) zoom_scale_page_layout_view: u16,
    pub(super) pane: Option<Pane>,
    pub(super) selections: Vec<Selection>,
    pub(super) pivot_selections: Vec<PivotSelection>,
    pub(super) extensions: Vec<Extension>,
}
impl View {
    pub fn workbook_view_id(&self) -> u32 {
        self.workbook_view_id
    }
    pub fn window_protection(&self) -> bool {
        self.window_protection
    }
    pub fn show_formulas(&self) -> bool {
        self.show_formulas
    }
    pub fn show_grid_lines(&self) -> bool {
        self.show_grid_lines
    }
    pub fn show_row_col_headers(&self) -> bool {
        self.show_row_col_headers
    }
    pub fn show_zeros(&self) -> bool {
        self.show_zeros
    }
    pub fn right_to_left(&self) -> bool {
        self.right_to_left
    }
    pub fn tab_selected(&self) -> bool {
        self.tab_selected
    }
    pub fn show_ruler(&self) -> bool {
        self.show_ruler
    }
    pub fn show_outline_symbols(&self) -> bool {
        self.show_outline_symbols
    }
    pub fn default_grid_color(&self) -> bool {
        self.default_grid_color
    }
    pub fn show_white_space(&self) -> bool {
        self.show_white_space
    }
    pub fn view_type(&self) -> ViewType {
        self.view_type
    }
    pub fn top_left_cell(&self) -> Option<&CellReference> {
        self.top_left_cell.as_ref()
    }
    pub fn color_id(&self) -> u32 {
        self.color_id
    }
    pub fn zoom_scale(&self) -> u16 {
        self.zoom_scale
    }
    pub fn zoom_scale_normal(&self) -> u16 {
        self.zoom_scale_normal
    }
    pub fn zoom_scale_sheet_layout_view(&self) -> u16 {
        self.zoom_scale_sheet_layout_view
    }
    pub fn zoom_scale_page_layout_view(&self) -> u16 {
        self.zoom_scale_page_layout_view
    }
    pub fn pane(&self) -> Option<&Pane> {
        self.pane.as_ref()
    }
    pub fn selections(&self) -> &[Selection] {
        &self.selections
    }
    pub fn pivot_selections(&self) -> &[PivotSelection] {
        &self.pivot_selections
    }
    pub fn extensions(&self) -> &[Extension] {
        &self.extensions
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Views {
    pub(super) views: Vec<View>,
    pub(super) extensions: Vec<Extension>,
}
impl Views {
    pub fn views(&self) -> &[View] {
        &self.views
    }
    pub fn extensions(&self) -> &[Extension] {
        &self.extensions
    }
}
