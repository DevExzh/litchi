//! Typed worksheet-view models for `SpreadsheetML`.

use crate::error::Result;
use litchi_sheet::Rect;
use litchi_sheet::view::{Position, View};

use super::invalid;

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
pub struct Extension {
    pub(super) uri: String,
    pub(super) markup: Vec<u8>,
}
impl Extension {
    #[must_use]
    pub fn uri(&self) -> &str {
        &self.uri
    }
    /// MCE-processed extension markup. It is retained, not interpreted or executed.
    #[must_use]
    pub fn markup(&self) -> &[u8] {
        &self.markup
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
    pub(super) offset: Option<Rect>,
    pub(super) collapsed_levels_are_subtotals: bool,
    pub(super) axis: Option<PivotSelectionAxis>,
    pub(super) field_position: Option<u32>,
    pub(super) markup: Vec<u8>,
}
impl PivotArea {
    #[must_use]
    pub fn field(&self) -> Option<i32> {
        self.field
    }
    #[must_use]
    pub fn area_type(&self) -> PivotAreaType {
        self.area_type
    }
    #[must_use]
    pub fn data_only(&self) -> bool {
        self.data_only
    }
    #[must_use]
    pub fn label_only(&self) -> bool {
        self.label_only
    }
    #[must_use]
    pub fn grand_row(&self) -> bool {
        self.grand_row
    }
    #[must_use]
    pub fn grand_column(&self) -> bool {
        self.grand_column
    }
    #[must_use]
    pub fn cache_index(&self) -> bool {
        self.cache_index
    }
    #[must_use]
    pub fn outline(&self) -> bool {
        self.outline
    }
    #[must_use]
    pub fn offset(&self) -> Option<Rect> {
        self.offset
    }
    #[must_use]
    pub fn collapsed_levels_are_subtotals(&self) -> bool {
        self.collapsed_levels_are_subtotals
    }
    #[must_use]
    pub fn axis(&self) -> Option<PivotSelectionAxis> {
        self.axis
    }
    #[must_use]
    pub fn field_position(&self) -> Option<u32> {
        self.field_position
    }
    /// Complete, bounded, MCE-processed `pivotArea` markup, including references and extensions.
    #[must_use]
    pub fn markup(&self) -> &[u8] {
        &self.markup
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotSelection {
    pub(super) pane: Position,
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
    #[must_use]
    pub fn pane(&self) -> Position {
        self.pane
    }
    #[must_use]
    pub fn show_header(&self) -> bool {
        self.show_header
    }
    #[must_use]
    pub fn label(&self) -> bool {
        self.label
    }
    #[must_use]
    pub fn data(&self) -> bool {
        self.data
    }
    #[must_use]
    pub fn extendable(&self) -> bool {
        self.extendable
    }
    #[must_use]
    pub fn count(&self) -> u32 {
        self.count
    }
    #[must_use]
    pub fn axis(&self) -> Option<PivotSelectionAxis> {
        self.axis
    }
    #[must_use]
    pub fn dimension(&self) -> u32 {
        self.dimension
    }
    #[must_use]
    pub fn start(&self) -> u32 {
        self.start
    }
    #[must_use]
    pub fn min(&self) -> u32 {
        self.min
    }
    #[must_use]
    pub fn max(&self) -> u32 {
        self.max
    }
    #[must_use]
    pub fn active_row(&self) -> u32 {
        self.active_row
    }
    #[must_use]
    pub fn active_column(&self) -> u32 {
        self.active_column
    }
    #[must_use]
    pub fn previous_row(&self) -> u32 {
        self.previous_row
    }
    #[must_use]
    pub fn previous_column(&self) -> u32 {
        self.previous_column
    }
    #[must_use]
    pub fn click(&self) -> u32 {
        self.click
    }
    #[must_use]
    pub fn relationship_id(&self) -> Option<&str> {
        self.relationship_id.as_deref()
    }
    #[must_use]
    pub fn area(&self) -> &PivotArea {
        &self.area
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Entry {
    pub(super) view: View,
    pub(super) pivot_selections: Vec<PivotSelection>,
    pub(super) extensions: Vec<Extension>,
    pub(super) retained_xml: Vec<u8>,
}
impl Entry {
    pub fn view(&self) -> &View {
        &self.view
    }
    #[must_use]
    pub fn pivot_selections(&self) -> &[PivotSelection] {
        &self.pivot_selections
    }
    #[must_use]
    pub fn extensions(&self) -> &[Extension] {
        &self.extensions
    }
    /// Complete MCE-processed `sheetView` markup retained for source fidelity.
    #[must_use]
    pub fn retained_xml(&self) -> &[u8] {
        &self.retained_xml
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Collection {
    pub(super) entries: Vec<Entry>,
    pub(super) extensions: Vec<Extension>,
}
impl Collection {
    #[must_use]
    pub fn entries(&self) -> &[Entry] {
        &self.entries
    }
    #[must_use]
    pub fn extensions(&self) -> &[Extension] {
        &self.extensions
    }
}
