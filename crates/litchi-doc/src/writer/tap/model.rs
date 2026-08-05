//! Semantic DOC table authoring models.
//!
//! These types describe table rows, cells, styles, and revision snapshots in
//! DOC-native terms. Encoding is deliberately kept in the sibling codec
//! module, while invalid states use the shared validation vocabulary.

use crate::parts::tap::{
    BorderStyle, CellBorderTypes, CellBorders, CellShading, CellSpacing, TableHorizontalPosition,
    TableJustification, TableLook, TablePositioning, TableVerticalPosition, TableWidth,
    TextDirection, VerticalAlignment, VerticalMergeStatus,
};

use super::validation::TapBuildError;

/// Table cell descriptor
#[derive(Debug, Clone, Default)]
pub struct TableCell {
    /// Cell width (in twips)
    pub width: u16,
    /// This cell is merged into the preceding cell. The preceding cell is
    /// automatically encoded as the start of the horizontal merge.
    pub merged: bool,
    /// Vertical merge state for this cell
    pub vertical_merge: VerticalMergeStatus,
    /// Vertical alignment of cell contents
    pub vertical_alignment: VerticalAlignment,
    /// Cell text flow and rotation
    pub text_direction: TextDirection,
    /// Stretch contents to use the full cell width
    pub fit_text: bool,
    /// Prefer cell contents on a single unwrapped line
    pub no_wrap: bool,
    /// Hide the cell mark when every cell in the row is empty
    pub hide_mark: bool,
    /// Cell edge borders
    pub borders: CellBorders,
    /// Border-type-only overrides; `None` inherits that side's type
    pub border_type_overrides: CellBorderTypes,
    /// Complete legacy cell shading
    pub shading: Option<CellShading>,
    /// Cell padding in twips
    pub padding_top: Option<u16>,
    pub padding_left: Option<u16>,
    pub padding_bottom: Option<u16>,
    pub padding_right: Option<u16>,
}

/// Default borders for a DOC table row.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct TableBorders {
    pub top: Option<BorderStyle>,
    pub left: Option<BorderStyle>,
    pub bottom: Option<BorderStyle>,
    pub right: Option<BorderStyle>,
    pub horizontal: Option<BorderStyle>,
    pub vertical: Option<BorderStyle>,
}

/// Raw property revision metadata for a DOC table row.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TableRevisionMark {
    /// Whether this operand represents an active property revision.
    pub active: bool,
    /// Index into the document's `SttbfRMark` author table.
    pub author_index: u16,
    /// Packed MS-DOC `DTTM` value.
    pub timestamp: u32,
}

/// Table row properties
#[derive(Debug, Clone)]
pub struct TableRow {
    /// Cells in this row
    pub cells: Vec<TableCell>,
    /// Row height in twips (positive = at least, negative = exact, zero = auto)
    pub height: i16,
    /// Header row flag
    pub is_header: bool,
    /// Whether the row may split across page breaks
    pub allow_break: bool,
    /// Logical table justification
    pub justification: TableJustification,
    /// Preferred total table width
    pub preferred_width: Option<TableWidth>,
    /// Automatically resize columns to fit table contents
    pub auto_fit: bool,
    /// Preferred leading width before the first cell
    pub width_before: Option<TableWidth>,
    /// Preferred trailing width after the last cell
    pub width_after: Option<TableWidth>,
    /// Preferred leading indentation of the table
    pub preferred_indent: Option<TableWidth>,
    /// Avoid a page break between this row and the next row
    pub keep_with_next: bool,
    /// Table auto-format identity and optional look flags
    pub table_look: Option<TableLook>,
    /// Style-sheet index of the applied table style
    pub table_style_index: Option<u16>,
    /// Lay out the table from right to left
    pub right_to_left: bool,
    /// Whether this floating table may overlap other tables
    pub allow_overlap: bool,
    /// Anchor origins when this table is absolutely positioned
    pub positioning: Option<TablePositioning>,
    /// Horizontal alignment or physical offset from the anchor
    pub horizontal_position: TableHorizontalPosition,
    /// Vertical alignment or downward offset from the anchor
    pub vertical_position: TableVerticalPosition,
    /// Minimum text-wrapping distances on the physical sides, in twips
    pub distance_from_text_left: u16,
    pub distance_from_text_top: u16,
    pub distance_from_text_right: u16,
    pub distance_from_text_bottom: u16,
    /// Uniform spacing around every cell in this row
    pub cell_spacing: Option<CellSpacing>,
    /// Nonzero `PGPInfo.ipgpSelf` associated with this row
    pub paragraph_group_id: Option<u32>,
    /// Revision save ID associated with this table formatting
    pub revision_save_id: Option<u32>,
    /// Tracked row-property revision metadata
    pub formatting_revision: Option<TableRevisionMark>,
    /// Preserve pre-revision properties before the `sprmTWall` boundary
    pub properties_preserved_for_revision: bool,
    /// Full row state retained before the `sprmTWall` boundary
    pub preserved_properties_for_revision: Option<Box<TableRow>>,
    /// Default outer and inside borders for this row
    pub borders: TableBorders,
}

impl Default for TableRow {
    fn default() -> Self {
        Self {
            cells: Vec::new(),
            height: 0,
            is_header: false,
            allow_break: true,
            justification: TableJustification::Left,
            preferred_width: None,
            auto_fit: false,
            width_before: None,
            width_after: None,
            preferred_indent: None,
            keep_with_next: false,
            table_look: None,
            table_style_index: None,
            right_to_left: false,
            allow_overlap: true,
            positioning: None,
            horizontal_position: TableHorizontalPosition::Left,
            vertical_position: TableVerticalPosition::Inline,
            distance_from_text_left: 0,
            distance_from_text_top: 0,
            distance_from_text_right: 0,
            distance_from_text_bottom: 0,
            cell_spacing: None,
            paragraph_group_id: None,
            revision_save_id: None,
            formatting_revision: None,
            properties_preserved_for_revision: false,
            preserved_properties_for_revision: None,
            borders: TableBorders::default(),
        }
    }
}

/// TAP (Table Properties) builder
#[derive(Debug)]
pub struct TapBuilder {
    rows: Vec<TableRow>,
}

impl TapBuilder {
    /// Create a new TAP builder
    pub fn new() -> Self {
        Self { rows: Vec::new() }
    }

    /// Add a row to the table
    pub fn add_row(&mut self, row: TableRow) {
        self.rows.push(row);
    }

    /// Generate TAP SPRMs for a specific row
    pub fn generate_row_sprms(&self, row_index: usize) -> Vec<u8> {
        self.try_generate_row_sprms(row_index).unwrap_or_default()
    }

    /// Generate validated TAP SPRMs for a specific row.
    pub fn try_generate_row_sprms(&self, row_index: usize) -> Result<Vec<u8>, TapBuildError> {
        let row = self
            .rows
            .get(row_index)
            .ok_or(TapBuildError::RowOutOfBounds(row_index))?;
        super::codec::generate_row_sprms(row)
    }

    /// Borrow the configured rows.
    pub fn rows(&self) -> &[TableRow] {
        &self.rows
    }

    /// Get the number of rows
    pub fn row_count(&self) -> usize {
        self.rows.len()
    }
}

impl Default for TapBuilder {
    fn default() -> Self {
        Self::new()
    }
}

/// Helper to create a simple table
pub fn create_simple_table(rows: usize, cols: usize, cell_width: u16) -> TapBuilder {
    let mut builder = TapBuilder::new();

    for _ in 0..rows {
        let cells = vec![
            TableCell {
                width: cell_width,
                merged: false,
                ..TableCell::default()
            };
            cols
        ];
        builder.add_row(TableRow {
            cells,
            ..TableRow::default()
        });
    }

    builder
}
