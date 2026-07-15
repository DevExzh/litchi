/// Table Properties (TAP) parser for DOC files.
///
/// TAP structures define table-level formatting including:
/// - Table borders and shading
/// - Row and cell definitions
/// - Table positioning
/// - Cell margins and spacing
use super::super::package::Result;

/// Table Properties structure.
///
/// Contains formatting and structural information for a table.
/// Based on Apache POI's TableProperties class.
#[derive(Debug, Clone)]
pub struct TableProperties {
    /// Number of cells in the row
    pub cell_count: usize,
    /// Cell boundaries (positions in twips from left margin)
    pub cell_boundaries: Vec<i16>,
    /// Cell properties for each cell
    pub cell_properties: Vec<CellProperties>,
    /// Table justification (alignment)
    pub justification: TableJustification,
    /// Half the width of spacing between cells (dxaGapHalf)
    pub gap_half: i16,
    /// Table indent from left margin (twips)
    pub indent_left: i16,
    /// Preferred table width
    pub preferred_width: Option<TableWidth>,
    /// Row height in twips (positive = at least, negative = exact)
    pub row_height: Option<i16>,
    /// Row is header row
    pub is_header_row: bool,
    /// Allow row to break across pages
    pub allow_row_break: bool,
    /// Whether the row has a tracked property change
    pub has_formatting_revision: Option<bool>,
    /// Row revision author index in `SttbfRMark`
    pub formatting_revision_author_index: Option<u16>,
    /// Packed row revision DTTM
    pub formatting_revision_timestamp: Option<u32>,
    /// Whether pre-revision table properties are preserved
    pub properties_preserved_for_revision: bool,
    /// Table borders
    pub border_top: Option<BorderStyle>,
    pub border_left: Option<BorderStyle>,
    pub border_bottom: Option<BorderStyle>,
    pub border_right: Option<BorderStyle>,
    pub border_horizontal: Option<BorderStyle>,
    pub border_vertical: Option<BorderStyle>,
}

/// Cell Properties structure.
///
/// Contains formatting for an individual table cell.
/// Based on Apache POI's TableCellDescriptor class.
#[derive(Debug, Clone, Default)]
pub struct CellProperties {
    /// Horizontal merge status
    pub merge_status: CellMergeStatus,
    /// Vertical merge status
    pub vertical_merge_status: VerticalMergeStatus,
    /// Vertical alignment
    pub vertical_alignment: VerticalAlignment,
    /// Cell background color (RGB)
    pub background_color: Option<(u8, u8, u8)>,
    /// Cell borders
    pub borders: CellBorders,
    /// Text direction
    pub text_direction: TextDirection,
    /// Stretch contents to use the full cell width
    pub fit_text: bool,
    /// Prefer cell contents on a single unwrapped line
    pub no_wrap: bool,
    /// Hide the cell mark when every cell in the row is empty
    pub hide_mark: bool,
    /// Preferred cell width
    pub preferred_width: Option<TableWidth>,
    /// Cell padding (in twips)
    pub padding_top: Option<i16>,
    pub padding_left: Option<i16>,
    pub padding_bottom: Option<i16>,
    pub padding_right: Option<i16>,
}

/// Cell merge status.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum CellMergeStatus {
    /// Not horizontally merged
    #[default]
    None,
    /// First cell in a horizontal merge
    First,
    /// Continuation of a horizontal merge
    Merged,
}

/// Vertical cell merge status.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum VerticalMergeStatus {
    /// Not vertically merged
    #[default]
    None,
    /// First cell in a vertical merge
    First,
    /// Continuation of a vertical merge
    Merged,
}

/// Vertical alignment within a cell.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum VerticalAlignment {
    #[default]
    Top,
    Center,
    Bottom,
}

/// Table justification (alignment).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum TableJustification {
    #[default]
    Left,
    Center,
    Right,
}

/// Text direction in a cell.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum TextDirection {
    /// Left to right, top to bottom
    #[default]
    LrTb,
    /// Top to bottom, right to left (vertical)
    TbRl,
    /// Bottom to top, left to right (vertical)
    BtLr,
    /// Left to right, bottom to top
    LrBt,
    /// Top to bottom, left to right (vertical)
    TbLr,
}

/// Table or cell width specification.
#[derive(Debug, Clone, Copy)]
pub struct TableWidth {
    /// Width value
    pub value: i16,
    /// Width type
    pub width_type: WidthType,
}

/// Width type for tables and cells.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum WidthType {
    /// Width in twips (1/1440 inch)
    Twips,
    /// Width as percentage (value * 50)
    Percentage,
    /// Auto width
    Auto,
}

/// Cell borders.
#[derive(Debug, Clone, Default)]
pub struct CellBorders {
    pub top: Option<BorderStyle>,
    pub left: Option<BorderStyle>,
    pub bottom: Option<BorderStyle>,
    pub right: Option<BorderStyle>,
}

/// Border style.
#[derive(Debug, Clone, Copy)]
pub struct BorderStyle {
    /// Line width in 1/8 points
    pub width: u8,
    /// Border color
    pub color: Option<(u8, u8, u8)>,
    /// Border type
    pub border_type: BorderType,
}

/// Border types.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BorderType {
    None,
    Single,
    Thick,
    Double,
    Dotted,
    Dashed,
    DotDash,
    DotDotDash,
    Triple,
    ThinThickSmall,
    ThickThinSmall,
    ThinThickThinSmall,
}

impl Default for TableProperties {
    fn default() -> Self {
        Self {
            cell_count: 0,
            cell_boundaries: Vec::new(),
            cell_properties: Vec::new(),
            justification: TableJustification::Left,
            gap_half: 0,
            indent_left: 0,
            preferred_width: None,
            row_height: None,
            is_header_row: false,
            allow_row_break: true,
            has_formatting_revision: None,
            formatting_revision_author_index: None,
            formatting_revision_timestamp: None,
            properties_preserved_for_revision: false,
            border_top: None,
            border_left: None,
            border_bottom: None,
            border_right: None,
            border_horizontal: None,
            border_vertical: None,
        }
    }
}

impl TableProperties {
    /// Create new TableProperties with default values.
    pub fn new() -> Self {
        Self::default()
    }

    /// Create TableProperties with specified cell count.
    ///
    /// Initializes cell boundaries and properties arrays.
    pub fn with_cell_count(cell_count: usize) -> Self {
        Self {
            cell_count,
            cell_boundaries: vec![0; cell_count + 1],
            cell_properties: vec![CellProperties::default(); cell_count],
            ..Default::default()
        }
    }

    /// Parse table properties from SPRM (Single Property Modifier) data.
    ///
    /// # Arguments
    ///
    /// * `grpprl` - Group of SPRMs (property modifications)
    pub fn from_sprm(grpprl: &[u8]) -> Result<Self> {
        let arena = bumpalo::Bump::new();
        super::tap_parser::TapParser::new(&arena).parse_tap(grpprl)
    }

    /// Get cell width in twips for a given cell index.
    pub fn get_cell_width(&self, cell_index: usize) -> Option<i16> {
        if cell_index < self.cell_boundaries.len().saturating_sub(1) {
            Some(self.cell_boundaries[cell_index + 1] - self.cell_boundaries[cell_index])
        } else {
            None
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_default_tap() {
        let tap = TableProperties::new();
        assert_eq!(tap.cell_count, 0);
        assert_eq!(tap.justification, TableJustification::Left);
        assert!(tap.allow_row_break);
    }

    #[test]
    fn test_cell_merge_status() {
        let none = CellMergeStatus::None;
        let first = CellMergeStatus::First;
        assert_ne!(none, first);
    }

    #[test]
    fn test_vertical_alignment() {
        let top = VerticalAlignment::Top;
        let center = VerticalAlignment::Center;
        assert_ne!(top, center);
    }

    #[test]
    fn test_table_definition() {
        let operand = vec![
            2, // 2 cells
            0, 0, // Start at 0
            100, 0, // First boundary at 100 twips
            200, 0, // End at 200 twips
        ];
        let mut data = 0xD608u16.to_le_bytes().to_vec();
        data.extend_from_slice(&u16::try_from(operand.len() + 1).unwrap().to_le_bytes());
        data.extend_from_slice(&operand);

        let tap = TableProperties::from_sprm(&data).unwrap();
        assert_eq!(tap.cell_count, 2);
        assert_eq!(tap.cell_boundaries.len(), 3);
        assert_eq!(tap.get_cell_width(0), Some(100));
        assert_eq!(tap.get_cell_width(1), Some(100));
    }
}
