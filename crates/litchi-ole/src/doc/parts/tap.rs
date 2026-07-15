/// Table Properties (TAP) parser for DOC files.
///
/// TAP structures define table-level formatting including:
/// - Table borders and shading
/// - Row and cell definitions
/// - Table positioning
/// - Cell margins and spacing
use super::super::package::Result;

bitflags::bitflags! {
    /// Optional formats enabled by a DOC table style or auto-format (`Fatl`).
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
    pub struct TableLookFlags: u16 {
        const BORDERS = 0x0001;
        const SHADING = 0x0002;
        const FONT = 0x0004;
        const COLOR = 0x0008;
        const BEST_FIT = 0x0010;
        const HEADER_ROW = 0x0020;
        const LAST_ROW = 0x0040;
        const HEADER_COLUMN = 0x0080;
        const LAST_COLUMN = 0x0100;
        const NO_ROW_BANDING = 0x0200;
        const NO_COLUMN_BANDING = 0x0400;
    }
}

/// Table auto-format identity and optional look flags (`TLP`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TableLook {
    /// Application-specific predefined auto-format index; -1 means none.
    pub autoformat_index: i16,
    /// Optional formats enabled by the table style or auto-format.
    pub flags: TableLookFlags,
}

/// Vertical origin used for an absolutely positioned DOC table.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableVerticalAnchor {
    Margin,
    Page,
    Paragraph,
    None,
}

/// Horizontal origin used for an absolutely positioned DOC table.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableHorizontalAnchor {
    Column,
    Margin,
    Page,
    None,
}

/// Anchor origins for a floating DOC table (`PositionCodeOperand`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TablePositioning {
    pub vertical_anchor: TableVerticalAnchor,
    pub horizontal_anchor: TableHorizontalAnchor,
}

/// Source of a DOC table's uniform inter-cell spacing.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CellSpacingSource {
    /// Explicit twip spacing (`ftsDxa`).
    Explicit,
    /// Spacing produced by table-border application (`ftsDxaSys`).
    TableBorder,
}

/// Uniform spacing applied around every cell in a DOC table row.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CellSpacing {
    pub width: u16,
    pub source: CellSpacingSource,
}

/// Horizontal table position, including the special alignment sentinels.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum TableHorizontalPosition {
    #[default]
    Left,
    Center,
    Right,
    Inside,
    Outside,
    /// Physical offset in twips after decoding XAS_plusOne.
    Offset(i16),
}

/// Vertical table position, including the special alignment sentinels.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum TableVerticalPosition {
    #[default]
    Inline,
    Top,
    Center,
    Bottom,
    Inside,
    Outside,
    /// Downward offset in twips after decoding YAS_plusOne.
    Offset(i16),
}

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
    /// Resolved logical table justification
    pub justification: TableJustification,
    /// Legacy physical justification source, if explicitly specified
    pub legacy_physical_justification: Option<TableJustification>,
    /// Modern logical justification source, if explicitly specified
    pub modern_logical_justification: Option<TableJustification>,
    /// Scalar defaults parsed from a table style's `UpxTapx`
    pub style_defaults: TableStyleDefaults,
    /// Half the width of spacing between cells (dxaGapHalf)
    pub gap_half: i16,
    /// Table indent from left margin (twips)
    pub indent_left: i16,
    /// Preferred table width
    pub preferred_width: Option<TableWidth>,
    /// Automatically resize columns to fit table contents
    pub auto_fit: bool,
    /// Preferred leading space before the first cell
    pub width_before: Option<TableWidth>,
    /// Preferred trailing space after the last cell
    pub width_after: Option<TableWidth>,
    /// Preferred leading indentation of the table
    pub preferred_indent: Option<TableWidth>,
    /// Avoid a page break between this row and the following row
    pub keep_with_next: bool,
    /// Table auto-format identity and enabled optional formatting
    pub table_look: Option<TableLook>,
    /// Style-sheet index of the applied table style
    pub table_style_index: Option<u16>,
    /// Final right-to-left layout state
    pub right_to_left: bool,
    /// Legacy right-to-left source retained for correct Bool16 OR semantics
    pub legacy_right_to_left: bool,
    /// Modern right-to-left source retained for correct Bool16 OR semantics
    pub modern_right_to_left: bool,
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
    /// Nonzero `PGPInfo.ipgpSelf` associated with this row
    pub paragraph_group_id: Option<u32>,
    /// Revision save ID associated with this table formatting
    pub revision_save_id: Option<u32>,
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
    /// Complete legacy cell shading descriptor
    pub shading: Option<CellShading>,
    /// `ShdNil` from a raw shading operand defers to the table style
    pub shading_inherits_from_style: bool,
    /// Cell borders
    pub borders: CellBorders,
    /// Border-type-only overrides; `None` means inherit that side's type
    pub border_type_overrides: CellBorderTypes,
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

/// Per-side border type overrides from `sprmTCellBrcType`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct CellBorderTypes {
    pub top: Option<BorderType>,
    pub left: Option<BorderType>,
    pub bottom: Option<BorderType>,
    pub right: Option<BorderType>,
}

/// Scalar cell and band defaults stored in a DOC table style.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct TableStyleDefaults {
    pub padding_top: Option<u16>,
    pub padding_left: Option<u16>,
    pub padding_bottom: Option<u16>,
    pub padding_right: Option<u16>,
    pub vertical_alignment: Option<VerticalAlignment>,
    pub no_wrap: Option<bool>,
    pub horizontal_band_size: Option<u8>,
    pub vertical_band_size: Option<u8>,
    pub border_top: Option<TableStyleBorder>,
    pub border_bottom: Option<TableStyleBorder>,
    pub border_left: Option<TableStyleBorder>,
    pub border_right: Option<TableStyleBorder>,
    pub border_inside_horizontal: Option<TableStyleBorder>,
    pub border_inside_vertical: Option<TableStyleBorder>,
    pub border_diagonal_down: Option<TableStyleBorder>,
    pub border_diagonal_up: Option<TableStyleBorder>,
    pub shading: Option<TableStyleShading>,
}

/// An explicitly specified border value in a DOC table style.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableStyleBorder {
    /// Explicitly clear the border.
    NoBorder,
    /// Apply this border style.
    Border(BorderStyle),
}

/// An explicitly specified shading value in a DOC table style.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableStyleShading {
    /// `ShdAuto`: explicitly apply no shading.
    NoShading,
    /// Apply this shading descriptor.
    Shading(CellShading),
}

/// Legacy Word table-cell shading.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct CellShading {
    /// Pattern foreground color, or automatic when absent
    pub foreground_color: Option<(u8, u8, u8)>,
    /// Pattern background color, or automatic when absent
    pub background_color: Option<(u8, u8, u8)>,
    /// Pattern used to combine the foreground and background colors
    pub pattern: ShadingPattern,
}

/// Shading patterns representable by the two-byte DOC `Shd80` structure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
#[repr(u8)]
pub enum ShadingPattern {
    #[default]
    Auto = 0x00,
    Solid = 0x01,
    Percent5 = 0x02,
    Percent10 = 0x03,
    Percent20 = 0x04,
    Percent25 = 0x05,
    Percent30 = 0x06,
    Percent40 = 0x07,
    Percent50 = 0x08,
    Percent60 = 0x09,
    Percent70 = 0x0A,
    Percent75 = 0x0B,
    Percent80 = 0x0C,
    Percent90 = 0x0D,
    DarkHorizontal = 0x0E,
    DarkVertical = 0x0F,
    DarkReverseDiagonal = 0x10,
    DarkDiagonal = 0x11,
    DarkCross = 0x12,
    DarkDiagonalCross = 0x13,
    Horizontal = 0x14,
    Vertical = 0x15,
    ReverseDiagonal = 0x16,
    Diagonal = 0x17,
    Cross = 0x18,
    DiagonalCross = 0x19,
    Percent2Point5 = 0x23,
    Percent7Point5 = 0x24,
    Percent12Point5 = 0x25,
    Percent15 = 0x26,
    Percent17Point5 = 0x27,
    Percent22Point5 = 0x28,
    Percent27Point5 = 0x29,
    Percent32Point5 = 0x2A,
    Percent35 = 0x2B,
    Percent37Point5 = 0x2C,
    Percent42Point5 = 0x2D,
    Percent45 = 0x2E,
    Percent47Point5 = 0x2F,
    Percent52Point5 = 0x30,
    Percent55 = 0x31,
    Percent57Point5 = 0x32,
    Percent62Point5 = 0x33,
    Percent65 = 0x34,
    Percent67Point5 = 0x35,
    Percent72Point5 = 0x36,
    Percent77Point5 = 0x37,
    Percent82Point5 = 0x38,
    Percent85 = 0x39,
    Percent87Point5 = 0x3A,
    Percent92Point5 = 0x3B,
    Percent95 = 0x3C,
    Percent97Point5 = 0x3D,
}

impl ShadingPattern {
    pub(crate) fn from_u8(value: u8) -> Option<Self> {
        Some(match value {
            0x00 => Self::Auto,
            0x01 => Self::Solid,
            0x02 => Self::Percent5,
            0x03 => Self::Percent10,
            0x04 => Self::Percent20,
            0x05 => Self::Percent25,
            0x06 => Self::Percent30,
            0x07 => Self::Percent40,
            0x08 => Self::Percent50,
            0x09 => Self::Percent60,
            0x0A => Self::Percent70,
            0x0B => Self::Percent75,
            0x0C => Self::Percent80,
            0x0D => Self::Percent90,
            0x0E => Self::DarkHorizontal,
            0x0F => Self::DarkVertical,
            0x10 => Self::DarkReverseDiagonal,
            0x11 => Self::DarkDiagonal,
            0x12 => Self::DarkCross,
            0x13 => Self::DarkDiagonalCross,
            0x14 => Self::Horizontal,
            0x15 => Self::Vertical,
            0x16 => Self::ReverseDiagonal,
            0x17 => Self::Diagonal,
            0x18 => Self::Cross,
            0x19 => Self::DiagonalCross,
            0x23 => Self::Percent2Point5,
            0x24 => Self::Percent7Point5,
            0x25 => Self::Percent12Point5,
            0x26 => Self::Percent15,
            0x27 => Self::Percent17Point5,
            0x28 => Self::Percent22Point5,
            0x29 => Self::Percent27Point5,
            0x2A => Self::Percent32Point5,
            0x2B => Self::Percent35,
            0x2C => Self::Percent37Point5,
            0x2D => Self::Percent42Point5,
            0x2E => Self::Percent45,
            0x2F => Self::Percent47Point5,
            0x30 => Self::Percent52Point5,
            0x31 => Self::Percent55,
            0x32 => Self::Percent57Point5,
            0x33 => Self::Percent62Point5,
            0x34 => Self::Percent65,
            0x35 => Self::Percent67Point5,
            0x36 => Self::Percent72Point5,
            0x37 => Self::Percent77Point5,
            0x38 => Self::Percent82Point5,
            0x39 => Self::Percent85,
            0x3A => Self::Percent87Point5,
            0x3B => Self::Percent92Point5,
            0x3C => Self::Percent95,
            0x3D => Self::Percent97Point5,
            _ => return None,
        })
    }
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
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
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
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct CellBorders {
    pub top: Option<BorderStyle>,
    pub left: Option<BorderStyle>,
    pub bottom: Option<BorderStyle>,
    pub right: Option<BorderStyle>,
    /// Diagonal from the top-left to the bottom-right
    pub diagonal_down: Option<BorderStyle>,
    /// Diagonal from the top-right to the bottom-left
    pub diagonal_up: Option<BorderStyle>,
}

/// Border style.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct BorderStyle {
    /// Line width in 1/8 points
    pub width: u8,
    /// Border color
    pub color: Option<(u8, u8, u8)>,
    /// Border type
    pub border_type: BorderType,
    /// Distance from cell contents in points
    pub spacing: u8,
    /// Draw a shadow effect
    pub shadow: bool,
    /// Reverse the border for a frame effect
    pub frame: bool,
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
    ThinThickMedium,
    ThickThinMedium,
    ThinThickThinMedium,
    ThinThickLarge,
    ThickThinLarge,
    ThinThickThinLarge,
    Wave,
    DoubleWave,
    DashSmall,
    DashDotStroked,
    Emboss,
    Engrave,
    Outset,
    Inset,
}

impl Default for TableProperties {
    fn default() -> Self {
        Self {
            cell_count: 0,
            cell_boundaries: Vec::new(),
            cell_properties: Vec::new(),
            justification: TableJustification::Left,
            legacy_physical_justification: None,
            modern_logical_justification: None,
            style_defaults: TableStyleDefaults::default(),
            gap_half: 0,
            indent_left: 0,
            preferred_width: None,
            auto_fit: false,
            width_before: None,
            width_after: None,
            preferred_indent: None,
            keep_with_next: false,
            table_look: None,
            table_style_index: None,
            right_to_left: false,
            legacy_right_to_left: false,
            modern_right_to_left: false,
            allow_overlap: true,
            positioning: None,
            horizontal_position: TableHorizontalPosition::Left,
            vertical_position: TableVerticalPosition::Inline,
            distance_from_text_left: 0,
            distance_from_text_top: 0,
            distance_from_text_right: 0,
            distance_from_text_bottom: 0,
            cell_spacing: None,
            row_height: None,
            is_header_row: false,
            allow_row_break: true,
            has_formatting_revision: None,
            formatting_revision_author_index: None,
            formatting_revision_timestamp: None,
            properties_preserved_for_revision: false,
            paragraph_group_id: None,
            revision_save_id: None,
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
