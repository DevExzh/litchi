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
#[derive(Debug, Clone)]
pub struct TableProperties {
    /// Number of cells in the row
    pub cell_count: usize,
    /// Cell boundaries (positions in twips)
    pub cell_boundaries: Vec<i16>,
    /// Cell properties for each cell
    pub cell_properties: Vec<CellProperties>,
    /// Table justification
    pub justification: TableJustification,
    /// Table indent from left margin (twips)
    pub indent_left: i16,
    /// Preferred table width (twips or percentage)
    pub preferred_width: Option<TableWidth>,
    /// Row height (twips)
    pub row_height: Option<i16>,
    /// Row is header row
    pub is_header_row: bool,
    /// Allow row to break across pages
    pub allow_row_break: bool,
}

/// Cell Properties structure.
///
/// Contains formatting for an individual table cell.
#[derive(Debug, Clone, Default)]
pub struct CellProperties {
    /// Merged cell status
    pub merge_status: CellMergeStatus,
    /// Vertical alignment
    pub vertical_alignment: VerticalAlignment,
    /// Cell background color
    pub background_color: Option<(u8, u8, u8)>,
    /// Cell borders
    pub borders: CellBorders,
    /// Text direction
    pub text_direction: TextDirection,
    /// Preferred cell width
    pub preferred_width: Option<TableWidth>,
}

/// Cell merge status.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum CellMergeStatus {
    /// Not merged
    #[default]
    None,
    /// First cell in merge
    First,
    /// Continuation of merged cell
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
            indent_left: 0,
            preferred_width: None,
            row_height: None,
            is_header_row: false,
            allow_row_break: true,
        }
    }
}

impl TableProperties {
    /// Create new TableProperties with default values.
    pub fn new() -> Self {
        Self::default()
    }

    /// Parse table properties from SPRM (Single Property Modifier) data.
    ///
    /// # Arguments
    ///
    /// * `grpprl` - Group of SPRMs (property modifications)
    pub fn from_sprm(grpprl: &[u8]) -> Result<Self> {
        let mut tap = Self::default();
        let mut offset = 0;

        while offset < grpprl.len() {
            if offset + 1 > grpprl.len() {
                break;
            }

            // Read SPRM opcode
            let sprm = u16::from_le_bytes([grpprl[offset], grpprl[offset + 1]]);
            offset += 2;

            match sprm {
                // Table justification (sprmTJc)
                0x5400 => {
                    if offset < grpprl.len() {
                        tap.justification = match grpprl[offset] {
                            0 => TableJustification::Left,
                            1 => TableJustification::Center,
                            2 => TableJustification::Right,
                            _ => TableJustification::Left,
                        };
                        offset += 1;
                    }
                }
                // Table definition (sprmTDefTable)
                0xD608 => {
                    // This is complex - contains cell count and boundaries
                    if offset + 1 < grpprl.len() {
                        let size = u16::from_le_bytes([grpprl[offset], grpprl[offset + 1]]) as usize;
                        offset += 2;

                        if offset + size <= grpprl.len() {
                            let def_data = &grpprl[offset..offset + size];
                            tap.parse_table_definition(def_data)?;
                            offset += size;
                        }
                    }
                }
                // Row height (sprmTDyaRowHeight)
                0x9407 => {
                    if offset + 1 < grpprl.len() {
                        tap.row_height = Some(i16::from_le_bytes([grpprl[offset], grpprl[offset + 1]]));
                        offset += 2;
                    }
                }
                // Header row (sprmTTableHeader)
                0x3403 => {
                    if offset < grpprl.len() {
                        tap.is_header_row = grpprl[offset] != 0;
                        offset += 1;
                    }
                }
                // Can't split row (sprmTFCantSplit)
                0x3404 => {
                    if offset < grpprl.len() {
                        tap.allow_row_break = grpprl[offset] == 0;
                        offset += 1;
                    }
                }
                // Table indent (sprmTDxaLeft)
                0x9601 => {
                    if offset + 1 < grpprl.len() {
                        tap.indent_left = i16::from_le_bytes([grpprl[offset], grpprl[offset + 1]]);
                        offset += 2;
                    }
                }
                // Cell properties (various sprmTCxxx)
                0xD605..=0xD620 => {
                    // Cell-specific SPRMs
                    let size = Self::get_sprm_size(sprm);
                    if offset + size <= grpprl.len() {
                        // Parse cell properties if needed
                        offset += size;
                    }
                }
                // Unknown SPRM - skip
                _ => {
                    let size = Self::get_sprm_size(sprm);
                    offset += size;
                }
            }
        }

        Ok(tap)
    }

    /// Parse table definition structure.
    ///
    /// Format: cell count (1 byte) + cell boundaries (2 bytes each)
    fn parse_table_definition(&mut self, data: &[u8]) -> Result<()> {
        if data.is_empty() {
            return Ok(());
        }

        // Read cell count
        self.cell_count = data[0] as usize;

        // Read cell boundaries
        let mut offset = 1;
        self.cell_boundaries.clear();

        for _ in 0..=self.cell_count {
            if offset + 1 < data.len() {
                let boundary = i16::from_le_bytes([data[offset], data[offset + 1]]);
                self.cell_boundaries.push(boundary);
                offset += 2;
            }
        }

        // Initialize cell properties
        self.cell_properties = vec![CellProperties::default(); self.cell_count];

        Ok(())
    }

    /// Get the size of an SPRM operand.
    fn get_sprm_size(sprm: u16) -> usize {
        let sprm_type = sprm & 0x07;
        match sprm_type {
            0 | 1 => 1,
            2 | 4 | 5 => 2,
            3 => 4,
            6 => 1, // Variable - simplified
            7 => 3,
            _ => 1,
        }
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
        let mut tap = TableProperties::new();
        
        // Create simple table definition: 2 cells
        // Format: count(1) + boundaries(3 * 2 bytes)
        let data = vec![
            2,         // 2 cells
            0, 0,      // Start at 0
            100, 0,    // First boundary at 100 twips
            200, 0,    // End at 200 twips
        ];

        tap.parse_table_definition(&data).unwrap();
        assert_eq!(tap.cell_count, 2);
        assert_eq!(tap.cell_boundaries.len(), 3);
        assert_eq!(tap.get_cell_width(0), Some(100));
        assert_eq!(tap.get_cell_width(1), Some(100));
    }
}

