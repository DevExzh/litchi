//! TAP (Table Properties) generation for DOC files
//!
//! TAP structures define table layout, borders, and cell properties.
//!
//! Based on Microsoft's "[MS-DOC]" specification and Apache POI's TableProperties.

use super::sprm::SprmBuilder;

/// Error returned when table row properties cannot be represented in DOC TAP.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum TapBuildError {
    /// Requested row index is not present in the builder.
    RowOutOfBounds(usize),
    /// DOC table rows can contain at most 63 cells.
    InvalidCellCount(usize),
    /// Cumulative cell boundaries exceed signed 16-bit twip coordinates.
    CellWidthsOverflow,
    /// Positive row heights are limited to signed 16-bit twip values.
    InvalidRowHeight(u16),
    /// A merge continuation cannot occur in the first cell.
    MergeWithoutPrecedingCell,
}

impl std::fmt::Display for TapBuildError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::RowOutOfBounds(index) => write!(f, "table row {index} does not exist"),
            Self::InvalidCellCount(count) => {
                write!(
                    f,
                    "DOC table rows must contain between 1 and 63 cells, got {count}"
                )
            },
            Self::CellWidthsOverflow => {
                write!(
                    f,
                    "DOC cell widths exceed the signed 16-bit coordinate range"
                )
            },
            Self::InvalidRowHeight(height) => {
                write!(f, "DOC row height {height} exceeds 32767 twips")
            },
            Self::MergeWithoutPrecedingCell => {
                write!(f, "the first DOC table cell cannot be a merge continuation")
            },
        }
    }
}

impl std::error::Error for TapBuildError {}

/// Table cell descriptor
#[derive(Debug, Clone, Default)]
pub struct TableCell {
    /// Cell width (in twips)
    pub width: u16,
    /// This cell is merged into the preceding cell. The preceding cell is
    /// automatically encoded as the start of the horizontal merge.
    pub merged: bool,
}

/// Table row properties
#[derive(Debug, Clone, Default)]
pub struct TableRow {
    /// Cells in this row
    pub cells: Vec<TableCell>,
    /// Row height (in twips)
    pub height: u16,
    /// Header row flag
    pub is_header: bool,
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
        generate_row_sprms(row)
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

pub(crate) fn generate_row_sprms(row: &TableRow) -> Result<Vec<u8>, TapBuildError> {
    let cell_count = row.cells.len();
    if !(1..=63).contains(&cell_count) {
        return Err(TapBuildError::InvalidCellCount(cell_count));
    }
    if row.cells.first().is_some_and(|cell| cell.merged) {
        return Err(TapBuildError::MergeWithoutPrecedingCell);
    }
    if row.height > i16::MAX as u16 {
        return Err(TapBuildError::InvalidRowHeight(row.height));
    }

    let effective_widths = if row.cells.iter().all(|cell| cell.width == 0) {
        const DEFAULT_TABLE_WIDTH: u32 = 8640;
        (0..cell_count)
            .map(|index| {
                let left = DEFAULT_TABLE_WIDTH * index as u32 / cell_count as u32;
                let right = DEFAULT_TABLE_WIDTH * (index + 1) as u32 / cell_count as u32;
                (right - left) as u16
            })
            .collect::<Vec<_>>()
    } else {
        row.cells.iter().map(|cell| cell.width).collect()
    };

    let mut boundaries = Vec::with_capacity(cell_count + 1);
    boundaries.push(0i16);
    let mut boundary = 0u32;
    for width in &effective_widths {
        boundary = boundary
            .checked_add(u32::from(*width))
            .ok_or(TapBuildError::CellWidthsOverflow)?;
        boundaries.push(i16::try_from(boundary).map_err(|_| TapBuildError::CellWidthsOverflow)?);
    }

    let mut builder = SprmBuilder::new();
    if row.is_header {
        builder.add_bool(0x3404, true);
    }
    if row.height > 0 {
        builder.add_word(0x9407, row.height);
    }
    let mut sprms = builder.build();

    let mut operand = Vec::with_capacity(1 + (cell_count + 1) * 2 + cell_count * 20);
    operand.push(cell_count as u8);
    for boundary in boundaries {
        operand.extend_from_slice(&boundary.to_le_bytes());
    }
    for (index, width) in effective_widths.into_iter().enumerate() {
        let horizontal_merge = if row.cells[index].merged {
            1
        } else if row.cells.get(index + 1).is_some_and(|cell| cell.merged) {
            2
        } else {
            0
        };
        let flags = horizontal_merge | (3u16 << 9); // horzMerge + ftsDxa
        operand.extend_from_slice(&flags.to_le_bytes());
        operand.extend_from_slice(&width.to_le_bytes());
        operand.extend_from_slice(&[0; 16]); // Four default Brc80MayBeNil values.
    }

    let encoded_size =
        u16::try_from(operand.len() + 1).map_err(|_| TapBuildError::CellWidthsOverflow)?;
    sprms.extend_from_slice(&0xD608u16.to_le_bytes());
    sprms.extend_from_slice(&encoded_size.to_le_bytes());
    sprms.extend_from_slice(&operand);
    Ok(sprms)
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
                merged: false
            };
            cols
        ];
        builder.add_row(TableRow {
            cells,
            height: 0, // Auto height
            is_header: false,
        });
    }

    builder
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_tap_builder() {
        let mut builder = TapBuilder::new();
        builder.add_row(TableRow {
            cells: vec![
                TableCell {
                    width: 1000,
                    merged: false,
                },
                TableCell {
                    width: 1000,
                    merged: false,
                },
            ],
            height: 200,
            is_header: true,
        });

        let sprms = builder.try_generate_row_sprms(0).unwrap();
        let tap = crate::doc::parts::tap::TableProperties::from_sprm(&sprms).unwrap();
        assert_eq!(tap.cell_boundaries, [0, 1000, 2000]);
        assert_eq!(tap.row_height, Some(200));
        assert!(tap.is_header_row);
        assert_eq!(
            tap.cell_properties[0].preferred_width.unwrap().width_type,
            crate::doc::parts::tap::WidthType::Twips
        );
    }

    #[test]
    fn test_tap_builder_empty() {
        let builder = TapBuilder::new();
        assert_eq!(builder.row_count(), 0);
        assert_eq!(
            builder.try_generate_row_sprms(0),
            Err(TapBuildError::RowOutOfBounds(0))
        );
    }

    #[test]
    fn test_tap_builder_multiple_rows() {
        let mut builder = TapBuilder::new();
        for i in 0..5 {
            builder.add_row(TableRow {
                cells: vec![
                    TableCell {
                        width: 1000,
                        merged: false,
                    },
                    TableCell {
                        width: 1000,
                        merged: false,
                    },
                    TableCell {
                        width: 1000,
                        merged: false,
                    },
                ],
                height: 200 + (i as u16 * 50),
                is_header: i == 0,
            });
        }

        assert_eq!(builder.row_count(), 5);
        let sprms = builder.generate_row_sprms(0);
        assert!(!sprms.is_empty());
    }

    #[test]
    fn test_create_simple_table() {
        let table = create_simple_table(3, 4, 1440); // 3 rows, 4 cols, 1 inch cells
        assert_eq!(table.row_count(), 3);
    }

    #[test]
    fn test_create_simple_table_single_cell() {
        let table = create_simple_table(1, 1, 1000);
        assert_eq!(table.row_count(), 1);
        assert_eq!(table.rows[0].cells.len(), 1);
    }

    #[test]
    fn test_create_simple_table_large() {
        let table = create_simple_table(10, 10, 500);
        assert_eq!(table.row_count(), 10);
        assert_eq!(table.rows[0].cells.len(), 10);
    }

    #[test]
    fn test_table_row_count() {
        let table = create_simple_table(5, 3, 1000);
        assert_eq!(table.row_count(), 5);
    }

    #[test]
    fn rejects_unrepresentable_rows() {
        let mut builder = TapBuilder::new();
        builder.add_row(TableRow {
            cells: vec![TableCell {
                width: 1000,
                merged: true,
            }],
            height: 0,
            is_header: false,
        });
        assert_eq!(
            builder.try_generate_row_sprms(0),
            Err(TapBuildError::MergeWithoutPrecedingCell)
        );

        let mut builder = TapBuilder::new();
        builder.add_row(TableRow {
            cells: vec![TableCell {
                width: u16::MAX,
                merged: false,
            }],
            height: 0,
            is_header: false,
        });
        assert_eq!(
            builder.try_generate_row_sprms(0),
            Err(TapBuildError::CellWidthsOverflow)
        );
    }
}
