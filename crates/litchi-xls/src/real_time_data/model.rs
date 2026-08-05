//! Typed semantic values for BIFF8 RealTimeData records.

use crate::error::{XlsError, XlsResult};

/// The last value an RTD server returned for a topic (`RTDOper.rtdVt`).
#[derive(Debug, Clone, PartialEq)]
pub enum Value {
    /// A floating-point value (`Xnum`).
    Number(f64),
    /// A text value (`RTDOperStr`).
    Text(String),
    /// A Boolean value.
    Boolean(bool),
    /// A signed integer that indicates an error code.
    Error(i32),
    /// A signed integer used for purposes other than an error code.
    Integer(i32),
}

/// A cell subscribed to an RTD topic (`RTDEItem`, MS-XLS 2.5.223).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Cell {
    /// Zero-based row index of the cell.
    pub row: u16,
    /// Zero-based column index of the cell.
    pub column: u8,
    /// Zero-based index of the sheet containing the cell (`TabIndex`).
    pub sheet_index: u16,
}

impl Cell {
    /// Create a checked RTD subscriber cell from raw zero-based indices.
    pub fn new(row: u32, column: u16, sheet_index: usize) -> XlsResult<Self> {
        let invalid = || {
            XlsError::InvalidCellReference(format!(
                "RTD subscriber row {row}, column {column} is outside the BIFF8 grid"
            ))
        };
        let row = u16::try_from(row).map_err(|_| invalid())?;
        let column = u8::try_from(column).map_err(|_| invalid())?;
        let sheet_index = u16::try_from(sheet_index)
            .map_err(|_| XlsError::WorksheetNotFound(format!("Sheet {sheet_index}")))?;
        Ok(Self {
            row,
            column,
            sheet_index,
        })
    }
}
/// Typed `RealTimeData` record content (MS-XLS 2.4.214).
#[derive(Debug, Clone, PartialEq)]
pub struct Record {
    /// Number of leading characters this topic shares with the previous
    /// record's topic (`ichSamePrefix`); always zero for the first record.
    pub common_prefix_len: u32,
    /// The topic sub-strings as stored (without the shared prefix). The
    /// first is the RTD server ProgID, the second the server name (empty for
    /// a local server), and the rest combine into the unique topic.
    pub topic_segments: Vec<String>,
    /// The fully reconstructed topic: the shared prefix of the previous
    /// topic followed by the stored sub-strings.
    pub topic: String,
    /// The last value returned by the RTD server (`rtdOper`).
    pub value: Value,
    /// The cells subscribed to this topic (`rgRTDE`).
    pub cells: Vec<Cell>,
}
