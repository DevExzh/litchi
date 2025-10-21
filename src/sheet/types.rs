//! Common types for spreadsheet operations.

/// Error type for spreadsheet operations.
pub type Result<T> = std::result::Result<T, Box<dyn std::error::Error>>;

/// Types of data that can be stored in a cell.
#[derive(Debug, Clone, PartialEq)]
pub enum CellValue {
    /// Empty cell
    Empty,
    /// Boolean value
    Bool(bool),
    /// 64-bit signed integer
    Int(i64),
    /// 64-bit floating point number
    Float(f64),
    /// String value
    String(String),
    /// Date/time value (stored as serial number)
    DateTime(f64),
    /// Error value
    Error(String),
}
