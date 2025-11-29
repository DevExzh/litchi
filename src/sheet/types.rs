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
    /// Formula with optional cached result
    Formula {
        /// Formula expression (without leading '=')
        formula: String,
        /// Optional cached result value
        cached_value: Option<Box<CellValue>>,
    },
}

impl CellValue {
    /// Static reference to an empty cell value for zero-copy returns.
    pub const EMPTY: &'static CellValue = &CellValue::Empty;

    /// Get the value as a string slice if it's a String variant.
    pub fn as_str(&self) -> Option<&str> {
        match self {
            CellValue::String(s) => Some(s),
            _ => None,
        }
    }

    /// Get the value as a float if it's a Float variant.
    pub fn as_float(&self) -> Option<f64> {
        match self {
            CellValue::Float(f) => Some(*f),
            _ => None,
        }
    }
}

// Implement From for convenient cell value creation

impl From<bool> for CellValue {
    fn from(b: bool) -> Self {
        CellValue::Bool(b)
    }
}

impl From<i32> for CellValue {
    fn from(i: i32) -> Self {
        CellValue::Int(i as i64)
    }
}

impl From<i64> for CellValue {
    fn from(i: i64) -> Self {
        CellValue::Int(i)
    }
}

impl From<u32> for CellValue {
    fn from(i: u32) -> Self {
        CellValue::Int(i as i64)
    }
}

impl From<usize> for CellValue {
    fn from(i: usize) -> Self {
        CellValue::Int(i as i64)
    }
}

impl From<f32> for CellValue {
    fn from(f: f32) -> Self {
        CellValue::Float(f as f64)
    }
}

impl From<f64> for CellValue {
    fn from(f: f64) -> Self {
        CellValue::Float(f)
    }
}

impl From<String> for CellValue {
    fn from(s: String) -> Self {
        CellValue::String(s)
    }
}

impl From<&str> for CellValue {
    fn from(s: &str) -> Self {
        CellValue::String(s.to_string())
    }
}

impl From<&String> for CellValue {
    fn from(s: &String) -> Self {
        CellValue::String(s.clone())
    }
}
