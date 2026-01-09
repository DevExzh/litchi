//! Common types for spreadsheet operations.

/// Error type for spreadsheet operations.
pub type Result<T> = std::result::Result<T, Box<dyn std::error::Error + Send + Sync>>;

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
        /// Whether this formula is an array formula (CSE / dynamic array)
        is_array: bool,
        /// Optional referenced range for array formulas (e.g., "A1:C3")
        array_range: Option<String>,
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

    /// Infer cell value type from string representation.
    ///
    /// This function attempts to parse the string in order:
    /// 1. Empty string -> Empty
    /// 2. Integer -> Int
    /// 3. Float -> Float  
    /// 4. Boolean keywords (TRUE/FALSE/1/0/YES/NO/ON/OFF) -> Bool
    /// 5. Everything else -> String
    pub fn infer_from_str<S: AsRef<str>>(s: S) -> Self {
        let s = s.as_ref();
        if s.is_empty() {
            return Self::Empty;
        }

        if let Ok(i) = s.parse::<i64>() {
            return Self::Int(i);
        }

        if let Ok(f) = fast_float2::parse(s) {
            return Self::Float(f);
        }

        match s.to_uppercase().as_str() {
            "TRUE" | "1" | "YES" | "ON" => Self::Bool(true),
            "FALSE" | "0" | "NO" | "OFF" => Self::Bool(false),
            _ => Self::String(s.to_string()),
        }
    }
}

// Implement From for convenient cell value creation

impl From<bool> for CellValue {
    fn from(b: bool) -> Self {
        Self::Bool(b)
    }
}

impl From<i32> for CellValue {
    fn from(i: i32) -> Self {
        Self::Int(i as i64)
    }
}

impl From<i64> for CellValue {
    fn from(i: i64) -> Self {
        Self::Int(i)
    }
}

impl From<u32> for CellValue {
    fn from(i: u32) -> Self {
        Self::Int(i as i64)
    }
}

impl From<usize> for CellValue {
    fn from(i: usize) -> Self {
        Self::Int(i as i64)
    }
}

impl From<f32> for CellValue {
    fn from(f: f32) -> Self {
        Self::Float(f as f64)
    }
}

impl From<f64> for CellValue {
    fn from(f: f64) -> Self {
        Self::Float(f)
    }
}

impl From<String> for CellValue {
    fn from(s: String) -> Self {
        Self::String(s)
    }
}

impl From<&str> for CellValue {
    fn from(s: &str) -> Self {
        Self::String(s.to_string())
    }
}

impl From<&String> for CellValue {
    fn from(s: &String) -> Self {
        Self::String(s.clone())
    }
}
