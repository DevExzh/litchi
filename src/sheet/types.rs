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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_cell_value_variants() {
        let empty = CellValue::Empty;
        assert!(matches!(empty, CellValue::Empty));

        let boolean = CellValue::Bool(true);
        assert!(matches!(boolean, CellValue::Bool(true)));

        let int = CellValue::Int(42);
        assert!(matches!(int, CellValue::Int(42)));

        let float = CellValue::Float(3.14);
        assert!(matches!(float, CellValue::Float(3.14)));

        let string = CellValue::String("test".to_string());
        assert!(matches!(string, CellValue::String(_)));

        let datetime = CellValue::DateTime(44561.5);
        assert!(matches!(datetime, CellValue::DateTime(44561.5)));

        let error = CellValue::Error("#DIV/0!".to_string());
        assert!(matches!(error, CellValue::Error(_)));
    }

    #[test]
    fn test_cell_value_formula() {
        let formula = CellValue::Formula {
            formula: "A1+B1".to_string(),
            cached_value: Some(Box::new(CellValue::Float(10.0))),
            is_array: false,
            array_range: None,
        };
        assert!(matches!(formula, CellValue::Formula { .. }));
    }

    #[test]
    fn test_as_str() {
        let string_val = CellValue::String("hello".to_string());
        assert_eq!(string_val.as_str(), Some("hello"));

        let int_val = CellValue::Int(42);
        assert_eq!(int_val.as_str(), None);

        let empty_val = CellValue::Empty;
        assert_eq!(empty_val.as_str(), None);
    }

    #[test]
    fn test_as_float() {
        let float_val = CellValue::Float(3.14);
        assert!((float_val.as_float().unwrap() - 3.14).abs() < 0.001);

        let int_val = CellValue::Int(42);
        assert_eq!(int_val.as_float(), None);

        let string_val = CellValue::String("3.14".to_string());
        assert_eq!(string_val.as_float(), None);
    }

    #[test]
    fn test_infer_from_str_empty() {
        assert!(matches!(CellValue::infer_from_str(""), CellValue::Empty));
    }

    #[test]
    fn test_infer_from_str_integer() {
        assert!(matches!(
            CellValue::infer_from_str("42"),
            CellValue::Int(42)
        ));
        assert!(matches!(
            CellValue::infer_from_str("-123"),
            CellValue::Int(-123)
        ));
        assert!(matches!(CellValue::infer_from_str("0"), CellValue::Int(0)));
    }

    #[test]
    fn test_infer_from_str_float() {
        let float_val = CellValue::infer_from_str("3.14");
        assert!(matches!(float_val, CellValue::Float(_)));
        if let CellValue::Float(f) = float_val {
            assert!((f - 3.14).abs() < 0.001);
        }

        let float_val = CellValue::infer_from_str("-0.5");
        assert!(matches!(float_val, CellValue::Float(_)));
    }

    #[test]
    fn test_infer_from_str_boolean() {
        assert!(matches!(
            CellValue::infer_from_str("TRUE"),
            CellValue::Bool(true)
        ));
        assert!(matches!(
            CellValue::infer_from_str("true"),
            CellValue::Bool(true)
        ));
        // Note: "1" and "0" are parsed as integers before boolean check
        assert!(matches!(
            CellValue::infer_from_str("YES"),
            CellValue::Bool(true)
        ));
        assert!(matches!(
            CellValue::infer_from_str("ON"),
            CellValue::Bool(true)
        ));

        assert!(matches!(
            CellValue::infer_from_str("FALSE"),
            CellValue::Bool(false)
        ));
        assert!(matches!(
            CellValue::infer_from_str("false"),
            CellValue::Bool(false)
        ));
        assert!(matches!(
            CellValue::infer_from_str("NO"),
            CellValue::Bool(false)
        ));
        assert!(matches!(
            CellValue::infer_from_str("OFF"),
            CellValue::Bool(false)
        ));
    }

    #[test]
    fn test_infer_from_str_string() {
        assert!(matches!(
            CellValue::infer_from_str("hello"),
            CellValue::String(_)
        ));
        assert!(matches!(
            CellValue::infer_from_str("3.14.15"),
            CellValue::String(_)
        ));
        assert!(matches!(
            CellValue::infer_from_str("abc123"),
            CellValue::String(_)
        ));
    }

    #[test]
    fn test_from_conversions() {
        assert!(matches!(CellValue::from(true), CellValue::Bool(true)));
        assert!(matches!(CellValue::from(false), CellValue::Bool(false)));

        assert!(matches!(CellValue::from(42i32), CellValue::Int(42)));
        assert!(matches!(CellValue::from(42i64), CellValue::Int(42)));
        assert!(matches!(CellValue::from(42u32), CellValue::Int(42)));

        let float_val: CellValue = 3.14f32.into();
        assert!(matches!(float_val, CellValue::Float(_)));

        let float_val: CellValue = 3.14f64.into();
        assert!(matches!(float_val, CellValue::Float(_)));

        assert!(matches!(
            CellValue::from("hello".to_string()),
            CellValue::String(_)
        ));
        assert!(matches!(CellValue::from("hello"), CellValue::String(_)));
        assert!(matches!(
            CellValue::from(&"hello".to_string()),
            CellValue::String(_)
        ));
    }

    #[test]
    fn test_empty_constant() {
        assert!(matches!(CellValue::EMPTY, CellValue::Empty));
    }
}
