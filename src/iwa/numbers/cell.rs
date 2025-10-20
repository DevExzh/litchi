//! Cell Value Types for Numbers Spreadsheets
//!
//! Numbers supports various cell types including text, numbers, dates, formulas, and more.

use std::fmt;

/// Represents a cell value in a Numbers table
#[derive(Debug, Clone, Default)]
pub enum CellValue {
    /// Empty cell
    #[default]
    Empty,
    /// Text/string value
    Text(String),
    /// Numeric value (integer or floating-point)
    Number(f64),
    /// Boolean value
    Boolean(bool),
    /// Date value (stored as timestamp)
    Date(String),
    /// Duration/time value
    Duration(f64),
    /// Formula (stored as string representation)
    Formula(String),
    /// Error value
    Error(String),
}

impl CellValue {
    /// Check if cell is empty
    pub fn is_empty(&self) -> bool {
        matches!(self, CellValue::Empty)
    }

    /// Get cell type
    pub fn cell_type(&self) -> CellType {
        match self {
            CellValue::Empty => CellType::Empty,
            CellValue::Text(_) => CellType::Text,
            CellValue::Number(_) => CellType::Number,
            CellValue::Boolean(_) => CellType::Boolean,
            CellValue::Date(_) => CellType::Date,
            CellValue::Duration(_) => CellType::Duration,
            CellValue::Formula(_) => CellType::Formula,
            CellValue::Error(_) => CellType::Error,
        }
    }

    /// Get text representation of the cell value
    pub fn as_text(&self) -> String {
        match self {
            CellValue::Empty => String::new(),
            CellValue::Text(s) => s.clone(),
            CellValue::Number(n) => format!("{}", n),
            CellValue::Boolean(b) => format!("{}", b),
            CellValue::Date(d) => d.clone(),
            CellValue::Duration(d) => format!("{}", d),
            CellValue::Formula(f) => f.clone(),
            CellValue::Error(e) => format!("ERROR: {}", e),
        }
    }

    /// Try to get as a number
    pub fn as_number(&self) -> Option<f64> {
        match self {
            CellValue::Number(n) => Some(*n),
            CellValue::Duration(d) => Some(*d),
            CellValue::Text(s) => s.parse::<f64>().ok(),
            CellValue::Boolean(b) => Some(if *b { 1.0 } else { 0.0 }),
            _ => None,
        }
    }

    /// Try to get as a boolean
    pub fn as_boolean(&self) -> Option<bool> {
        match self {
            CellValue::Boolean(b) => Some(*b),
            CellValue::Number(n) => Some(*n != 0.0),
            CellValue::Text(s) => match s.to_lowercase().as_str() {
                "true" | "yes" | "1" => Some(true),
                "false" | "no" | "0" => Some(false),
                _ => None,
            },
            _ => None,
        }
    }
}

impl fmt::Display for CellValue {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        // Escape CSV special characters
        let text = self.as_text();
        if text.contains(',') || text.contains('"') || text.contains('\n') {
            write!(f, "\"{}\"", text.replace('"', "\"\""))
        } else {
            write!(f, "{}", text)
        }
    }
}

/// Cell type enumeration
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CellType {
    Empty,
    Text,
    Number,
    Boolean,
    Date,
    Duration,
    Formula,
    Error,
}

impl CellType {
    /// Get human-readable name
    pub fn name(&self) -> &'static str {
        match self {
            CellType::Empty => "Empty",
            CellType::Text => "Text",
            CellType::Number => "Number",
            CellType::Boolean => "Boolean",
            CellType::Date => "Date",
            CellType::Duration => "Duration",
            CellType::Formula => "Formula",
            CellType::Error => "Error",
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_cell_value_types() {
        let empty = CellValue::Empty;
        assert!(empty.is_empty());
        assert_eq!(empty.cell_type(), CellType::Empty);

        let text = CellValue::Text("Hello".to_string());
        assert!(!text.is_empty());
        assert_eq!(text.cell_type(), CellType::Text);
        assert_eq!(text.as_text(), "Hello");

        let number = CellValue::Number(42.5);
        assert_eq!(number.as_number(), Some(42.5));
        assert_eq!(number.as_text(), "42.5");

        let boolean = CellValue::Boolean(true);
        assert_eq!(boolean.as_boolean(), Some(true));
        assert_eq!(boolean.as_number(), Some(1.0));
    }

    #[test]
    fn test_cell_value_display() {
        let text = CellValue::Text("Simple".to_string());
        assert_eq!(format!("{}", text), "Simple");

        // Test CSV escaping
        let with_comma = CellValue::Text("Hello, World".to_string());
        assert_eq!(format!("{}", with_comma), "\"Hello, World\"");

        let with_quote = CellValue::Text("Say \"Hi\"".to_string());
        assert!(format!("{}", with_quote).contains("\"\""));
    }

    #[test]
    fn test_type_conversions() {
        let text_bool = CellValue::Text("true".to_string());
        assert_eq!(text_bool.as_boolean(), Some(true));

        let text_num = CellValue::Text("123.45".to_string());
        assert_eq!(text_num.as_number(), Some(123.45));

        let bool_num = CellValue::Boolean(false);
        assert_eq!(bool_num.as_number(), Some(0.0));
    }
}

