//! Cell data structures for ODS spreadsheets.

use super::CellAnnotation;
use litchi_core::{Result, xml::escape_xml};

/// Cell data types supported by ODF spreadsheets.
///
/// This enum represents the various data types that can be stored in
/// spreadsheet cells, following the ODF specification.
#[derive(Debug, Clone, PartialEq)]
pub enum CellValue {
    /// Empty cell
    Empty,
    /// Text string
    Text(String),
    /// Numeric value
    Number(f64),
    /// Boolean value
    Boolean(bool),
    /// Date/time value (stored as ISO 8601 string)
    Date(String),
    /// Currency value with currency code
    Currency(f64, String),
    /// Percentage value
    Percentage(f64),
    /// Time duration
    Time(String),
}

/// A cell in an ODS spreadsheet.
///
/// Cells contain typed values, optional formulas, and positioning information.
#[derive(Clone, Debug)]
pub struct Cell {
    /// The cell value
    pub value: CellValue,
    /// The raw text content of the cell
    pub text: String,
    /// The formula in the cell (if any), in ODF format
    pub formula: Option<String>,
    /// The optional ODF annotation (comment/note) attached to the cell.
    pub annotation: Option<CellAnnotation>,
    /// Name of the document-level content validation applied to this cell.
    pub validation_name: Option<String>,
    /// Legacy ODF `table:protect` state, preserved independently when present.
    pub protect: Option<bool>,
    /// ODF `table:protected` state, preserved independently when present.
    pub protected: Option<bool>,
    /// The row index (0-based)
    pub row: usize,
    /// The column index (0-based)
    pub col: usize,
}

impl Cell {
    /// Get the text content of the cell.
    ///
    /// Returns the displayed text value, which may differ from the
    /// underlying typed value for formatted numbers, dates, etc.
    pub fn text(&self) -> Result<&str> {
        Ok(&self.text)
    }

    /// Get the cell value.
    ///
    /// Returns the typed value stored in the cell.
    pub fn value(&self) -> Result<&CellValue> {
        Ok(&self.value)
    }

    /// Get the numeric value of the cell (if applicable).
    ///
    /// Returns `Some(value)` for Number, Currency, and Percentage types,
    /// `None` for other types.
    pub fn numeric_value(&self) -> Result<Option<f64>> {
        match &self.value {
            CellValue::Number(n) => Ok(Some(*n)),
            CellValue::Currency(n, _) => Ok(Some(*n)),
            CellValue::Percentage(p) => Ok(Some(*p)),
            _ => Ok(None),
        }
    }

    /// Get the formula in the cell.
    ///
    /// Returns the formula string if the cell contains a formula,
    /// None otherwise.
    pub fn formula(&self) -> Result<Option<&str>> {
        Ok(self.formula.as_deref())
    }

    /// Return the cell annotation, if present.
    pub fn annotation(&self) -> Option<&CellAnnotation> {
        self.annotation.as_ref()
    }

    /// Return a mutable cell annotation, if present.
    pub fn annotation_mut(&mut self) -> Option<&mut CellAnnotation> {
        self.annotation.as_mut()
    }

    /// Attach or replace the cell annotation.
    pub fn set_annotation(&mut self, annotation: CellAnnotation) {
        self.annotation = Some(annotation);
    }

    /// Remove and return the cell annotation.
    pub fn take_annotation(&mut self) -> Option<CellAnnotation> {
        self.annotation.take()
    }

    /// Check whether this cell has an annotation.
    pub fn has_annotation(&self) -> bool {
        self.annotation.is_some()
    }

    /// Return the document-level content-validation name applied to this cell.
    pub fn validation_name(&self) -> Option<&str> {
        self.validation_name.as_deref()
    }

    /// Apply a named document-level content validation to this cell.
    pub fn set_validation_name(&mut self, name: impl Into<String>) {
        self.validation_name = Some(name.into());
    }

    /// Remove the content-validation reference from this cell.
    pub fn clear_validation(&mut self) {
        self.validation_name = None;
    }

    /// Return the legacy `table:protect` state.
    pub fn protect(&self) -> Option<bool> {
        self.protect
    }

    /// Return the ODF `table:protected` state.
    pub fn protected(&self) -> Option<bool> {
        self.protected
    }

    /// Set both independently representable ODF cell-protection attributes.
    pub fn set_protection(&mut self, protect: Option<bool>, protected: Option<bool>) {
        self.protect = protect;
        self.protected = protected;
    }

    /// Remove both cell-protection attributes.
    pub fn clear_protection(&mut self) {
        self.protect = None;
        self.protected = None;
    }

    /// Parse and get the formula structure.
    ///
    /// Returns the parsed formula if the cell contains a formula,
    /// None otherwise. This provides access to the formula's tokens
    /// and structure for analysis.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::Spreadsheet;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut spreadsheet = Spreadsheet::open("data.ods")?;
    /// let sheets = spreadsheet.sheets()?;
    /// if let Some(first_sheet) = sheets.first() {
    ///     if let Some(cell) = first_sheet.rows.get(0).and_then(|row| row.cells.get(0)) {
    ///         if let Some(parsed_formula) = cell.parsed_formula()? {
    ///             println!("Formula tokens: {:?}", parsed_formula.tokens);
    ///         }
    ///     }
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn parsed_formula(&self) -> Result<Option<super::formula::Formula>> {
        if let Some(formula_str) = &self.formula {
            let parser = super::formula::FormulaParser::new(formula_str);
            Ok(Some(parser.parse()?))
        } else {
            Ok(None)
        }
    }

    /// Check if the cell has a formula.
    ///
    /// Returns true if the cell contains a formula.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::Spreadsheet;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut spreadsheet = Spreadsheet::open("data.ods")?;
    /// let sheets = spreadsheet.sheets()?;
    /// if let Some(first_sheet) = sheets.first() {
    ///     if let Some(cell) = first_sheet.rows.get(0).and_then(|row| row.cells.get(0)) {
    ///         if cell.has_formula() {
    ///             println!("Cell A1 contains a formula");
    ///         }
    ///     }
    /// }
    /// # Ok(())
    /// # }
    /// ```
    #[inline]
    pub fn has_formula(&self) -> bool {
        self.formula.is_some()
    }

    /// Extract cell references from the formula.
    ///
    /// Returns a list of cell references used in the formula.
    /// Returns an empty vector if the cell has no formula.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::Spreadsheet;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut spreadsheet = Spreadsheet::open("data.ods")?;
    /// let sheets = spreadsheet.sheets()?;
    /// if let Some(first_sheet) = sheets.first() {
    ///     if let Some(cell) = first_sheet.rows.get(0).and_then(|row| row.cells.get(0)) {
    ///         let refs = cell.formula_cell_refs()?;
    ///         println!("Cell references: {:?}", refs);
    ///     }
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn formula_cell_refs(&self) -> Result<Vec<super::formula::CellRef>> {
        if let Some(formula) = self.parsed_formula()? {
            Ok(super::formula::extract_cell_refs(&formula)
                .into_iter()
                .cloned()
                .collect())
        } else {
            Ok(Vec::new())
        }
    }

    /// Extract function names used in the formula.
    ///
    /// Returns a list of function names used in the formula.
    /// Returns an empty vector if the cell has no formula.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::Spreadsheet;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut spreadsheet = Spreadsheet::open("data.ods")?;
    /// let sheets = spreadsheet.sheets()?;
    /// if let Some(first_sheet) = sheets.first() {
    ///     if let Some(cell) = first_sheet.rows.get(0).and_then(|row| row.cells.get(0)) {
    ///         let funcs = cell.formula_functions()?;
    ///         println!("Functions used: {:?}", funcs);
    ///     }
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn formula_functions(&self) -> Result<Vec<String>> {
        if let Some(formula) = self.parsed_formula()? {
            Ok(super::formula::extract_functions(&formula)
                .into_iter()
                .map(|s| s.to_string())
                .collect())
        } else {
            Ok(Vec::new())
        }
    }

    /// Get the cell coordinates (row, column).
    ///
    /// Returns a tuple of (row_index, column_index), both 0-based.
    pub fn coordinates(&self) -> (usize, usize) {
        (self.row, self.col)
    }

    /// Check if the cell is empty.
    ///
    /// Returns true if the cell value is `Empty`.
    pub fn is_empty(&self) -> bool {
        matches!(self.value, CellValue::Empty)
    }
}

pub(crate) fn write_cell_xml(output: &mut String, cell: &Cell) {
    output.push_str("<table:table-cell");
    if let Some(formula) = &cell.formula {
        output.push_str(" table:formula=\"");
        output.push_str(&escape_xml(formula));
        output.push('"');
    }
    if let Some(validation_name) = &cell.validation_name {
        output.push_str(" table:content-validation-name=\"");
        output.push_str(&escape_xml(validation_name));
        output.push('"');
    }
    if let Some(protect) = cell.protect {
        output.push_str(if protect {
            " table:protect=\"true\""
        } else {
            " table:protect=\"false\""
        });
    }
    if let Some(protected) = cell.protected {
        output.push_str(if protected {
            " table:protected=\"true\""
        } else {
            " table:protected=\"false\""
        });
    }

    match &cell.value {
        CellValue::Text(_) => output.push_str(" office:value-type=\"string\""),
        CellValue::Number(value) => {
            output.push_str(" office:value-type=\"float\" office:value=\"");
            output.push_str(&value.to_string());
            output.push('"');
        },
        CellValue::Currency(value, currency) => {
            output.push_str(" office:value-type=\"currency\" office:value=\"");
            output.push_str(&value.to_string());
            output.push_str("\" office:currency=\"");
            output.push_str(&escape_xml(currency));
            output.push('"');
        },
        CellValue::Percentage(value) => {
            output.push_str(" office:value-type=\"percentage\" office:value=\"");
            output.push_str(&value.to_string());
            output.push('"');
        },
        CellValue::Date(value) => {
            output.push_str(" office:value-type=\"date\" office:date-value=\"");
            output.push_str(&escape_xml(value));
            output.push('"');
        },
        CellValue::Time(value) => {
            output.push_str(" office:value-type=\"time\" office:time-value=\"");
            output.push_str(&escape_xml(value));
            output.push('"');
        },
        CellValue::Boolean(value) => {
            output.push_str(" office:value-type=\"boolean\" office:boolean-value=\"");
            output.push_str(if *value { "true" } else { "false" });
            output.push('"');
        },
        CellValue::Empty if cell.formula.is_some() => {
            output.push_str(" office:value-type=\"float\" office:value=\"0\"");
        },
        CellValue::Empty if cell.annotation.is_none() => {
            output.push_str("/>");
            return;
        },
        CellValue::Empty => {},
    }

    output.push('>');
    if let Some(annotation) = &cell.annotation {
        annotation.write_xml(output);
    }
    if matches!(cell.value, CellValue::Empty) {
        if cell.formula.is_some() {
            output.push_str("<text:p>0</text:p>");
        }
    } else {
        output.push_str("<text:p>");
        output.push_str(&escape_xml(&cell.text));
        output.push_str("</text:p>");
    }
    output.push_str("</table:table-cell>");
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_cell_value_empty() {
        let value = CellValue::Empty;
        assert_eq!(value, CellValue::Empty);
    }

    #[test]
    fn test_cell_value_text() {
        let value = CellValue::Text("Hello".to_string());
        assert_eq!(value, CellValue::Text("Hello".to_string()));
    }

    #[test]
    fn test_cell_value_number() {
        let value = CellValue::Number(42.5);
        assert_eq!(value, CellValue::Number(42.5));
    }

    #[test]
    fn test_cell_value_boolean() {
        let value = CellValue::Boolean(true);
        assert_eq!(value, CellValue::Boolean(true));
    }

    #[test]
    fn test_cell_value_date() {
        let value = CellValue::Date("2024-01-15".to_string());
        assert_eq!(value, CellValue::Date("2024-01-15".to_string()));
    }

    #[test]
    fn test_cell_value_currency() {
        let value = CellValue::Currency(100.0, "USD".to_string());
        match value {
            CellValue::Currency(amount, currency) => {
                assert!((amount - 100.0).abs() < f64::EPSILON);
                assert_eq!(currency, "USD");
            },
            _ => panic!("Expected Currency"),
        }
    }

    #[test]
    fn test_cell_value_percentage() {
        let value = CellValue::Percentage(0.25);
        assert_eq!(value, CellValue::Percentage(0.25));
    }

    #[test]
    fn test_cell_value_time() {
        let value = CellValue::Time("PT1H30M".to_string());
        assert_eq!(value, CellValue::Time("PT1H30M".to_string()));
    }

    #[test]
    fn test_cell_new() {
        let cell = Cell {
            value: CellValue::Empty,
            text: String::new(),
            formula: None,
            annotation: None,
            validation_name: None,
            protect: None,
            protected: None,
            row: 0,
            col: 0,
        };
        assert!(cell.is_empty());
        assert_eq!(cell.text, "");
        assert!(cell.formula.is_none());
    }

    #[test]
    fn test_cell_text() {
        let cell = Cell {
            value: CellValue::Text("Hello".to_string()),
            text: "Hello".to_string(),
            formula: None,
            annotation: None,
            validation_name: None,
            protect: None,
            protected: None,
            row: 0,
            col: 0,
        };
        assert_eq!(cell.text().unwrap(), "Hello");
    }

    #[test]
    fn test_cell_value() {
        let cell = Cell {
            value: CellValue::Number(42.0),
            text: "42".to_string(),
            formula: None,
            annotation: None,
            validation_name: None,
            protect: None,
            protected: None,
            row: 0,
            col: 0,
        };
        match cell.value().unwrap() {
            CellValue::Number(n) => assert!((n - 42.0).abs() < f64::EPSILON),
            _ => panic!("Expected Number"),
        }
    }

    #[test]
    fn test_cell_numeric_value() {
        let cell = Cell {
            value: CellValue::Number(42.0),
            text: "42".to_string(),
            formula: None,
            annotation: None,
            validation_name: None,
            protect: None,
            protected: None,
            row: 0,
            col: 0,
        };
        assert_eq!(cell.numeric_value().unwrap(), Some(42.0));

        let cell = Cell {
            value: CellValue::Currency(100.0, "USD".to_string()),
            text: "$100".to_string(),
            formula: None,
            annotation: None,
            validation_name: None,
            protect: None,
            protected: None,
            row: 0,
            col: 0,
        };
        assert_eq!(cell.numeric_value().unwrap(), Some(100.0));

        let cell = Cell {
            value: CellValue::Percentage(0.5),
            text: "50%".to_string(),
            formula: None,
            annotation: None,
            validation_name: None,
            protect: None,
            protected: None,
            row: 0,
            col: 0,
        };
        assert_eq!(cell.numeric_value().unwrap(), Some(0.5));

        let cell = Cell {
            value: CellValue::Text("Hello".to_string()),
            text: "Hello".to_string(),
            formula: None,
            annotation: None,
            validation_name: None,
            protect: None,
            protected: None,
            row: 0,
            col: 0,
        };
        assert_eq!(cell.numeric_value().unwrap(), None);
    }

    #[test]
    fn test_cell_formula() {
        let cell = Cell {
            value: CellValue::Number(42.0),
            text: "42".to_string(),
            formula: Some("=A1+B1".to_string()),
            annotation: None,
            validation_name: None,
            protect: None,
            protected: None,
            row: 0,
            col: 0,
        };
        assert_eq!(cell.formula().unwrap(), Some("=A1+B1"));
    }

    #[test]
    fn test_cell_no_formula() {
        let cell = Cell {
            value: CellValue::Text("Hello".to_string()),
            text: "Hello".to_string(),
            formula: None,
            annotation: None,
            validation_name: None,
            protect: None,
            protected: None,
            row: 0,
            col: 0,
        };
        assert_eq!(cell.formula().unwrap(), None);
    }

    #[test]
    fn test_cell_has_formula() {
        let cell_with = Cell {
            value: CellValue::Number(42.0),
            text: "42".to_string(),
            formula: Some("=A1".to_string()),
            annotation: None,
            validation_name: None,
            protect: None,
            protected: None,
            row: 0,
            col: 0,
        };
        assert!(cell_with.has_formula());

        let cell_without = Cell {
            value: CellValue::Text("Hello".to_string()),
            text: "Hello".to_string(),
            formula: None,
            annotation: None,
            validation_name: None,
            protect: None,
            protected: None,
            row: 0,
            col: 0,
        };
        assert!(!cell_without.has_formula());
    }

    #[test]
    fn test_cell_coordinates() {
        let cell = Cell {
            value: CellValue::Empty,
            text: String::new(),
            formula: None,
            annotation: None,
            validation_name: None,
            protect: None,
            protected: None,
            row: 5,
            col: 10,
        };
        assert_eq!(cell.coordinates(), (5, 10));
    }

    #[test]
    fn test_cell_is_empty() {
        let empty_cell = Cell {
            value: CellValue::Empty,
            text: String::new(),
            formula: None,
            annotation: None,
            validation_name: None,
            protect: None,
            protected: None,
            row: 0,
            col: 0,
        };
        assert!(empty_cell.is_empty());

        let text_cell = Cell {
            value: CellValue::Text("Hello".to_string()),
            text: "Hello".to_string(),
            formula: None,
            annotation: None,
            validation_name: None,
            protect: None,
            protected: None,
            row: 0,
            col: 0,
        };
        assert!(!text_cell.is_empty());
    }

    #[test]
    fn test_cell_equality() {
        let cell1 = CellValue::Number(42.0);
        let cell2 = CellValue::Number(42.0);
        let cell3 = CellValue::Number(43.0);

        assert_eq!(cell1, cell2);
        assert_ne!(cell1, cell3);
    }

    #[test]
    fn test_cell_clone() {
        let cell = CellValue::Text("Hello".to_string());
        let cloned = cell.clone();
        assert_eq!(cell, cloned);
    }
}
