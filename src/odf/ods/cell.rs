//! Cell data structures for ODS spreadsheets.

use crate::common::Result;

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

    /// Parse and get the formula structure.
    ///
    /// Returns the parsed formula if the cell contains a formula,
    /// None otherwise. This provides access to the formula's tokens
    /// and structure for analysis.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi::odf::Spreadsheet;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let sheet = Spreadsheet::open("data.ods")?;
    /// if let Some(sheets) = sheet.sheets().ok() {
    ///     if let Some(first_sheet) = sheets.first() {
    ///         let cell = first_sheet.cell("A1")?;
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
    /// use litchi::odf::Spreadsheet;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let sheet = Spreadsheet::open("data.ods")?;
    /// if let Some(sheets) = sheet.sheets().ok() {
    ///     if let Some(first_sheet) = sheets.first() {
    ///         let cell = first_sheet.cell("A1")?;
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
    /// use litchi::odf::Spreadsheet;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let sheet = Spreadsheet::open("data.ods")?;
    /// if let Some(sheets) = sheet.sheets().ok() {
    ///     if let Some(first_sheet) = sheets.first() {
    ///         let cell = first_sheet.cell("A1")?;
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
    /// use litchi::odf::Spreadsheet;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let sheet = Spreadsheet::open("data.ods")?;
    /// if let Some(sheets) = sheet.sheets().ok() {
    ///     if let Some(first_sheet) = sheets.first() {
    ///         let cell = first_sheet.cell("A1")?;
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
