//! Unified spreadsheet API for Excel files.
//!
//! This module provides a unified interface for working with Excel spreadsheets,
//! similar to openpyxl but adapted for Rust idioms and performance.
//!
//! # Architecture
//!
//! The module defines traits that represent the core concepts:
//!
//! - [`Workbook`]: The top-level container for spreadsheet data
//! - [`Worksheet`]: Individual sheets within a workbook
//! - [`Cell`]: Individual cells containing data
//!
//! # Example
//!
//! ```rust,no_run
//! use litchi::sheet::{Workbook, Worksheet};
//!
//! // Open a workbook
//! let workbook = Workbook::open("workbook.xlsx")?;
//!
//! // Access worksheets
//! for worksheet in workbook.worksheets() {
//!     let worksheet = worksheet?;
//!     println!("Sheet: {}", worksheet.name());
//!
//!     // Access cells
//!     let cell = worksheet.cell(1, 1)?;
//!     println!("A1 value: {:?}", cell.value());
//!
//!     // Access by coordinate
//!     let cell_a1 = worksheet.cell_by_coordinate("A1")?;
//!     println!("A1 coordinate access: {:?}", cell_a1.value());
//!
//!     // Iterate over all cells
//!     for cell_result in worksheet.cells() {
//!         let cell = cell_result?;
//!         println!("Cell {}: {:?}", cell.coordinate(), cell.value());
//!     }
//! }
//!
//! // Access specific worksheet by name
//! let sheet1 = workbook.worksheet_by_name("Sheet1")?;
//! println!("Found worksheet: {}", sheet1.name());
//!
//! // Get active worksheet
//! let active = workbook.active_worksheet()?;
//! println!("Active sheet: {}", active.name());
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```

use std::fmt::Debug;

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

/// Represents an individual cell in a worksheet.
pub trait Cell {
    /// Get the row number (1-based).
    fn row(&self) -> u32;

    /// Get the column number (1-based).
    fn column(&self) -> u32;

    /// Get the cell coordinate (e.g., "A1").
    fn coordinate(&self) -> String;

    /// Get the cell value.
    fn value(&self) -> &CellValue;

    /// Check if the cell is empty.
    fn is_empty(&self) -> bool {
        matches!(self.value(), CellValue::Empty)
    }

    /// Check if the cell contains a formula.
    fn is_formula(&self) -> bool {
        false // Default implementation, can be overridden
    }

    /// Check if the cell contains a date/time value.
    fn is_date(&self) -> bool {
        matches!(self.value(), CellValue::DateTime(_))
    }
}

/// Iterator over cells in a worksheet.
pub trait CellIterator<'a> {
    /// Get the next cell.
    fn next(&mut self) -> Option<Result<Box<dyn Cell + 'a>>>;
}

/// Iterator over rows in a worksheet.
pub trait RowIterator<'a> {
    /// Get the next row (as a vector of cell values).
    fn next(&mut self) -> Option<Result<Vec<CellValue>>>;
}

/// Represents a worksheet (sheet) in a workbook.
pub trait Worksheet {
    /// Get the worksheet name.
    fn name(&self) -> &str;

    /// Get the number of rows in the worksheet.
    fn row_count(&self) -> usize;

    /// Get the number of columns in the worksheet.
    fn column_count(&self) -> usize;

    /// Get the dimensions as (min_row, min_col, max_row, max_col).
    /// Returns None if the worksheet is empty.
    fn dimensions(&self) -> Option<(u32, u32, u32, u32)>;

    /// Get a cell by row and column (1-based indexing).
    fn cell(&self, row: u32, column: u32) -> Result<Box<dyn Cell + '_>>;

    /// Get a cell by coordinate (e.g., "A1").
    fn cell_by_coordinate(&self, coordinate: &str) -> Result<Box<dyn Cell + '_>>;

    /// Get all cells as an iterator.
    fn cells(&self) -> Box<dyn CellIterator<'_> + '_>;

    /// Get all rows as an iterator.
    fn rows(&self) -> Box<dyn RowIterator<'_> + '_>;

    /// Get a specific row by index (0-based).
    fn row(&self, row_idx: usize) -> Result<Vec<CellValue>>;

    /// Get cell value by row and column (1-based indexing).
    fn cell_value(&self, row: u32, column: u32) -> Result<CellValue>;
}

/// Iterator over worksheets in a workbook.
pub trait WorksheetIterator<'a> {
    /// Get the next worksheet.
    fn next(&mut self) -> Option<Result<Box<dyn Worksheet + 'a>>>;
}

/// Represents a workbook (Excel file).
pub trait Workbook {
    /// Get the active worksheet.
    fn active_worksheet(&self) -> Result<Box<dyn Worksheet + '_>>;

    /// Get all worksheet names.
    fn worksheet_names(&self) -> Vec<String>;

    /// Get a worksheet by name.
    fn worksheet_by_name(&self, name: &str) -> Result<Box<dyn Worksheet + '_>>;

    /// Get a worksheet by index.
    fn worksheet_by_index(&self, index: usize) -> Result<Box<dyn Worksheet + '_>>;

    /// Get all worksheets as an iterator.
    fn worksheets(&self) -> Box<dyn WorksheetIterator<'_> + '_>;

    /// Get the number of worksheets.
    fn worksheet_count(&self) -> usize;

    /// Get the index of the active worksheet.
    fn active_sheet_index(&self) -> usize;
}

/// Open a workbook from a file path.
pub fn open_workbook<P: AsRef<std::path::Path>>(path: P) -> Result<Box<dyn Workbook>> {
    let package = crate::ooxml::opc::OpcPackage::open(path)?;
    let workbook = crate::ooxml::xlsx::Workbook::new(package)?;
    Ok(Box::new(workbook))
}

/// Open a workbook from bytes.
pub fn open_workbook_from_bytes(bytes: &[u8]) -> Result<Box<dyn Workbook>> {
    use std::io::Cursor;
    let cursor = Cursor::new(bytes);
    let package = crate::ooxml::opc::OpcPackage::from_reader(cursor)?;
    let workbook = crate::ooxml::xlsx::Workbook::new(package)?;
    Ok(Box::new(workbook))
}

/// Open an XLS workbook from a file path.
pub fn open_xls_workbook<P: AsRef<std::path::Path>>(path: P) -> Result<crate::ole::xls::XlsWorkbook<std::fs::File>> {
    use std::fs::File;
    let file = File::open(path)?;
    let workbook = crate::ole::xls::XlsWorkbook::new(file)?;
    Ok(workbook)
}

/// Open an XLS workbook from bytes.
pub fn open_xls_workbook_from_bytes(bytes: &[u8]) -> Result<crate::ole::xls::XlsWorkbook<std::io::Cursor<&[u8]>>> {
    use std::io::Cursor;
    let cursor = Cursor::new(bytes);
    let workbook = crate::ole::xls::XlsWorkbook::new(cursor)?;
    Ok(workbook)
}

/// Open an XLS workbook as a trait object from a file path.
pub fn open_xls_workbook_dyn<P: AsRef<std::path::Path>>(path: P) -> Result<Box<dyn Workbook>> {
    let workbook = open_xls_workbook(path)?;
    Ok(Box::new(workbook))
}

/// Open an XLS workbook as a trait object from bytes.
pub fn open_xls_workbook_from_bytes_dyn(bytes: &[u8]) -> Result<Box<dyn Workbook>> {
    use std::io::Cursor;
    let cursor = Cursor::new(bytes.to_vec());
    let workbook = crate::ole::xls::XlsWorkbook::new(cursor)?;
    Ok(Box::new(workbook))
}

/// Open an XLSB workbook from a file path.
pub fn open_xlsb_workbook<P: AsRef<std::path::Path>>(path: P) -> Result<crate::ooxml::xlsb::XlsbWorkbook> {
    use std::fs::File;
    let file = File::open(path)?;
    let workbook = crate::ooxml::xlsb::XlsbWorkbook::new(file)?;
    Ok(workbook)
}

/// Open an XLSB workbook from bytes.
pub fn open_xlsb_workbook_from_bytes(bytes: &[u8]) -> Result<crate::ooxml::xlsb::XlsbWorkbook> {
    use std::io::Cursor;
    let cursor = Cursor::new(bytes);
    let workbook = crate::ooxml::xlsb::XlsbWorkbook::new(cursor)?;
    Ok(workbook)
}

/// Open an XLSB workbook as a trait object from a file path.
pub fn open_xlsb_workbook_dyn<P: AsRef<std::path::Path>>(path: P) -> Result<Box<dyn Workbook>> {
    let workbook = open_xlsb_workbook(path)?;
    Ok(Box::new(workbook))
}

/// Open an XLSB workbook as a trait object from bytes.
pub fn open_xlsb_workbook_from_bytes_dyn(bytes: &[u8]) -> Result<Box<dyn Workbook>> {
    use std::io::Cursor;
    let cursor = Cursor::new(bytes.to_vec());
    let workbook = crate::ooxml::xlsb::XlsbWorkbook::new(cursor)?;
    Ok(Box::new(workbook))
}

/// Text-based format support (CSV, TSV, PRN, etc.)
pub mod text;

/// Open a CSV workbook from a file path.
pub fn open_csv_workbook<P: AsRef<std::path::Path>>(path: P) -> Result<Box<dyn Workbook>> {
    let workbook = crate::sheet::text::TextWorkbook::open(path)?;
    Ok(Box::new(workbook))
}

/// Open a CSV workbook from bytes.
pub fn open_csv_workbook_from_bytes(bytes: &[u8]) -> Result<Box<dyn Workbook>> {
    let workbook = crate::sheet::text::TextWorkbook::from_bytes(bytes, crate::sheet::text::TextConfig::default())?;
    Ok(Box::new(workbook))
}

/// Open a TSV workbook from a file path.
pub fn open_tsv_workbook<P: AsRef<std::path::Path>>(path: P) -> Result<Box<dyn Workbook>> {
    let config = crate::sheet::text::TextConfig::tsv();
    let workbook = crate::sheet::text::TextWorkbook::from_path_with_config(path, config)?;
    Ok(Box::new(workbook))
}

/// Open a TSV workbook from bytes.
pub fn open_tsv_workbook_from_bytes(bytes: &[u8]) -> Result<Box<dyn Workbook>> {
    let config = crate::sheet::text::TextConfig::tsv();
    let workbook = crate::sheet::text::TextWorkbook::from_bytes(bytes, config)?;
    Ok(Box::new(workbook))
}

/// Open a PRN workbook from a file path.
pub fn open_prn_workbook<P: AsRef<std::path::Path>>(path: P) -> Result<Box<dyn Workbook>> {
    let config = crate::sheet::text::TextConfig::prn();
    let workbook = crate::sheet::text::TextWorkbook::from_path_with_config(path, config)?;
    Ok(Box::new(workbook))
}

/// Open a PRN workbook from bytes.
pub fn open_prn_workbook_from_bytes(bytes: &[u8]) -> Result<Box<dyn Workbook>> {
    let config = crate::sheet::text::TextConfig::prn();
    let workbook = crate::sheet::text::TextWorkbook::from_bytes(bytes, config)?;
    Ok(Box::new(workbook))
}

/// Open a text workbook with custom configuration from a file path.
pub fn open_text_workbook_with_config<P: AsRef<std::path::Path>>(
    path: P,
    config: crate::sheet::text::TextConfig
) -> Result<Box<dyn Workbook>> {
    let workbook = crate::sheet::text::TextWorkbook::from_path_with_config(path, config)?;
    Ok(Box::new(workbook))
}

/// Open a text workbook with custom configuration from bytes.
pub fn open_text_workbook_from_bytes_with_config(
    bytes: &[u8],
    config: crate::sheet::text::TextConfig
) -> Result<Box<dyn Workbook>> {
    let workbook = crate::sheet::text::TextWorkbook::from_bytes(bytes, config)?;
    Ok(Box::new(workbook))
}
