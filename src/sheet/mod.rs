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
///     let worksheet = worksheet?;
///     println!("Sheet: {}", worksheet.name());
///
///     // Access cells
///     let cell = worksheet.cell(1, 1)?;
///     println!("A1 value: {:?}", cell.value());
///
///     // Access by coordinate
///     let cell_a1 = worksheet.cell_by_coordinate("A1")?;
///     println!("A1 coordinate access: {:?}", cell_a1.value());
///
///     // Iterate over all cells
///     for cell_result in worksheet.cells() {
///         let cell = cell_result?;
///         println!("Cell {}: {:?}", cell.coordinate(), cell.value());
///     }
/// }
///
/// // Access specific worksheet by name
/// let sheet1 = workbook.worksheet_by_name("Sheet1")?;
/// println!("Found worksheet: {}", sheet1.name());
///
/// // Get active worksheet
/// let active = workbook.active_worksheet()?;
/// println!("Active sheet: {}", active.name());
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```

// Submodule declarations
pub mod types;
pub mod traits;
pub mod functions;
pub mod text;

// Re-exports
pub use types::{Result, CellValue};
pub use traits::{Cell, CellIterator, RowIterator, Worksheet, WorksheetIterator, Workbook};
pub use functions::*;
