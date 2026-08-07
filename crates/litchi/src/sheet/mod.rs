//! Unified spreadsheet API for Excel and Numbers files.
//!
//! This module provides a unified interface for working with spreadsheets,
//! supporting multiple formats with automatic detection.
//!
//! # Supported Formats
//!
//! - `.xls` - Microsoft Excel 97-2003 (OLE2)
//! - `.xlsx` - Microsoft Excel 2007+ (Office Open XML)
//! - `.xlsb` - Microsoft Excel Binary Workbook
//! - `.ods` - OpenDocument Spreadsheet
//! - `.numbers` - Apple Numbers (iWork Archive)
//!
//! # Quick Start
//!
//! ```rust,ignore
//! use litchi::sheet::Workbook;
//!
//! # fn main() -> Result<(), Box<dyn std::error::Error + Send + Sync>> {
//! // Open any spreadsheet format - auto-detected
//! let workbook = Workbook::open("data.numbers")?;
//!
//! // Get worksheet names
//! let names = workbook.worksheet_names()?;
//! println!("Worksheets: {:?}", names);
//!
//! // Extract all text
//! let text = workbook.text()?;
//! println!("{}", text);
//!
//! // Get metadata
//! let metadata = workbook.metadata()?;
//! if let Some(title) = metadata.title {
//!     println!("Title: {}", title);
//! }
//! # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
//! # }
//! ```
//!
//! # Architecture
//!
//! The module provides both:
//! - **Unified API**: `Workbook` struct for high-level operations
//! - **Trait-based API**: `Workbook`, `Worksheet`, `Cell` traits for advanced use

/// Canonical worksheet-view semantics shared by XLSX and XLSB.
pub mod view {
    pub use litchi_sheet::view::*;
    pub use litchi_sheet::{Cell, Rect};
}

pub use litchi_core::sheet::{traits, types};

// Unified workbook facade implementations require a concrete workbook format.
#[cfg(feature = "eval")]
pub mod eval {
    pub use litchi_eval::*;
}
#[cfg(any(
    feature = "xls",
    feature = "xlsx",
    feature = "xlsb",
    feature = "ods",
    feature = "iwork"
))]
mod adapters;
#[cfg(any(
    feature = "xls",
    feature = "xlsx",
    feature = "xlsb",
    feature = "ods",
    feature = "iwork"
))]
pub mod functions;
pub mod text;
#[cfg(any(
    feature = "xls",
    feature = "xlsx",
    feature = "xlsb",
    feature = "ods",
    feature = "iwork"
))]
mod workbook;
#[cfg(any(
    feature = "xls",
    feature = "xlsx",
    feature = "xlsb",
    feature = "ods",
    feature = "iwork"
))]
mod workbook_types;

// Re-exports
#[cfg(feature = "eval")]
pub use eval::FormulaEvaluator;
#[cfg(any(
    feature = "xls",
    feature = "xlsx",
    feature = "xlsb",
    feature = "ods",
    feature = "iwork"
))]
pub use functions::*;
pub use traits::{Cell, CellIterator, RowIterator, WorkbookTrait, Worksheet, WorksheetIterator};
pub use types::{CellValue, Result};
#[cfg(any(
    feature = "xls",
    feature = "xlsx",
    feature = "xlsb",
    feature = "ods",
    feature = "iwork"
))]
pub use workbook::Workbook;
