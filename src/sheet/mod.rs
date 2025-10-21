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
//! - `.numbers` - Apple Numbers (iWork Archive)
//!
//! # Quick Start
//!
//! ```rust,no_run
//! use litchi::sheet::Workbook;
//!
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
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```
//!
//! # Architecture
//!
//! The module provides both:
//! - **Unified API**: `Workbook` struct for high-level operations
//! - **Trait-based API**: `Workbook`, `Worksheet`, `Cell` traits for advanced use

// Submodule declarations
pub mod functions;
pub mod text;
pub mod traits;
pub mod types;
mod workbook;
mod workbook_types;

// Re-exports
pub use functions::*;
pub use traits::{Cell, CellIterator, RowIterator, WorkbookTrait, Worksheet, WorksheetIterator};
pub use types::{CellValue, Result};
pub use workbook::Workbook;
