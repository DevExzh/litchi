//! Excel (.xlsx) spreadsheet support.
//!
//! This module provides parsing and manipulation of Microsoft Excel spreadsheets
//! in the Office Open XML (OOXML) format (.xlsx files).
//!
//! # Architecture
//!
//! The module follows a similar pattern to other OOXML modules:
//!
//! - `Workbook`: The main workbook content and API
//! - `Worksheet`: Individual sheet content and data access
//! - Various internal parsers for styles, shared strings, etc.
//!
//! # Example
//!
//! ```rust,no_run
//! use litchi::ooxml::xlsx::Workbook;
//!
//! // Open a workbook
//! let workbook = Workbook::open("workbook.xlsx")?;
//!
//! // Access worksheets
//! for worksheet in workbook.worksheets() {
//!     println!("Sheet: {}", worksheet.name());
//!
//!     // Access cells
//!     let cell = worksheet.cell(1, 1)?;
//!     println!("A1 value: {:?}", cell.value());
//! }
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```

pub mod cell;
pub mod parsers;
pub mod shared_strings;
pub mod styles;
pub mod workbook;
pub mod worksheet;

// Re-export main types for convenience
pub use cell::Cell;
pub use shared_strings::SharedStrings;
pub use styles::Styles;
pub use workbook::Workbook;
pub use worksheet::Worksheet;
