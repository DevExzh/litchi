//! Text-based spreadsheet format support (CSV, TSV, etc.)
//!
//! This module provides support for delimited text formats like CSV, TSV, and PRN files.
//! It implements the unified sheet API with configurable delimiters and efficient parsing.
//!
//! # Features
//!
//! - **Configurable delimiters**: Support for CSV (comma), TSV (tab), PRN (semicolon), or custom delimiters
//! - **Streaming parsing**: Memory-efficient processing of large files
//! - **Quote handling**: Proper support for quoted fields with escape sequences
//! - **Zero-copy operations**: Minimize allocations where possible
//! - **High performance**: Optimized for speed and memory usage
//!
//! # Example
//!
//! ```rust,no_run
//! use litchi::sheet::text::{TextWorkbook, TextConfig};
//!
//! // Open a CSV file with default settings
//! let workbook = TextWorkbook::open("data.csv")?;
//!
//! // Open with custom delimiter (TSV)
//! let config = TextConfig::new().with_delimiter(b'\t');
//! let workbook = TextWorkbook::from_path_with_config("data.tsv", config)?;
//!
//! // Access worksheets and cells
//! for worksheet in workbook.worksheets() {
//!     let worksheet = worksheet?;
//!     println!("Sheet: {}", worksheet.name());
//!
//!     // Read cells
//!     for row in worksheet.rows() {
//!         let row = row?;
//!         for cell_value in row {
//!             println!("{:?}", cell_value);
//!         }
//!     }
//! }
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```

pub mod workbook;
pub mod worksheet;
pub mod cell;
pub mod iterators;
pub mod parser;

pub use workbook::{TextWorkbook, TextConfig};
pub use worksheet::TextWorksheet;
pub use cell::TextCell;
pub use parser::TextParser;

#[cfg(test)]
mod tests;
