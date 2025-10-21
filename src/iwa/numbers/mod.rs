//! Numbers Spreadsheet Support
//!
//! This module provides comprehensive support for parsing Apple Numbers spreadsheets,
//! including table extraction, cell data parsing, and formula support.
//!
//! ## Features
//!
//! - Sheet extraction
//! - Table parsing with cell data
//! - Formula extraction
//! - CSV export
//! - Cell formatting information
//!
//! ## Example
//!
//! ```rust,no_run
//! use litchi::iwa::numbers::NumbersDocument;
//!
//! let doc = NumbersDocument::open("spreadsheet.numbers")?;
//! let sheets = doc.sheets()?;
//!
//! for sheet in sheets {
//!     println!("Sheet: {}", sheet.name);
//!     for table in &sheet.tables {
//!         println!("  Table: {}", table.name);
//!         println!("{}", table.to_csv());
//!     }
//! }
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```

pub mod cell;
pub mod document;
pub mod sheet;
pub mod table;
pub mod table_extractor;

pub use cell::{CellType, CellValue};
pub use document::NumbersDocument;
pub use sheet::NumbersSheet;
pub use table::NumbersTable;
pub use table_extractor::TableDataExtractor;
