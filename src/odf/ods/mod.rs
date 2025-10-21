//! OpenDocument Spreadsheet (.ods) implementation.
//!
//! This module provides comprehensive support for parsing and working with
//! OpenDocument Spreadsheet documents (.ods files), which are the open standard
//! equivalent of Microsoft Excel spreadsheets.

mod cell;
mod parser;
mod row;
mod sheet;
mod spreadsheet;

pub use cell::{Cell, CellValue};
pub use row::Row;
pub use sheet::Sheet;
pub use spreadsheet::Spreadsheet;
