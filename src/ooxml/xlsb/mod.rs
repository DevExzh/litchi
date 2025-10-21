//! Excel XLSB (.xlsb) file format reader
//!
//! This module provides functionality to parse Microsoft Excel XLSB files,
//! which are Excel 2007+ binary format files stored in ZIP containers.
//! The implementation follows the MS-XLSB specification.

/// Error types for XLSB parsing
mod error;

/// XLSB record parsing utilities
mod records;

/// Workbook parsing implementation
mod workbook;

/// Worksheet parsing implementation
mod worksheet;

/// Cell value parsing and representation
mod cell;

/// XLSB cells reader
mod cells_reader;

/// Shared parsing utilities
mod utils;

pub use cell::XlsbCell;
pub use error::{XlsbError, XlsbResult};
pub use workbook::XlsbWorkbook;
pub use worksheet::XlsbWorksheet;
