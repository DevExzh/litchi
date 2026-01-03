//! Legacy Excel (.xls) file format reader
//!
//! This module provides functionality to parse Microsoft Excel files
//! in the legacy binary format (.xls files), which are OLE2-based files.
//! The implementation is based on the BIFF (Binary Interchange File Format)
//! specification and draws inspiration from other spreadsheet libraries.

/// Error types for XLS parsing
mod error;

/// BIFF record parsing utilities
pub mod records;

/// Workbook parsing implementation
mod workbook;

/// Worksheet parsing implementation
mod worksheet;

/// Cell value parsing and representation
mod cell;

/// Shared parsing utilities
mod utils;

/// XLS file writing
pub mod writer;

pub use cell::XlsCell;
pub use error::{XlsError, XlsResult};
pub use workbook::XlsWorkbook;
pub use worksheet::XlsWorksheet;
pub use writer::XlsWriter;
