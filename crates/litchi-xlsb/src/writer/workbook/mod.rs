//! XLSB workbook writer implementation.
//!
//! This module provides functionality to create complete XLSB files with multiple worksheets,
//! shared strings, styles, and advanced features.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use model::WorkbookWriter;
