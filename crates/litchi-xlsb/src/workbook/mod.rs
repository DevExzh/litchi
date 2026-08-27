//! Workbook implementation for XLSB files.
//!
//! The public [`Workbook`] facade is backed by a typed model, BIFF12 codecs,
//! and OPC/package integration kept in separate layers.

mod codec;
mod model;
mod package;
mod source;

#[cfg(test)]
mod tests;

pub use model::{Workbook, WorksheetIterator};
pub use source::{SourceBackedExternalLink, SourceBackedWorkbook, SourceBackedWorksheet};
