//! Workbook implementation for XLSB files.
//!
//! The public [`Workbook`] facade is backed by a typed model, BIFF12 codecs,
//! and OPC/package integration kept in separate layers.

pub(crate) mod cache;
pub(crate) mod catalog;
mod codec;
mod model;
mod package;
mod source;

#[cfg(test)]
mod tests;

pub use cache::Stats;
pub(crate) use model::DrawingLoadPolicy;
pub use model::{Workbook, WorksheetIterator};
pub use source::{
    CellCursor, SourceBackedExternalLink, SourceBackedIndexedWorksheet, SourceBackedWorkbook,
    SourceBackedWorksheet,
};
