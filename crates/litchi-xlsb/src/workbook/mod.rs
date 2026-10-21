//! Workbook implementation for XLSB files.
//!
//! The public [`Workbook`] facade is backed by a typed model, BIFF12 codecs,
//! and OPC/package integration kept in separate layers.

pub(crate) mod cache;
pub(crate) mod catalog;
mod codec;
mod features;
mod model;
mod package;
mod source;

#[cfg(test)]
mod tests;

pub use cache::Stats;
pub(crate) use model::DrawingLoadPolicy;
pub use model::{Workbook, Worksheets};
pub use source::{
    CellCursor, SourceBackedExternalLink, SourceBackedIndexedWorksheet, SourceBackedWorkbook,
    SourceBackedWorksheet,
};

/// A bounded dimension of this family; an exceeded bound is reported as
/// [`crate::package::PackageError::Limit`] with this kind.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Records of one worksheet indexed by a source-backed owner.
    IndexedRecords,
    /// UTF-16 units of the indexed shared strings.
    IndexedStringUnits,
    /// Bytes of one indexed worksheet part.
    IndexedPartBytes,
}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(match self {
            Self::IndexedRecords => "indexed worksheet records",
            Self::IndexedStringUnits => "indexed shared-string code units",
            Self::IndexedPartBytes => "indexed worksheet bytes",
        })
    }
}
