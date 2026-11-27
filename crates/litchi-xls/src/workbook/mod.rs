//! Semantic facade for the legacy XLS workbook owner.
//!
//! The facade keeps the stable crate::workbook module path while separating
//! the typed workbook model from BIFF substream decoding and OLE package
//! orchestration. The split follows the MS-XLS compound-file -> stream ->
//! substream -> record layering.

mod codec;
pub(crate) mod model;
pub mod package;
mod query_cache;
pub mod source;
mod validation_only;

#[cfg(test)]
mod tests;
#[cfg(test)]
mod validation_only_tests;

pub use crate::LimitKind;
pub use crate::source_limits::{SourceLimits, SourceLimitsBuilder};
pub use model::{OpenOptions, Worksheets};
pub use source::{
    CellCursor, SourceBackedCell, SourceBackedError, SourceBackedLimits, SourceBackedOptions,
    SourceBackedWorkbook, SourceBackedWorksheet, SourceBackedWorksheetDescriptor, StyleTables,
};
pub(crate) use validation_only::{KeptCells, ValidationWorkbook};
