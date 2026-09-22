//! Semantic facade for the legacy XLS workbook owner.
//!
//! The facade keeps the stable crate::workbook module path while separating
//! the typed workbook model from BIFF substream decoding and OLE package
//! orchestration. The split follows the MS-XLS compound-file -> stream ->
//! substream -> record layering.

mod codec;
mod model;
pub mod package;
mod query_cache;
pub mod source;
mod validation_only;

#[cfg(test)]
mod tests;
#[cfg(test)]
mod validation_only_tests;

pub use model::{OpenOptions, Workbook};
pub use source::{
    SourceBackedCell, SourceBackedError, SourceBackedLimits, SourceBackedWorkbook,
    SourceBackedWorksheet, SourceBackedWorksheetDescriptor,
};
pub(crate) use validation_only::{KeptCells, ValidationWorkbook};
