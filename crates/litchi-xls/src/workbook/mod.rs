//! Semantic facade for the legacy XLS workbook owner.
//!
//! The facade keeps the stable crate::workbook module path while separating
//! the typed workbook model from BIFF substream decoding and OLE package
//! orchestration. The split follows the MS-XLS compound-file -> stream ->
//! substream -> record layering.

mod codec;
mod model;
pub mod package;

#[cfg(test)]
mod tests;

pub use model::{OpenOptions, Workbook};
