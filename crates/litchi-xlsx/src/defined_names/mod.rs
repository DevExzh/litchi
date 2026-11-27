//! Source-bound publication of the inert workbook defined-name catalog.

mod patch;
mod snapshot;
pub mod source;

pub(crate) use patch::Patch;
pub use snapshot::Snapshot;
