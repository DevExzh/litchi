//! Guarded source-backed workbook tab visibility and selection publication.
//!
//! This owner is deliberately narrower than [`crate::workbook::Edit`]. It can
//! change visibility and the active tab of existing catalog entries, but it
//! cannot rename, reorder, create, remove, or retarget sheets.

mod patch;
mod snapshot;
pub mod source;

pub(crate) use patch::Patch;
pub(crate) use snapshot::GraphBinding;
pub use snapshot::{Snapshot, Tab};
