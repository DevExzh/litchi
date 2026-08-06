//! Source-checked edits for the fixed-width `[MS-OGRAPH]` `Series` record.
//!
//! The parent [`super`] module owns the chart-area record editor.  This
//! sibling owner keeps the `Series` record's six scalar fields together and
//! changes only the existing twelve-byte payload.  No record is inserted,
//! removed, reordered, or re-encoded, so opaque records remain byte-for-byte
//! intact.

mod codec;
mod model;
mod validation;

pub use model::{Change, Commit, Metadata, Patch, Snapshot, Transaction};

#[cfg(test)]
mod tests;
