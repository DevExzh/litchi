//! Typed Word caption metadata.
//!
//! The owner covers the two extended string tables addressed by the Word 97+
//! FIB: `SttbfCaption` and `SttbfAutoCaption`. Caption labels and OLE `ProgID`
//! references are inspected, edited, and serialized through immutable package
//! snapshots; no caption is inserted, no field is refreshed, and no host or
//! macro is opened or executed.

mod codec;
mod family;
mod model;
mod package;
mod semantic;
mod transaction;
mod validation;

#[cfg(test)]
mod tests;

/// Shared MS-OSHARED numbering vocabulary used by caption numbers.
pub use crate::parts::numbering::NumberFormat as Format;

pub use family::{Commit, Diagnostics, Edit, Error, LimitKind, Patch, Result};
pub use model::{
    AutoEntry, AutoTable, Definition, Heading, Info, LabelTable, Location, Numbering, Separator,
};
pub use semantic::Tables;

/// The part-level caption codec: the semantic [`record::Snapshot`] and its
/// [`record::Edit`] staging, which publishes a new semantic snapshot;
/// package edits go through [`Edit`].
pub mod record {
    pub use super::transaction::{Commit, Diagnostics, Edit, Error, Patch, Snapshot};
}

// The document state predates the contextual facade and uses this name
// internally. It is crate-visible only; the public path is `captions::Tables`.
pub(crate) use semantic::Tables as CaptionTables;
