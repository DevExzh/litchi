//! Layered ownership of the inert MS-DOC `SttbfAtnBkmk` table.
//!
//! The table stores one fixed `ATNBE` record for each annotation bookmark.
//! This owner exposes only the opaque tag identity and never resolves comment
//! text, bookmark ranges, or external data. Package edits append a new table
//! payload and update only the FIB range; semantic no-ops return the original
//! CFB bytes unchanged.

mod codec;
mod family;
mod model;
mod package;
mod transaction;
mod validation;

#[cfg(test)]
mod tests;

pub use codec::{parse, parse_bytes, to_bytes};
pub use family::{Commit, Diagnostics, Edit, Error, LimitKind, Patch, Result};
pub use model::{Tag, TagId, Tags};

/// The part-level annotation-bookmark codec: the semantic [`record::Snapshot`] and its
/// [`record::Edit`] staging, which publishes a new semantic snapshot;
/// package edits go through [`Edit`].
pub mod record {
    pub use super::transaction::{Commit, Diagnostics, Edit, Error, Patch, Snapshot};
}

/// FIB index of `fcSttbfAtnBkmk`/`lcbSttbfAtnBkmk`.
pub const FIB_INDEX: usize = validation::FIB_INDEX;

/// Maximum annotation-bookmark count permitted by MS-DOC §2.9.277.
pub(crate) const MAX_ENTRIES: usize = validation::MAX_ENTRIES;

/// Maximum bounded payload retained by the document-level deferred reader.
pub(crate) const MAX_TABLE_BYTES: usize = validation::MAX_TABLE_BYTES;
