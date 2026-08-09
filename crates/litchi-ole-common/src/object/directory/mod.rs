//! Typed CFB directory metadata shared by embedded-object owners.
//!
//! The CFB container already validates its directory tree.  This module adds
//! the small, format-neutral projection that DOC, PPT, and XLS hosts need
//! when deciding whether a captured storage or stream is the object they own.
//! It carries directory identity and containment links without exposing the
//! container's mutable internals or activating any OLE payload.

mod catalog;
mod codec;
mod model;
mod patch;
mod snapshot;
mod transaction;
mod validation;

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_unrelated,
    clippy::default_trait_access,
    reason = "tests use concise assertions while exercising fallible malformed-input paths"
)]
mod tests;

pub use catalog::{Catalog, Entry};
pub use model::{EntryKind, Limits, Links, MAX_REGULAR_SID, Metadata, NOSTREAM, Sid};
pub use patch::{Change, Patch};
pub use snapshot::Snapshot;
pub use transaction::{Commit, Revision, Transaction, update};

pub(crate) use codec::{decode, format_class_id};
