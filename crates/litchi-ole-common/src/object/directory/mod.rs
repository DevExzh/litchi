//! Typed CFB directory metadata shared by embedded-object owners.
//!
//! The CFB container already validates its directory tree.  This module adds
//! the small, format-neutral projection that DOC, PPT, and XLS hosts need
//! when deciding whether a captured storage or stream is the object they own.
//! It carries directory identity and containment links without exposing the
//! container's mutable internals or activating any OLE payload.

mod codec;
mod model;
mod validation;

#[cfg(test)]
mod tests;

pub use model::{EntryKind, Links, MAX_REGULAR_SID, Metadata, NOSTREAM, Sid};

pub(crate) use codec::{decode, format_class_id};
