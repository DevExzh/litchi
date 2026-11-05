//! Typed, bounded `MsoEnvelopeCLSID` metadata from a DOC table stream.

mod codec;
mod family;
mod model;
mod package;
mod transaction;
mod validation;

pub use family::{Commit, Diagnostics, Edit, Error, LimitKind, Patch, Result};
pub use model::{
    Attachment, Envelope, FollowUpStatus, Importance, MSO_ENVELOPE_CLSID, Message, Payload,
    PropertyValue, RecipientCollection, RecipientProperties, RecipientProperty, SecurityFlags,
    Sensitivity, Text, Version,
};

/// The part-level envelope codec: the semantic [`record::Snapshot`] of one
/// `fcMsoEnvelope` range and its [`record::Edit`] staging, which publishes a
/// new semantic snapshot; package edits go through [`Edit`].
pub mod record {
    pub use super::transaction::{Commit, Diagnostics, Edit, Error, Patch, Snapshot};
}

/// FIB index of `fcMsoEnvelope`/`lcbMsoEnvelope` in `FibRgFcLcb2000`.
pub const FIB_INDEX: usize = validation::FIB_INDEX;

#[cfg(test)]
mod tests;
