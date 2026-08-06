//! Typed, bounded `MsoEnvelopeCLSID` metadata from a DOC table stream.

mod codec;
mod model;
mod package;
mod transaction;
mod validation;

pub use model::{
    Attachment, Envelope, FollowUpStatus, Importance, MSO_ENVELOPE_CLSID, Message, Payload,
    PropertyValue, RecipientCollection, RecipientProperties, RecipientProperty, SecurityFlags,
    Sensitivity, Text, Version,
};
pub use package::{Commit as PackageCommit, Editor, Snapshot as PackageSnapshot};
pub use transaction::{Commit, Error as TransactionError, Patch, Snapshot, Transaction};

/// FIB index of `fcMsoEnvelope`/`lcbMsoEnvelope` in `FibRgFcLcb2000`.
pub const FIB_INDEX: usize = validation::FIB_INDEX;

#[cfg(test)]
mod tests;
