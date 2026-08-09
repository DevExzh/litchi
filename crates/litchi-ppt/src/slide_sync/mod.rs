//! Layered `PowerPoint` 12 slide-library synchronization owner.
//!
//! The owner keeps the binary record graph and the synchronization vocabulary
//! separate. [`Snapshot`] captures one complete slide record, while
//! [`Editor`] stages one optional `RoundTripSlideSyncInfo12` container and
//! publishes it atomically. Records outside that container are retained in
//! their original order and remain opaque to the semantic editor.

mod codec;
mod model;
mod transaction;
mod validation;

#[cfg(test)]
mod tests;

pub use model::{
    LibraryUrl, Limits, MAX_TEXT_BYTES, ServerId, Snapshot, Synchronization, SystemTime,
};
pub use transaction::{Change, ChangeSet, Commit, Editor, Revision};
