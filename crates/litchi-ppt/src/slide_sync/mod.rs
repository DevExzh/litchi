//! Layered `PowerPoint` 12 slide-library synchronization owner.
//!
//! The owner keeps the binary record graph and the synchronization vocabulary
//! separate. [`Snapshot`] captures one complete slide record, while
//! [`Edit`] (from [`crate::Package::edit_slide_sync`]) stages one optional
//! `RoundTripSlideSyncInfo12` container of a slide and publishes it as a new
//! package. Records outside that container are retained in
//! their original order and remain opaque to the semantic editor.

mod codec;
mod edit;
mod model;
#[allow(
    dead_code,
    reason = "the record-level owner behind the package-level Edit; its full surface stays exercised by the unit tests (record 0357)"
)]
mod transaction;
mod validation;

#[cfg(test)]
mod tests;

pub use edit::Edit;
pub use model::{
    LibraryUrl, Limits, MAX_TEXT_BYTES, ServerId, Snapshot, Synchronization, SystemTime,
};
pub use transaction::{Change, Revision};

crate::scoped::family!("slide-sync");
