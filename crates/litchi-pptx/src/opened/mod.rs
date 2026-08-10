//! Atomic semantic transactions over an opened presentation.
//!
//! One immutable root composes slide/text/notes edits with common, grouped, and
//! endpoint-closed connector shape transfer; picture-part creation/removal;
//! ordinary charts, media, table styles, legacy/modern comments and typed
//! extensions; master/layout authoring; and dependency relationship transfer.
//! Publication uses an exact finite OPC resource patch and classifies shape
//! transfer refusals before mutation.

mod model;
mod patch;
mod transaction;
mod xml;

#[cfg(test)]
mod tests;

pub use model::{Limits, Slide, Snapshot};
pub use patch::{Conflict, History, Patch, Resolution, ThreeWayPlan};
pub use transaction::{Commit, Transaction};

pub(crate) use model::capture;
pub(crate) use patch::apply;
