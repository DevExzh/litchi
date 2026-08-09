//! Atomic semantic transactions over an opened presentation.
//!
//! This owner deliberately edits only existing slide-order entries, existing
//! shape text runs, and existing notes text runs. Every other OPC part,
//! relationship, and unknown XML byte remains outside the write set.

mod model;
mod patch;
mod transaction;
mod xml;

#[cfg(test)]
mod tests;

pub use model::{Limits, Slide, Snapshot};
pub use patch::{History, Patch};
pub use transaction::{Commit, Transaction};

pub(crate) use model::capture;
pub(crate) use patch::apply;
