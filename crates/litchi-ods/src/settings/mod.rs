//! Contextual ODS calculation-settings ownership.
//!
//! This module is intentionally self-contained so the ODS crate root can
//! expose it when the package facade is ready.  The schema model and XML
//! codec remain in `litchi-odf-common`; this family layer owns only the
//! `content.xml` context, immutable snapshots, and atomic edits.

mod codec;
mod model;
mod transaction;

#[cfg(test)]
mod tests;

pub use litchi_odf_common::calculation::{Iteration, IterationStatus, NullDate, Settings};
pub use model::Snapshot;
pub use transaction::{Commit, Editor, Transaction};
