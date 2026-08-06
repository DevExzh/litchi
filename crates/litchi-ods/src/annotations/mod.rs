//! Contextual ODS cell-annotation ownership.
//!
//! The semantic value is the shared [`litchi_odf_common::annotation::Annotation`]
//! model.  This owner adds the spreadsheet-specific cell selector, bounded
//! source scanner, exact XML spans, snapshots, reversible patches, and
//! failure-atomic package-facing transactions around it.

mod codec;
mod model;
mod package;
mod transaction;
mod validation;

#[cfg(test)]
mod tests;

pub use litchi_odf_common::annotation::{Annotation, Element, Node};
pub use model::{Cell, Entry, Selector};
pub use transaction::{Commit, Operation, Patch, Snapshot, Transaction, update};

pub(crate) use package::replace_content;
