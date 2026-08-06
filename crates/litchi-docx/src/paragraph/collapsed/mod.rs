//! Word 2012 paragraph-collapse metadata.
//!
//! The extension is intentionally owned by a focused paragraph module. The
//! semantic value is a closed enum, the XML reader/writer is bounded and
//! loss-preserving, and [`Transaction`] publishes a new snapshot only after
//! the candidate paragraph has been validated.

mod codec;
mod model;
mod package;
mod transaction;
mod validation;

#[cfg(test)]
mod tests;

pub use model::Collapsed;
pub use transaction::{Commit, Patch, Snapshot, Transaction};

pub(crate) use codec::append_xml;
