//! Word 2012 footnote-column layout attached to a `w:sectPr`.
//!
//! The extension is deliberately scoped below [`crate::section`]. Its model
//! is package-neutral, its codec owns only the bounded `sectPr` XML seam, the
//! package adapter discovers section properties in a document part, and the
//! transaction layer publishes lossless snapshots and preconditioned patches.

mod codec;
mod model;
mod package;
mod transaction;
mod validation;

#[cfg(test)]
mod tests;

pub use model::Layout;
pub use package::parse_part;
pub use transaction::{Commit, Patch, Snapshot, Transaction};
