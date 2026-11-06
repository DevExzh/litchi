//! Word 2012 footnote-column layout attached to a `w:sectPr`.
//!
//! The extension is deliberately scoped below [`crate::section`]. Its model
//! is package-neutral, its codec owns only the bounded `sectPr` XML seam, the
//! package adapter discovers section properties in a document part, and the
//! transaction layer publishes lossless snapshots and preconditioned patches.

mod codec;
mod model;
mod package;
#[allow(
    dead_code,
    unreachable_pub,
    reason = "the detached commit and patch are exercised by unit tests; Section publishes snapshots"
)]
mod transaction;
mod validation;

#[cfg(test)]
mod tests;

pub use model::Layout;
pub(crate) use package::active_sections;
pub use package::parse_part;
pub use transaction::Snapshot;
