//! Typed, inert worksheet page-break ownership.
//!
//! The owner validates the core `rowBreaks` and `colBreaks` vocabulary,
//! applies Office's collection limits, and supports byte-preserving worksheet
//! rewrites. It does not paginate, print, or infer automatic breaks.

mod codec;
mod model;
mod package;
mod patch;
mod snapshot;
mod transaction;

pub use codec::{parse, replace, write};
pub use model::{Axis, Break, Collection, MAX_HORIZONTAL_BREAKS, MAX_VERTICAL_BREAKS, PageBreaks};
pub use package::{apply_patch, edit, load};
pub use patch::{Commit, Diagnostics, Patch};
pub use snapshot::Snapshot;
pub use transaction::Transaction;

#[cfg(test)]
mod tests;
