//! Slide-ID-owned PowerPoint Designer tags.
//!
//! [MS-PPTX] 2.2.20 attaches the inert PowerPoint 2020 tag list to a
//! presentation-level `p:sldId`, not to the related slide part. This owner
//! retains duplicate extension entries for inventory, but singular mutation
//! deliberately refuses that ambiguous producer state.

mod codec;
mod model;
mod package;
mod transaction;

#[cfg(test)]
mod tests;

pub use crate::shape::designer::{Limits, Tag, Tags};
pub use model::{Binding, Snapshot};
pub use package::{
    apply_commit, apply_patch, load, load_snapshot, load_snapshot_with_limits, remove, store,
};
pub use transaction::{Commit, Edit, Patch, Revision};
