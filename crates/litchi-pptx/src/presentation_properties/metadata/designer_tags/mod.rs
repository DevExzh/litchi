//! Slide-ID-owned `PowerPoint` Designer tags.
//!
//! [MS-PPTX] 2.2.20 attaches the inert `PowerPoint` 2020 tag list to a
//! presentation-level `p:sldId`, not to the related slide part. This owner
//! retains duplicate extension entries for inventory, but singular mutation
//! deliberately refuses that ambiguous producer state.

mod codec;
mod edit;
mod model;
mod package;
mod transaction;

#[cfg(test)]
mod tests;

pub use crate::shape::designer::{Limits, Tag, Tags};
pub use edit::{Commit, Diagnostics, Edit, Error, LimitKind, Patch};
pub use model::{Binding, Snapshot};
pub use package::{load, load_snapshot};
pub use transaction::Revision;

#[cfg(test)]
pub(crate) use package::apply_patch;
pub(crate) use package::load_snapshot_with_limits;
pub(crate) use package::{remove, store};
pub(crate) use transaction::{Change, Draft, Prepared};
