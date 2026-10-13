//! Extended guide values and bounded `PresentationML` extension codec.

mod codec;
mod model;
mod package;
mod transaction;

pub use crate::diagnostics::Diagnostics;
pub use model::{Color, ColorKind, Guide, Guides, List, ListKind, Orientation};
pub use package::{apply_commit, apply_patch, load, load_snapshot, remove, store};
pub use transaction::{Commit, Edit, Patch, Revision, Snapshot};

#[cfg(test)]
mod tests;
