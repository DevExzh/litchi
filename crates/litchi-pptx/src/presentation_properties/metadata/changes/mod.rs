//! Changes Information values, bounded XML, and OPC lifecycle.

mod codec;
mod model;
mod package;
mod transaction;

pub use crate::diagnostics::Diagnostics;
pub use codec::{CONTENT_TYPE, RELATIONSHIP_TYPE};
pub use model::{Data, Descriptor, Info, Kind, List, Namespace, Part};
pub use package::{apply_commit, apply_patch, load, load_snapshot, remove, store};
pub use transaction::{Commit, Edit, Patch, Revision, Snapshot};

#[cfg(test)]
mod tests;
