//! Changes Information values, bounded XML, and OPC lifecycle.

mod codec;
mod model;
mod package;
mod transaction;

pub use codec::{CONTENT_TYPE, RELATIONSHIP_TYPE};
pub use model::{Data, Descriptor, Info, Kind, List, Namespace, Part};
pub use package::{apply_commit, apply_patch, load, load_snapshot, remove, store};
pub use transaction::{Commit, Patch, Revision, Snapshot, Transaction};

#[cfg(test)]
mod tests;
