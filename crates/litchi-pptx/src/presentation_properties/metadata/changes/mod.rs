//! Changes Information values, bounded XML, and OPC lifecycle.

mod codec;
mod edit;
mod model;
mod package;
mod transaction;
pub use edit::{Commit, Diagnostics, Edit, EditOptions, Error, LimitKind, Patch};

pub use codec::{CONTENT_TYPE, Limits, RELATIONSHIP_TYPE};
pub use model::{Data, Descriptor, Info, Kind, List, Namespace, Part};
pub use package::{load, load_snapshot};
pub use transaction::{Draft, Revision, Snapshot};

#[cfg(test)]
mod tests;
