//! Presentation ordering, custom shows, and sections as one validated graph.

mod codec;
mod edit;
mod model;
mod package;
mod transaction;
pub use edit::{Commit, Diagnostics, Edit, Error, LimitKind, Patch};

#[cfg(test)]
mod tests;

pub use super::Limits;
pub use model::{Graph, Reference};
pub use package::{find_custom_show, find_section, load, load_snapshot};
pub use transaction::{Draft, Revision, Snapshot};
