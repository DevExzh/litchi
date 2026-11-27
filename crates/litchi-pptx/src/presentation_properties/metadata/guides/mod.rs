//! Extended guide values and bounded `PresentationML` extension codec.

mod codec;
mod edit;
mod model;
mod package;
mod transaction;
pub use edit::{Commit, Diagnostics, Edit, Error, LimitKind, Patch};

pub use super::Limits;
pub use model::{Color, ColorKind, Guide, Guides, List, ListKind, Orientation};
pub use package::{load, load_snapshot};
pub use transaction::{Draft, Revision, Snapshot};

#[cfg(test)]
mod tests;
