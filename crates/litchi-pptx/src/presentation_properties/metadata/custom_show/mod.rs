//! Custom-show values and `PresentationML` fragment codec.

mod codec;
mod edit;
mod model;
mod package;
mod transaction;
pub use edit::{Commit, Diagnostics, Edit, Error, LimitKind, Patch};
mod wire;

#[cfg(test)]
mod tests;

pub use model::{List, Show};
pub use package::{load, load_snapshot};
pub use transaction::{Draft, Revision, Snapshot};
