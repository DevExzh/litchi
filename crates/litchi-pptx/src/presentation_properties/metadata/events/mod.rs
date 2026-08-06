//! Slide-show event values and bounded extension storage.

mod codec;
mod model;
mod package;
mod transaction;
mod validation;

#[cfg(test)]
mod tests;

pub use codec::EXTENSION_URI;
pub use model::{Draft, Event, Kind, Trigger};
pub use package::{apply_commit, apply_patch, load, load_snapshot, remove, store};
pub use transaction::{Commit, Patch, Revision, Snapshot, Transaction};
