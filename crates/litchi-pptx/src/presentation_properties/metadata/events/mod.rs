//! Slide-show event values and bounded extension storage.

mod codec;
mod edit;
mod model;
mod package;
mod transaction;
pub use edit::{Commit, Diagnostics, Edit, EditOptions, Error, LimitKind, Patch};
mod validation;

#[cfg(test)]
mod tests;

pub(crate) use codec::validate_slide_xml;
pub use codec::{EXTENSION_URI, Limits};
pub use model::{Draft, Event, Kind, Trigger};
pub use package::{load, load_snapshot};
pub use transaction::{ListDraft, Revision, Snapshot};
