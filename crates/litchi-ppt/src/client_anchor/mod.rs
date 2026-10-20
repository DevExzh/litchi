//! Source-preserving ownership of `PowerPoint` shape anchors.
//!
//! [`Anchor`] and its rectangle values form the semantic model, and
//! [`Snapshot`] retains one exact MS-PPT `OfficeArtClientAnchor` record. A
//! shape's anchor changes through
//! [`Edit::set_shape_anchor`](crate::package::edit::Edit::set_shape_anchor)
//! on the primary [`crate::Package::edit`]; the record-level edit, the
//! binary codec and the validation policy are private layers (record 0357).

mod codec;
mod model;
#[allow(
    dead_code,
    reason = "the record-level owner the primary Package::edit drives; its full surface stays exercised by the unit tests (record 0357)"
)]
mod transaction;
mod validation;

#[cfg(test)]
mod tests;

pub use codec::RECORD_TYPE;
pub use model::{Anchor, Data, Encoding, Limits, Rect, SmallRect};
#[allow(
    unused_imports,
    reason = "the record-level edit is reached by its unit tests and the slide-order root"
)]
pub(crate) use transaction::{Change, Commit, Patch, Transaction};
pub use transaction::{Revision, Snapshot};
