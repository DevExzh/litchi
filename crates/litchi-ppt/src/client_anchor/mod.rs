//! Source-preserving ownership of `PowerPoint` shape anchors.
//!
//! [`Anchor`] and its rectangle values form the semantic model, [`Snapshot`]
//! retains one exact MS-PPT `OfficeArtClientAnchor` record, and
//! [`Transaction`] publishes isolated edits with a reversible, source-checked
//! [`Patch`]. The binary codec and validation policy remain private layers.

mod codec;
mod model;
mod transaction;
mod validation;

#[cfg(test)]
mod tests;

pub use codec::RECORD_TYPE;
pub use model::{Anchor, Data, Encoding, Limits, Rect, SmallRect};
pub use transaction::{Change, Commit, Patch, Revision, Snapshot, Transaction};
