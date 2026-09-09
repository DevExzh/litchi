//! Source-backed owners for MS-OLEPS non-simple Property Set storages.
//!
//! A non-simple property set is a CFB storage whose direct `CONTENTS` stream
//! carries the PropertySetStream and whose indirect values name direct child
//! streams or storages.  The owner keeps the CFB source and directory catalog
//! lazy and shared; it parses the contents stream only when requested.  A
//! changed publication uses a bounded canonical CFB rewrite and refuses
//! unsupported directory kinds rather than silently dropping producer data.

mod codec;
mod model;
mod transaction;

pub use crate::property_set::model::{VT_STORAGE, VT_STORED_OBJECT, VT_STREAM, VT_STREAMED_OBJECT};
pub use model::{Element, ElementKind, Limits, Snapshot};
pub use transaction::{Commit, Editor, Patch, Revision, update};

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_unrelated,
    reason = "tests exercise fallible malformed CFB and transaction paths"
)]
mod tests;
