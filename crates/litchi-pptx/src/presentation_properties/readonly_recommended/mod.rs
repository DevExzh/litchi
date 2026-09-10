//! Source-preserving PresentationML 2017/10 read-only recommendation.
//!
//! The recommendation is a document hint stored in the presentation
//! properties `p:extLst`. It is intentionally inert: this owner never
//! enforces a host edit policy.

mod codec;
mod package;
mod transaction;

#[cfg(test)]
mod tests;

pub use package::{apply_commit, apply_patch, load, load_snapshot, put, remove};
pub use transaction::{Commit, Patch, Revision, Snapshot, Transaction};

/// The Microsoft extension URI that owns `p1710:readonlyRecommended`.
pub const EXTENSION_URI: &str = "{1BD7E111-0CB8-44D6-8891-C1BB2F81B7CC}";

/// The namespace URI used by the `readonlyRecommended` element.
pub const NAMESPACE: &str = "http://schemas.microsoft.com/office/powerpoint/2017/10/main";
