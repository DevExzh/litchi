//! Typed, inert worksheet hyperlink projections.
//!
//! Hyperlink targets are metadata only. Parsing never resolves, opens, or
//! fetches an external target.

pub(crate) mod codec;
pub(crate) mod model;
mod patch;
mod snapshot;
mod source;

pub use model::{Hyperlink, HyperlinkReference};
pub use patch::{Commit, Diagnostics, Patch};
pub use snapshot::Snapshot;
pub use source::{SourceBackedEditor, SourceEdit};

/// Focused source-backed editor for one worksheet's direct hyperlinks.
pub type SourceBackedHyperlinkEditor = SourceBackedEditor;
/// Isolated source-backed edit for one worksheet's direct hyperlinks.
pub type SourceBackedHyperlinkEdit = SourceEdit;
/// Immutable source-backed worksheet hyperlink state.
pub type SourceBackedHyperlinkSnapshot = Snapshot;
/// Exact reversible source-backed worksheet hyperlink patch.
pub type SourceBackedHyperlinkPatch = Patch;
/// Successful source-backed worksheet hyperlink commit.
pub type SourceBackedHyperlinkCommit = Commit;
/// Immutable source-backed worksheet hyperlink state.
pub type HyperlinkSnapshot = Snapshot;
/// Exact reversible source-backed worksheet hyperlink patch.
pub type HyperlinkPatch = Patch;
/// Successful source-backed worksheet hyperlink commit.
pub type HyperlinkCommit = Commit;

pub(crate) use codec::parse;
