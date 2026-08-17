//! Source-backed visibility edits for existing worksheet row owners.
//!
//! This capability owns only the direct unqualified `hidden` attribute on an
//! already stored `<row>` element. Hiding writes the canonical `hidden="1"`
//! form and unhiding removes the attribute. Row contents, layout, styles, the
//! containing worksheet structure, and every other package member are retained
//! exactly. The conservative scalar-cell source closure is reused, so formulas,
//! relationships, protection, markup compatibility, macros, and ambiguous
//! package graphs are refused rather than partially interpreted.

mod patch;
pub(crate) mod rewrite;
mod snapshot;
mod source;

pub use patch::{Commit, Diagnostics, Patch};
pub use snapshot::Snapshot;
pub use source::{MAX_BATCH_EDITS, RowVisibilityEdit, SourceBackedEditor, SourceEdit};
