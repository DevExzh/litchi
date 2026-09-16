//! Source-backed visibility edits for existing worksheet row owners.
//!
//! This capability owns only the direct unqualified `hidden` attribute on an
//! already stored `<row>` element. Hiding writes the canonical `hidden="1"`
//! form and unhiding removes the attribute. Row contents, layout, styles, the
//! containing worksheet structure, and every other package member are retained
//! exactly. The conservative scalar-cell source closure is reused, so formulas,
//! macros and ambiguous package graphs are refused rather than partially
//! interpreted.
//!
//! Change 0657 replaced that closure's allow-lists with a dependency rule, and
//! this module's admission follows it on the same terms, asking of each
//! construct whether its meaning depends on *which rows are hidden*:
//!
//! * a **worksheet relationship** is admitted. The rewrite copies every byte
//!   but the `<row>` tags it owns, and a hyperlink, drawing, comment or
//!   printer-settings part is anchored by address, so none of them moves when
//!   a row is hidden. A table or query table stays refused by the shared
//!   closure, because its meaning depends on a cell's value.
//! * a **sheet protection** forbids the edit outright, and an **autofilter**,
//!   a **sort state** or a **custom sheet view** carries its own hidden-row
//!   state in exactly the attribute this module owns. All four are refused
//!   here, by this module's own scan, because the shared closure no longer
//!   refuses them on its behalf.
//! * **markup-compatibility content** is refused because the cell store this
//!   module reuses is built from the preprocessed bytes while its rewrite is
//!   lexical over the source bytes, and the two views must be the same bytes.

mod patch;
pub(crate) mod rewrite;
mod snapshot;
mod source;

pub use patch::{Commit, Diagnostics, Patch};
pub use snapshot::Snapshot;
pub use source::{MAX_BATCH_EDITS, RowVisibilityEdit, SourceBackedEditor, SourceEdit};
