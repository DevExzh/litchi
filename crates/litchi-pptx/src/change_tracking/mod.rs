//! Inert `PowerPoint` change-tracking identifiers.
//!
//! `MS-PPTX` 2.2.9 attaches a creation identifier to `p:cSld` and a
//! modification identifier to each shape's application non-visual properties.
//! The identifiers are document data only: this crate does not infer history,
//! authors, clocks, or collaboration behavior from them.

mod codec;
mod edit;
mod model;
mod package;
mod transaction;

#[cfg(test)]
mod tests;

pub use edit::{Commit, Diagnostics, Edit, Error, LimitKind, Patch};
pub use model::{Id, Shape, Snapshot, State};
pub(crate) use transaction::Draft;

pub(crate) use package::{apply_commit_with_limits, load_with_limits};

/// Office 2010 `PowerPoint` `p14` namespace used by both identifier elements.
pub const NAMESPACE: &str = "http://schemas.microsoft.com/office/powerpoint/2010/main";

/// Extension URI for `p14:creationId` below `p:cSld`.
pub const CREATION_EXTENSION_URI: &str = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}";

/// Extension URI for `p14:modId` below a shape's `p:nvPr`.
pub const MODIFICATION_EXTENSION_URI: &str = "{D42A27DB-BD31-4B8C-83A1-F6EECF244321}";
