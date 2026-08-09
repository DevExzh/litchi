//! Shape-owned `PowerPoint` classification metadata.
//!
//! [MS-PPTX] stores the classification outcome in the direct non-visual
//! properties of a shape:
//!
//! ```text
//! p:{sp,pic,cxnSp,graphicFrame,grpSp}
//!   / p:nv*Pr / p:nvPr / p:extLst / p:ext[@uri=classification]
//!     / p184:classification[@val]
//! ```
//!
//! The owner is deliberately below [`crate::shape`]. A classification is a
//! property of one selected shape, not a slide-wide setting. Package graph
//! traversal and atomic publication live in [`package`], XML ranges live in
//! [`codec`], bounded value checks live in [`validation`], and staged edits
//! live in [`transaction`].

mod codec;
mod model;
mod package;
mod transaction;
mod validation;

#[cfg(test)]
mod tests;

pub use model::{Opaque, Outcome, Snapshot};
pub use transaction::Editor;

pub(crate) use package::{load, put, remove};

/// URI identifying the `PowerPoint` classification extension.
pub const EXTENSION_URI: &str = "{1162E1C5-73C7-4A58-AE30-91384D911F3F}";

/// Namespace introduced by the classification extension.
pub const NAMESPACE: &str = "http://schemas.microsoft.com/office/powerpoint/2018/4/main";
