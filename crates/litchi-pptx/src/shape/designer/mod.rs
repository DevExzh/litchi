//! Shape-owned PowerPoint Designer metadata.
//!
//! This bounded owner currently covers the narrowest safely authorable member
//! of the Designer family: the `p15:designElem` boolean from [MS-PPTX] 2.2.17
//! and 2.5. The owner is attached to a selected shape's direct `p:nvPr` and
//! edits only the recognized extension range; unrelated extension entries and
//! bytes remain opaque.

mod codec;
mod model;
mod package;
mod transaction;
mod validation;

#[cfg(test)]
mod tests;

pub use model::{Opaque, Snapshot};
pub use transaction::Editor;

pub(crate) use package::{load, put, remove};

/// URI assigned to the design-element extension by [MS-PPTX] 2.2.17.
pub const EXTENSION_URI: &str = "{386F3935-93C4-4BCD-93E2-E3B085C9AB24}";

/// Namespace assigned to `p15:designElem` by [MS-PPTX] 2.5.
pub const NAMESPACE: &str = "http://schemas.microsoft.com/office/powerpoint/2015/main";
