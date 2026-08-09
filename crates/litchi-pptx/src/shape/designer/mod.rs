//! Shape-owned `PowerPoint` Designer metadata.
//!
//! This bounded owner currently covers the narrowest safely authorable member
//! of the Designer family: the `p15:designElem` boolean from [MS-PPTX] 2.2.17
//! and 2.5. The owner is attached to a selected shape's direct `p:nvPr` and
//! edits only the recognized extension range; unrelated extension entries and
//! bytes remain opaque.

mod codec;
mod model;
#[allow(
    dead_code,
    reason = "detached readers are also used by crate-level round-trip tests"
)]
mod p202;
mod package;
mod properties;
mod transaction;
mod validation;

#[cfg(test)]
mod tests;

pub use model::{DrawingProperties, Limits, Opaque, Snapshot, Tag, Tags};
pub use properties::{PropertiesCommit, PropertiesEdit, PropertiesSnapshot};
pub use transaction::Editor;

pub(crate) use p202::{read_tags_with_prefix, write_properties, write_tags};

#[cfg(test)]
pub(crate) use p202::{read_properties, read_tags};
pub(crate) use package::{load, put, remove};
pub(crate) use properties::load_properties_with_limits;
#[allow(
    unused_imports,
    reason = "staged crate-private host API; public facades deliberately remain unchanged"
)]
// Staged crate-private host API; public facades deliberately remain unchanged.
pub(crate) use properties::{apply_properties, load_properties, put_properties, remove_properties};

/// URI assigned to the design-element extension by [MS-PPTX] 2.2.17.
pub const EXTENSION_URI: &str = "{386F3935-93C4-4BCD-93E2-E3B085C9AB24}";

/// Namespace assigned to `p15:designElem` by [MS-PPTX] 2.5.
pub const NAMESPACE: &str = "http://schemas.microsoft.com/office/powerpoint/2015/main";

/// URI assigned to shape Designer properties by [MS-PPTX] 2.2.19.
pub const PROPERTIES_EXTENSION_URI: &str = "{E7BDC344-281C-4309-B0C6-D0EE65EED2A8}";

/// URI assigned to slide Designer tags by [MS-PPTX] 2.2.20.
pub const TAGS_EXTENSION_URI: &str = "{E3EDB536-0D56-4F60-86BA-61A60CA02DAB}";

/// Namespace of the `PowerPoint` 2020 Designer schema.
pub const P202_NAMESPACE: &str = "http://schemas.microsoft.com/office/powerpoint/2020/02/main";
