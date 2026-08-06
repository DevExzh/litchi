//! Shape-owned PresentationML programmable-tag attachments.
//!
//! A shape tag list is anchored below the selected shape's application
//! non-visual properties. The relationship itself remains owned by the
//! containing slide, layout, master, notes, or handout part:
//!
//! ```text
//! p:{sp,pic,cxnSp,graphicFrame,grpSp}
//!   / p:nv*Pr / p:nvPr / p:custDataLst / p:tags@r:id
//! ```
//!
//! Name lookup is the ordinary entry point. A checked depth-first position is
//! retained by [`crate::shape::Key`] for duplicate-name repair and source-order
//! workflows. Relationship IDs and tag-part names never participate in the
//! safe selector.

mod anchor;
mod codec;
mod model;
mod validation;

#[cfg(test)]
mod tests;

pub use codec::{load, put, remove};

// Classification and other shape-owned semantic owners need the same
// selector-to-byte-span mapping as programmable tags. Keep the scanner
// implementation private while exposing this crate-internal seam to sibling
// shape owners.
pub(crate) use codec::selected_raw_span;

// `tag::package` uses this narrow lexical seam while discovering package-level
// tag anchors. The implementation remains private to the tag subsystem.
pub(crate) use codec::attribute_value_span;
