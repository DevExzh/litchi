//! Package-facing resource ceilings for one ODF annotation fragment.
//!
//! The package readers retain the annotation model in memory while walking
//! content.xml. These limits bound the resulting tree even when the caller
//! supplies events from a larger document.

pub(crate) const MAX_ANNOTATION_BODY_ELEMENTS: usize = 65_536;
pub(crate) const MAX_ANNOTATION_ELEMENTS: usize = 65_536;
pub(crate) const MAX_ANNOTATION_NESTING: usize = 128;
pub(crate) const MAX_ANNOTATION_ATTRIBUTES: usize = 4_096;
pub(crate) const MAX_ANNOTATION_ATTRIBUTE_BYTES: usize = 16 * 1024 * 1024;
pub(crate) const MAX_ANNOTATION_TEXT_BYTES: usize = 16 * 1024 * 1024;
