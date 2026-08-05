//! Semantic access to point and range cross-reference targets.

mod codec;
mod model;
mod writing;

#[cfg(test)]
mod tests;

pub use model::ReferenceMark;
pub use writing::{
    ReferenceMarkFragments, insert_reference_mark_xml, remove_reference_mark_xml,
    replace_reference_mark_xml,
};

pub(crate) use codec::parse_reference_marks;

pub(super) const MAX_DEPTH: usize = 4_096;
pub(super) const MAX_MARKS: usize = 1_000_000;
