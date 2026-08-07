//! Layered semantic owner for `OpenDocument` index source marks.
//!
//! The facade keeps the public mark vocabulary and mutation entry points stable;
//! parsing, semantic values, package edits, and regression tests live in their
//! dedicated layers.

mod codec;
mod model;
mod package;
#[cfg(test)]
mod tests;

pub(super) const MAX_MARK_DEPTH: usize = 4_096;
pub(super) const MAX_MARKS: usize = 1_000_000;

pub(crate) use codec::parse_text_index_marks;
pub(crate) use codec::{end_kind, is_bibliography_type, point_kind, start_kind};
pub use model::{
    TextAlphabeticalMarkMetadata, TextIndexMark, TextIndexMarkFragments, TextIndexMarkKind,
};
pub use package::{
    insert_text_index_mark_xml, remove_text_index_mark_xml, replace_text_index_mark_xml,
};
