//! Typed, bounded ODF variable declarations.
//!
//! The public surface is kept at the owner boundary while the implementation
//! is split into semantic models, XML codecs, package aggregation, and
//! regression tests. The codec deliberately edits only the selected XML span,
//! so unrelated and unknown document content remains byte-for-byte intact.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub(super) const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
pub(super) const STYLE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
pub(super) const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
pub(super) const MAX_XML_BYTES: usize = 64 * 1_048_576;
pub(super) const MAX_DEPTH: usize = 256;
pub(super) const MAX_GROUPS: usize = 4_096;
pub(super) const MAX_DECLARATIONS: usize = 65_536;
pub(super) const MAX_NAME_BYTES: usize = 65_536;
pub(super) const MAX_VALUE_BYTES: usize = 1_048_576;
pub(super) const MAX_AGGREGATE_BYTES: usize = 16 * 1_048_576;

pub(crate) use codec::splice_publication;
pub use codec::{remove_xml, set_xml};
pub use model::{
    Body, DateValue, Declaration, Declarations, Group, HeaderFooter, Kind, Part, Scope, Value,
    ValueType,
};
pub(crate) use package::parse_parts;
