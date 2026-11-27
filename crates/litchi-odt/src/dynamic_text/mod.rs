//! Namespace-aware mutation of dynamic text fields in ODT `content.xml`.
//!
//! The owner keeps the public XML mutation facade at the package boundary while
//! separating bounded scan state, XML mechanics, and regression coverage.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

#[cfg(test)]
pub(super) use codec::scan;
pub(super) use model::{ParagraphSite, Scan, Span};

pub use package::{
    insert_database_field_xml, insert_dynamic_text_field_xml, remove_database_field_xml,
    remove_dynamic_text_field_xml, replace_database_field_xml, replace_dynamic_text_field_xml,
};

pub(super) const TEXT_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
pub(super) const MAX_CONTENT_XML_BYTES: usize = 64 * 1024 * 1024;
pub(super) const MAX_DYNAMIC_FIELDS: usize = 1_000_000;
pub(super) const MAX_PARAGRAPHS: usize = 1_000_000;
pub(super) const MAX_XML_DEPTH: usize = 4_096;
