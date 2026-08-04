//! Semantic ODF font-face declarations.
//!
//! Linked font resources are exposed as inert metadata. This owner never loads
//! a URI, installs a font, or interprets embedded font data.

mod codec;
mod model;
mod package;
#[cfg(test)]
mod tests;

use litchi_core::{Error, Result};

pub(super) const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
pub(super) const STYLE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
pub(super) const SVG_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
pub(super) const XLINK_NAMESPACE: &[u8] = b"http://www.w3.org/1999/xlink";
pub(super) const MAX_DOCUMENT_XML_BYTES: usize = 32 * 1024 * 1024;
pub(super) const MAX_FONT_FACES: usize = 4_096;
pub(super) const MAX_SOURCES_PER_FACE: usize = 1_024;
pub(super) const MAX_FORMATS_PER_SOURCE: usize = 64;
pub(super) const MAX_VALUE_BYTES: usize = 64 * 1024;
pub(super) const MAX_AGGREGATE_TEXT_BYTES: usize = 8 * 1024 * 1024;
pub(super) const MAX_XML_DEPTH: usize = 128;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub(super) enum NamespaceKind {
    Office,
    Style,
    Svg,
    Xlink,
    Other,
}

pub(super) fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}

pub(super) fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::InvalidFormat(format!("invalid ODF font-face XML: {error}"))
}

pub use codec::parse_font_face_declarations;
pub use model::{
    Face, Faces, GenericFamily, Link, Metric, MetricKind, Pitch, PositiveLength, Source, Stretch,
    Style, Variant, Weight,
};
pub(crate) use package::{
    parse_content_font_face_declarations, parse_styles_font_face_declarations,
    remove_content_font_face_declarations_xml, remove_styles_font_face_declarations_xml,
    set_content_font_face_declarations_xml, set_styles_font_face_declarations_xml,
};
