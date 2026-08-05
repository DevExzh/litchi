//! Semantic ODF outline numbering-style facade.
//!
//! The model layer owns typed outline values and their invariants. The codec
//! layer owns bounded XML conversion and byte-preserving edits. Package and
//! flat-document adapters remain at this contextual owner.

mod codec;
mod model;

#[cfg(test)]
mod tests;

use litchi_core::{Error, Result};

pub use codec::{parse_outline_styles, remove_outline_style_xml, set_outline_style_xml};
pub use model::{
    Attribute, LevelStyle, ListProperties, NumberFormat, PositionMode, PositiveInteger, Style,
    Styles, TextAlign, TextProperties,
};

use crate::{FlatDocument, Package};

pub(super) const OFFICE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
pub(super) const TEXT: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
pub(super) const STYLE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
pub(super) const FO: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
pub(super) const MAX_XML_BYTES: usize = 64 * 1024 * 1024;
pub(super) const MAX_DEPTH: usize = 256;
pub(super) const MAX_STYLES: usize = 1_024;
pub const MAX_OUTLINE_LEVELS: u16 = 1_024;
pub(super) const MAX_VALUE_BYTES: usize = 64 * 1024;
pub(super) const MAX_TOTAL_ATTRIBUTE_BYTES: usize = 16 * 1024 * 1024;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub(super) enum NamespaceKind {
    Office,
    Text,
    Style,
    Fo,
    Other,
}

pub(super) fn namespace_kind(namespace: &[u8]) -> NamespaceKind {
    match namespace {
        OFFICE => NamespaceKind::Office,
        TEXT => NamespaceKind::Text,
        STYLE => NamespaceKind::Style,
        FO => NamespaceKind::Fo,
        _ => NamespaceKind::Other,
    }
}

pub(super) fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(invalid_error(message))
}

pub(super) fn invalid_error(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

impl Package {
    /// Return outline numbering styles from styles.xml.
    pub fn outline_styles(&self) -> Result<Styles> {
        self.styles_xml()?.map_or_else(
            || Ok(Styles::default()),
            |styles| parse_outline_styles(&styles),
        )
    }
}

impl FlatDocument {
    /// Return outline numbering styles without interpreting heading content.
    pub fn outline_styles(&self) -> Result<Styles> {
        parse_outline_styles(self.xml())
    }
}
