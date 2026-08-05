//! Semantic facade for ODF text line-numbering configuration.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use codec::parse;
pub use model::{Configuration, Format, NonNegativeLength, Position, Separator};

pub(crate) use codec::{remove_xml, set_xml};

pub(super) const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
pub(super) const TEXT_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
pub(super) const STYLE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
pub(super) const MAX_DOCUMENT_XML_BYTES: usize = 32 * 1024 * 1024;
pub(super) const MAX_VALUE_BYTES: usize = 64 * 1024;
pub(super) const MAX_XML_DEPTH: usize = 128;

pub(super) fn invalid<T>(message: impl Into<String>) -> litchi_core::Result<T> {
    Err(litchi_core::Error::InvalidFormat(message.into()))
}

pub(super) fn xml_error(error: impl std::fmt::Display) -> litchi_core::Error {
    litchi_core::Error::InvalidFormat(format!("invalid ODF line-numbering XML: {error}"))
}
