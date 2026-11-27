//! Typed ODF `style:footnote-sep` page-layout properties.

mod codec;
mod model;
mod package;

pub(super) const STYLE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
pub(super) const MAX_XML_BYTES: usize = 64 * 1_048_576;
pub(super) const MAX_DEPTH: usize = 256;
pub(super) const MAX_SEPARATORS: usize = 65_536;
pub(super) const MAX_VALUE_BYTES: usize = 4_096;
pub(super) const MAX_AGGREGATE_BYTES: usize = 16 * 1_048_576;

pub(super) fn invalid<T>(message: impl Into<String>) -> litchi_core::Result<T> {
    Err(make_error(message))
}

pub(super) fn make_error(message: impl Into<String>) -> litchi_core::Error {
    litchi_core::Error::InvalidFormat(message.into())
}

pub use codec::parse;
pub(crate) use codec::{
    parse_page_layout_property_footnote_separators, replace_page_layout_footnote_separator,
};
pub use model::{Adjustment, Length, LineStyle, Percent, Separator};
