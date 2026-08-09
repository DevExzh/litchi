//! Shared ODF data-style grammar.
//!
//! The public facade is intentionally small: semantic models and bounded XML
//! operations are exposed here, while token parsing and XML traversal remain
//! implementation layers below it.

mod codec;
mod model;
mod package;
mod tokens;

#[cfg(test)]
mod tests;

use litchi_core::{Error, Result};

pub use codec::parse_data_styles_xml;
pub use model::{
    Calendar, Clock, Currency, EmbeddedText, FormatSource, Fraction, Kind, Locale, Map, Month,
    NumberToken, Part, Scientific, Seconds, Section, ShortLong, Style, Styles, TextProperties,
    Token, TransliterationStyle, Version, WeekOfYear,
};
pub use package::{parse_flat, parse_package, remove_data_style_xml, set_data_style_xml};

pub(crate) use codec::{
    Attribute, Frame, Node, byte_offset, collect_attributes, decode, direct_style_section,
    ensure_empty_node, ensure_no_children, ensure_whitespace, event_start, frame_section,
    namespace_uri, parse_locale, read_document_version, reject_remaining, required, required_i64,
    take, take_bool, take_f64, take_i64, take_versioned_bool, take_versioned_i64,
    take_versioned_u64, validate_cell_address, validate_locale, validate_name,
    validate_optional_string, validate_text,
};

pub(crate) const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
pub(crate) const STYLE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
pub(crate) const NUMBER: &str = "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0";
pub(crate) const LOEXT: &str =
    "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0";
pub(crate) const MAX_XML_BYTES: usize = 64 * 1_048_576;
pub(crate) const MAX_DEPTH: usize = 256;
pub(crate) const MAX_EVENTS: usize = 2_000_000;
pub(crate) const MAX_STYLES: usize = 65_536;
pub(crate) const MAX_PARTS: usize = 4_096;
pub(crate) const MAX_MAPS: usize = 1_024;
pub(crate) const MAX_ATTRIBUTES: usize = 128;
pub(crate) const MAX_VALUE_BYTES: usize = 65_536;
pub(crate) const MAX_AGGREGATE_BYTES: usize = 32 * 1_048_576;

pub(crate) fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(bad(message))
}

pub(crate) fn bad(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
