//! Inert, bounded discovery of `text:alphabetical-index-auto-mark-file` links.
//!
//! The owner keeps the public reference model, XML codec, package-part
//! aggregation, and regression coverage in separate layers. An alphabetical-
//! index auto-mark file names an external concordance file whose words
//! generate alphabetical index entries automatically; the reference is
//! retained for inspection only and the file is never fetched or loaded.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

use litchi_core::{Error, Result};

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const XLINK: &str = "http://www.w3.org/1999/xlink";
const MAX_XML_BYTES: usize = 64 * 1_048_576;
const MAX_DEPTH: usize = 256;
const MAX_OCCURRENCES: usize = 1_024;
const MAX_VALUE_BYTES: usize = 65_536;
const MAX_AGGREGATE_BYTES: usize = 16 * 1_048_576;

pub use model::AlphabeticalIndexAutoMarkFile;
pub(crate) use package::parse_auto_mark_file_parts;

pub(super) fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(make_error(message))
}

pub(super) fn make_error(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
