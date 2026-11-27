//! XML-map wire codec facade.

mod encode;
mod parse;

use crate::Result;

pub use encode::write;
pub(crate) use parse::{parse_stream, validate_opaque_element};

/// # Errors
///
/// Returns an error if validation, decoding, encoding, or the requested operation fails.
pub fn parse(input: &[u8]) -> Result<super::model::MapInfo> {
    parse_stream(input)
}
