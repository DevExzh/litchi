//! MS-OLEPS parsing, serialization, and transactional CFB integration.

mod binary;
mod editor;
mod package;
mod semantic;
mod support;
#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_unrelated,
    clippy::cast_possible_truncation,
    reason = "tests use concise assertions and checked fixture-sized literals"
)]
mod tests;

pub use editor::Editor;
pub use package::{PropertySetReader, SharedPropertySetReader};

use super::model::{Section, Value};
use litchi_cfb::OleError;

pub(crate) fn validate_section(section: &Section, version: u16) -> Result<(), OleError> {
    semantic::validate_section(section, version)
}

pub(crate) fn parse_non_simple_stream(data: &[u8]) -> Result<super::model::Stream, OleError> {
    binary::parse_non_simple_stream(data)
}

pub(crate) fn decode_non_simple_unknown_property(
    variant_type: u16,
    data: &[u8],
    codepage: u16,
    property_identifier: u32,
) -> Result<Value, OleError> {
    let length = data
        .len()
        .checked_add(4)
        .ok_or_else(|| super::model::invalid("non-simple indirect value length overflows"))?;
    let mut encoded = Vec::new();
    encoded
        .try_reserve_exact(length)
        .map_err(|source| OleError::Allocation {
            resource: "non-simple indirect value validation",
            source,
        })?;
    encoded.extend_from_slice(&variant_type.to_le_bytes());
    encoded.extend_from_slice(&0u16.to_le_bytes());
    encoded.extend_from_slice(data);
    binary::parse_typed_property_for_non_simple_property(&encoded, codepage, 0, property_identifier)
}

// Keep the existing private property_set::codec test seam available to the
// parent module without exposing binary implementation details to consumers.
#[cfg(test)]
pub(super) use binary::{filetime_to_date, filetime_to_duration, parse_typed_property};
