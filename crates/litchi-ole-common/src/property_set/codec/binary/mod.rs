//! Binary MS-OLEPS stream and VARIANT codec facade.

mod composite;
#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_unrelated,
    reason = "tests use concise assertions while exercising fallible malformed-input paths"
)]
mod composite_tests;
mod semantic;
#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_unrelated,
    reason = "tests use concise assertions while exercising fallible malformed-input paths"
)]
mod tests;
mod validation;
mod wire;

use super::super::model::Stream;
use super::super::model::Value;
use litchi_cfb::OleError;

pub(crate) use semantic::{filetime_to_date, filetime_to_duration};

impl Stream {
    /// Parse a complete Property Set stream with section-local bounds.
    ///
    /// # Errors
    ///
    /// Returns an error if `data` is malformed, exceeds parser limits, or does
    /// not satisfy the Property Set semantic constraints.
    pub fn parse(data: &[u8]) -> Result<Self, OleError> {
        validation::parse_stream(data)
    }

    /// Serializes this stream in the canonical Property Set binary format.
    ///
    /// # Errors
    ///
    /// Returns an error if the stream violates serialization constraints, a
    /// length cannot be represented on the wire, or allocating output fails.
    pub fn to_bytes(&self) -> Result<Vec<u8>, OleError> {
        validation::serialize_stream(self)
    }
}

pub(crate) fn parse_typed_property(
    data: &[u8],
    codepage: u16,
    property_offset: usize,
) -> Result<Value, OleError> {
    semantic::parse_typed_property(data, codepage, property_offset)
}

pub(crate) fn parse_typed_property_for_property(
    data: &[u8],
    codepage: u16,
    property_offset: usize,
    property_identifier: u32,
) -> Result<Value, OleError> {
    semantic::parse_typed_property_for_property(
        data,
        codepage,
        property_offset,
        property_identifier,
    )
}
