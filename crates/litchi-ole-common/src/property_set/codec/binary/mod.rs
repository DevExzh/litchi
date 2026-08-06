//! Binary MS-OLEPS stream and VARIANT codec facade.

mod semantic;
#[cfg(test)]
mod tests;
mod validation;
mod wire;

use super::super::model::Stream;
use super::super::model::Value;
use litchi_cfb::OleError;

pub(crate) use semantic::{filetime_to_date, filetime_to_duration};

pub(crate) fn parse_typed_property(
    data: &[u8],
    codepage: u16,
    property_offset: usize,
) -> Result<Value, OleError> {
    semantic::parse_typed_property(data, codepage, property_offset)
}

impl Stream {
    /// Parse a complete Property Set stream with section-local bounds.
    pub fn parse(data: &[u8]) -> Result<Self, OleError> {
        validation::parse_stream(data)
    }

    pub fn to_bytes(&self) -> Result<Vec<u8>, OleError> {
        validation::serialize_stream(self)
    }
}
