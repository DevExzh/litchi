//! Layered binary codecs for Word stylesheets.

mod parser;

use crate::package::Error as PackageError;

#[cfg(test)]
pub(super) use parser::{NIL_STYLE, parse_style_revision};

pub(super) fn corrupted(message: &str) -> PackageError {
    PackageError::Corrupted(format!("invalid stylesheet: {message}"))
}

pub(super) fn read_u16(data: &[u8], offset: usize, field: &str) -> crate::package::Result<u16> {
    litchi_core::binary::read_u16_le(data, offset)
        .map_err(|error| corrupted(&format!("invalid {field}: {error}")))
}

pub(super) fn read_i16(data: &[u8], offset: usize, field: &str) -> crate::package::Result<i16> {
    Ok(read_u16(data, offset, field)? as i16)
}

pub(super) fn read_u32(data: &[u8], offset: usize, field: &str) -> crate::package::Result<u32> {
    litchi_core::binary::read_u32_le(data, offset)
        .map_err(|error| corrupted(&format!("invalid {field}: {error}")))
}
