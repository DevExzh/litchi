//! Inert wildcard XML payloads.

use crate::{Error, Result};

const MAX_OPAQUE_BYTES: usize = 16 * 1024 * 1024;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OpaqueXml(Vec<u8>);

impl OpaqueXml {
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn try_new(bytes: impl Into<Vec<u8>>) -> Result<Self> {
        let bytes = bytes.into();
        if bytes.is_empty() || bytes.len() > MAX_OPAQUE_BYTES {
            return Err(invalid(
                "opaque XML element is empty or exceeds its 16 MiB limit",
            ));
        }
        crate::xml_map::codec::validate_opaque_element(&bytes)?;
        Ok(Self(bytes))
    }

    pub(crate) fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        if bytes.is_empty() || bytes.len() > MAX_OPAQUE_BYTES {
            return Err(invalid(
                "opaque XML element is empty or exceeds its 16 MiB limit",
            ));
        }
        std::str::from_utf8(&bytes)
            .map_err(|error| invalid(format!("opaque XML element is not UTF-8: {error}")))?;
        Ok(Self(bytes))
    }

    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        &self.0
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidData(format!("XML map: {}", message.into()))
}
