//! Owned UTF-8 storage shared by decoded ODF XML parts.

use crate::core::XmlStreamLimits;
use crate::limits::{LimitKind, RawXmlLimitKind};
use litchi_core::{Error, Result};

/// A validated, immutable UTF-8 XML part.
#[allow(
    clippy::module_name_repetitions,
    reason = "`XmlPart` is the established public name for the XML storage primitive."
)]
#[derive(Debug)]
pub struct XmlPart {
    content: String,
}

impl XmlPart {
    /// Copy one XML part admitted under the raw policy's input-byte bound.
    ///
    /// # Errors
    ///
    /// Returns [`LimitKind::RawXml`] when the part exceeds the policy's
    /// input bytes, and an invalid-format error when it is not UTF-8.
    pub fn from_bytes(
        bytes: &[u8],
        limits: XmlStreamLimits,
    ) -> std::result::Result<Self, crate::Error> {
        admit_raw(bytes.len(), limits)?;
        Ok(Self::copy(bytes)?)
    }

    /// Adopt one owned XML part admitted under the raw policy's input-byte
    /// bound, without copying the buffer.
    ///
    /// # Errors
    ///
    /// As [`Self::from_bytes`].
    pub fn from_vec(
        bytes: Vec<u8>,
        limits: XmlStreamLimits,
    ) -> std::result::Result<Self, crate::Error> {
        admit_raw(bytes.len(), limits)?;
        Ok(Self::adopt(bytes)?)
    }

    /// Copy bytes already admitted by the caller's part policy.
    pub(crate) fn copy(bytes: &[u8]) -> Result<Self> {
        let value = std::str::from_utf8(bytes).map_err(|error| {
            Error::InvalidFormat(format!("Invalid UTF-8 in XML content: {error}"))
        })?;
        let mut content = String::new();
        content
            .try_reserve_exact(value.len())
            .map_err(|source| Error::Allocation {
                resource: "ODF XML part",
                source,
            })?;
        content.push_str(value);
        Ok(Self { content })
    }

    /// Adopt an owned buffer already admitted by the caller's part policy.
    pub(crate) fn adopt(bytes: Vec<u8>) -> Result<Self> {
        let content = String::from_utf8(bytes).map_err(|error| {
            Error::InvalidFormat(format!("Invalid UTF-8 in XML content: {error}"))
        })?;
        Ok(Self { content })
    }

    /// Borrow the raw XML content.
    #[must_use]
    pub fn content(&self) -> &str {
        &self.content
    }

    /// Borrow the raw XML bytes without another allocation.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        self.content.as_bytes()
    }
}

/// Admit one part under a raw XML policy's input-byte bound.
pub(crate) fn admit_raw(
    observed: usize,
    limits: XmlStreamLimits,
) -> std::result::Result<(), crate::Error> {
    let observed = u64::try_from(observed).unwrap_or(u64::MAX);
    if observed > limits.max_bytes() {
        return Err(crate::Error::Limit {
            kind: LimitKind::RawXml(RawXmlLimitKind::InputBytes),
            observed,
            maximum: limits.max_bytes(),
        });
    }
    Ok(())
}
