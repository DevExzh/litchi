//! BIFF8 `EntExU2` record (0x01C2, MS-XLS 2.4.102): an application-specific
//! cache of information.
//!
//! The record SHOULD NOT be written and SHOULD be ignored (MS-XLS 2.4.102):
//! the cache exists for performance reasons only and can be rebuilt from
//! information stored elsewhere in the file. Everything in this module is
//! therefore INERT: the bytes are stored verbatim so the record round-trips
//! unchanged, and they are never interpreted. The `rgb` field has no
//! spec-defined structure or size constraints, so any payload is accepted.
//!
//! # References
//!
//! - MS-XLS 2.4.102 (EntExU2)

use super::Result;

/// Typed `EntExU2` record content (MS-XLS 2.4.102): an application-specific
/// cache of information, preserved as opaque bytes.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EntExU2 {
    /// The opaque `rgb` cache bytes.
    cache: Vec<u8>,
}

impl EntExU2 {
    /// Parse an `EntExU2` record payload.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn parse(data: &[u8]) -> Result<Self> {
        Ok(Self {
            cache: data.to_vec(),
        })
    }

    /// Serialize back to a complete `EntExU2` record payload.
    #[must_use]
    pub fn to_payload(&self) -> Vec<u8> {
        self.cache.clone()
    }

    /// The opaque `rgb` cache bytes.
    #[must_use]
    pub fn cache(&self) -> &[u8] {
        &self.cache
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn round_trip() {
        // The rgb field has no spec-defined structure or size constraints.
        for payload in [&b""[..], &b"x"[..], &[0xDE, 0xAD, 0xBE, 0xEF], &[0; 64]] {
            let parsed = EntExU2::parse(payload).unwrap();
            assert_eq!(parsed.cache(), payload);
            assert_eq!(parsed.to_payload(), payload);
        }
    }
}
