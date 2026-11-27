//! BIFF8 user-interface collection markers of the Globals substream
//! (MS-XLS 2.1):
//!
//! - **InterfaceHdr** (0x00E1): the code page of the user interface and the
//!   beginning of the user-interface record collection (MS-XLS 2.4.146).
//! - **InterfaceEnd** (0x00E2): the end of that collection (MS-XLS 2.4.145).
//!
//! Everything in this module is INERT: the markers are stored verbatim and no
//! user-interface state is reconstructed.
//!
//! # References
//!
//! - MS-XLS 2.4.145 (InterfaceEnd), 2.4.146 (InterfaceHdr)

use super::{Error, Result};

/// Record type of the `InterfaceHdr` record (MS-XLS 2.4.146).
pub(crate) const INTERFACE_HDR_RECORD_TYPE: u16 = 0x00E1;

/// Byte length of an `InterfaceHdr` record payload.
const PAYLOAD_LEN: usize = 2;

/// The only legal `codePage` value: 0x04B0 (Unicode) (MS-XLS 2.4.146).
const CODE_PAGE_UNICODE: u16 = 0x04B0;

/// Typed `InterfaceHdr` record content (MS-XLS 2.4.146): the code page of
/// the user interface.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct InterfaceHdr {
    /// The user-interface code page. Always 0x04B0 (Unicode).
    code_page: u16,
}

impl InterfaceHdr {
    /// Parse an `InterfaceHdr` record payload.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    /// # Panics
    ///
    /// Panics only if an internal BIFF invariant has been violated.
    pub fn parse(data: &[u8]) -> Result<Self> {
        if data.len() != PAYLOAD_LEN {
            return Err(Error::InvalidLength {
                expected: PAYLOAD_LEN,
                found: data.len(),
            });
        }
        let code_page = u16::from_le_bytes(data[0..2].try_into().expect("length checked"));
        if code_page != CODE_PAGE_UNICODE {
            return Err(Error::InvalidRecord {
                record_type: INTERFACE_HDR_RECORD_TYPE,
                message: format!("InterfaceHdr codePage {code_page:#06X} is not 0x04B0"),
            });
        }
        Ok(Self { code_page })
    }

    /// Serialize back to a complete `InterfaceHdr` record payload.
    #[must_use]
    pub fn to_payload(&self) -> Vec<u8> {
        self.code_page.to_le_bytes().to_vec()
    }

    /// The user-interface code page (`codePage`). Always 0x04B0 (Unicode).
    #[must_use]
    pub fn code_page(&self) -> u16 {
        self.code_page
    }
}

/// Typed `InterfaceEnd` record content (MS-XLS 2.4.145): the end of the
/// user-interface record collection. The record has no fields.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct InterfaceEnd;

impl InterfaceEnd {
    /// Parse an `InterfaceEnd` record payload.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn parse(data: &[u8]) -> Result<Self> {
        if !data.is_empty() {
            return Err(Error::InvalidLength {
                expected: 0,
                found: data.len(),
            });
        }
        Ok(Self)
    }

    /// Serialize back to a complete `InterfaceEnd` record payload.
    #[must_use]
    pub fn to_payload(&self) -> Vec<u8> {
        Vec::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn interface_hdr_round_trip() {
        let payload = [0xB0, 0x04];
        let parsed = InterfaceHdr::parse(&payload).unwrap();
        assert_eq!(parsed.code_page(), 0x04B0);
        assert_eq!(parsed.to_payload(), payload);
    }

    #[test]
    fn interface_hdr_rejects_bad_length_and_code_page() {
        assert!(InterfaceHdr::parse(&[0xB0]).is_err());
        assert!(InterfaceHdr::parse(&[0xB0, 0x04, 0x00]).is_err());
        assert!(InterfaceHdr::parse(&[0xE4, 0x03]).is_err());
    }

    #[test]
    fn interface_end_round_trip() {
        let parsed = InterfaceEnd::parse(&[]).unwrap();
        assert_eq!(parsed.to_payload(), Vec::<u8>::new());
        assert!(InterfaceEnd::parse(&[0x00]).is_err());
    }
}
