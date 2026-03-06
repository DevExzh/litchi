//! Sheet protection parsing for XLS BIFF8 files.
//!
//! Parses the following records that define worksheet-level protection:
//!
//! - **PROTECT** (0x0012): Sheet is protected (boolean).
//! - **OBJECTPROTECT** (0x0063): Drawing objects are protected.
//! - **SCENPROTECT** (0x00DD): Scenarios are protected.
//! - **PASSWORD** (0x0013): Password hash for the protection.
//!
//! # Record Formats
//!
//! All four records share a trivial 2-byte payload (a single `u16`):
//!
//! | Record           | Type   | Value semantics                     |
//! |------------------|--------|-------------------------------------|
//! | PROTECT          | 0x0012 | 0 = unprotected, 1 = protected      |
//! | OBJECTPROTECT    | 0x0063 | 0 = unprotected, 1 = protected      |
//! | SCENPROTECT      | 0x00DD | 0 = unprotected, 1 = protected      |
//! | PASSWORD         | 0x0013 | 16-bit hash (0 = no password)       |

use crate::common::binary;
use crate::ole::xls::error::{XlsError, XlsResult};

/// PROTECT record type.
pub const PROTECT_TYPE: u16 = 0x0012;
/// OBJECTPROTECT record type.
pub const OBJECTPROTECT_TYPE: u16 = 0x0063;
/// SCENPROTECT record type.
pub const SCENPROTECT_TYPE: u16 = 0x00DD;
/// PASSWORD record type.
pub const PASSWORD_TYPE: u16 = 0x0013;

/// Sheet protection state parsed from BIFF8 records.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct SheetProtection {
    /// Whether the sheet is protected (PROTECT record).
    pub sheet_protected: bool,
    /// Whether drawing objects are protected (OBJECTPROTECT record).
    pub objects_protected: bool,
    /// Whether scenarios are protected (SCENPROTECT record).
    pub scenarios_protected: bool,
    /// Password hash from PASSWORD record (0 = no password).
    pub password_hash: u16,
}

impl SheetProtection {
    /// Returns `true` if any protection flag is active.
    #[inline]
    pub fn is_protected(&self) -> bool {
        self.sheet_protected
    }

    /// Returns `true` if a password hash is set.
    #[inline]
    pub fn has_password(&self) -> bool {
        self.password_hash != 0
    }
}

/// Parse a boolean protection record (PROTECT, OBJECTPROTECT, SCENPROTECT).
///
/// Returns `true` if the 2-byte payload is non-zero.
pub fn parse_protect_bool(data: &[u8]) -> XlsResult<bool> {
    if data.len() < 2 {
        return Err(XlsError::InvalidLength {
            expected: 2,
            found: data.len(),
        });
    }
    Ok(binary::read_u16_le_at(data, 0)? != 0)
}

/// Parse a PASSWORD record.
///
/// Returns the raw 16-bit password hash (0 = no password).
pub fn parse_password(data: &[u8]) -> XlsResult<u16> {
    if data.len() < 2 {
        return Err(XlsError::InvalidLength {
            expected: 2,
            found: data.len(),
        });
    }
    Ok(binary::read_u16_le_at(data, 0)?)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_protect_true() {
        let data = 1u16.to_le_bytes();
        assert!(parse_protect_bool(&data).unwrap());
    }

    #[test]
    fn test_parse_protect_false() {
        let data = 0u16.to_le_bytes();
        assert!(!parse_protect_bool(&data).unwrap());
    }

    #[test]
    fn test_parse_password() {
        let data = 0xCE4Bu16.to_le_bytes();
        assert_eq!(parse_password(&data).unwrap(), 0xCE4B);
    }

    #[test]
    fn test_parse_password_zero() {
        let data = 0u16.to_le_bytes();
        assert_eq!(parse_password(&data).unwrap(), 0);
    }

    #[test]
    fn test_sheet_protection_default() {
        let prot = SheetProtection::default();
        assert!(!prot.is_protected());
        assert!(!prot.has_password());
    }
}
