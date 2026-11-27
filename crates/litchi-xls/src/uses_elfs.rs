//! BIFF8 `UsesELFs` record (0x0160, MS-XLS 2.4.337) of the Globals substream
//! (MS-XLS 2.1): whether the file supports natural language formulas.
//!
//! Everything in this module is INERT: the flag is stored verbatim and no
//! natural language formula is evaluated.
//!
//! # References
//!
//! - MS-XLS 2.4.337 (UsesELFs), 2.5.14 (Boolean)

use super::{Error, Result};

/// Record type of the `UsesELFs` record (MS-XLS 2.4.337).
pub(crate) const USES_ELFS_RECORD_TYPE: u16 = 0x0160;

/// Byte length of a `UsesELFs` record payload: the `useselfs` field.
const PAYLOAD_LEN: usize = 2;

/// Typed `UsesELFs` record content (MS-XLS 2.4.337): whether the file
/// supports natural language formulas.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct UsesElfs {
    /// Whether the file supports natural language formulas (`useselfs`).
    uses_elfs: bool,
}

impl UsesElfs {
    /// Parse a `UsesELFs` record payload.
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
        let value = u16::from_le_bytes(data[0..2].try_into().expect("length checked"));
        // Boolean (MS-XLS 2.5.14): only 0x0000 and 0x0001 are legal.
        match value {
            0x0000 => Ok(Self { uses_elfs: false }),
            0x0001 => Ok(Self { uses_elfs: true }),
            other => Err(Error::InvalidRecord {
                record_type: USES_ELFS_RECORD_TYPE,
                message: format!("UsesELFs useselfs {other:#06X} is not a Boolean"),
            }),
        }
    }

    /// Serialize back to a complete `UsesELFs` record payload.
    #[must_use]
    pub fn to_payload(&self) -> Vec<u8> {
        u16::from(self.uses_elfs).to_le_bytes().to_vec()
    }

    /// Whether the file supports natural language formulas (`useselfs`).
    #[must_use]
    pub fn uses_elfs(&self) -> bool {
        self.uses_elfs
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn uses_elfs_round_trip() {
        for (payload, expected) in [([0x00, 0x00], false), ([0x01, 0x00], true)] {
            let record = UsesElfs::parse(&payload).unwrap();
            assert_eq!(record.uses_elfs(), expected);
            assert_eq!(record.to_payload(), payload);
        }
    }

    #[test]
    fn uses_elfs_rejects_bad_length_and_non_boolean() {
        assert!(UsesElfs::parse(&[0x01]).is_err());
        assert!(UsesElfs::parse(&[0x00, 0x00, 0x00]).is_err());
        // Boolean (MS-XLS 2.5.14) allows only 0x0000 and 0x0001.
        assert!(UsesElfs::parse(&[0x02, 0x00]).is_err());
        assert!(UsesElfs::parse(&[0x00, 0x01]).is_err());
    }
}
