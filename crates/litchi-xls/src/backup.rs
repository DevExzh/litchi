//! BIFF8 `Backup` record (0x0040, MS-XLS 2.4.14) of the Globals substream
//! (MS-XLS 2.1): whether a backup copy of the workbook is saved.
//!
//! Everything in this module is INERT: the flag is stored verbatim and no
//! file operation is performed.
//!
//! # References
//!
//! - MS-XLS 2.4.14 (Backup), 2.5.14 (Boolean)

use super::{Error, Result};

/// Record type of the `Backup` record (MS-XLS 2.4.14).
pub(crate) const BACKUP_RECORD_TYPE: u16 = 0x0040;

/// Byte length of a `Backup` record payload: a single two-byte field.
const PAYLOAD_LEN: usize = 2;

/// Typed `Backup` record content (MS-XLS 2.4.14): whether to save a backup
/// copy of the workbook.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Backup {
    /// Whether a backup copy is saved when the workbook is saved (`fBackup`).
    save_backup: bool,
}

impl Backup {
    /// Parse a `Backup` record payload.
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
            0x0000 => Ok(Self { save_backup: false }),
            0x0001 => Ok(Self { save_backup: true }),
            other => Err(Error::InvalidRecord {
                record_type: BACKUP_RECORD_TYPE,
                message: format!("Backup fBackup {other:#06X} is not a Boolean"),
            }),
        }
    }

    /// Serialize back to a complete `Backup` record payload.
    #[must_use]
    pub fn to_payload(&self) -> Vec<u8> {
        u16::from(self.save_backup).to_le_bytes().to_vec()
    }

    /// Whether a backup copy of the workbook is saved (`fBackup`).
    #[must_use]
    pub fn save_backup(&self) -> bool {
        self.save_backup
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn backup_round_trip() {
        for (payload, expected) in [([0x00, 0x00], false), ([0x01, 0x00], true)] {
            let record = Backup::parse(&payload).unwrap();
            assert_eq!(record.save_backup(), expected);
            assert_eq!(record.to_payload(), payload);
        }
    }

    #[test]
    fn backup_rejects_bad_length_and_non_boolean() {
        assert!(Backup::parse(&[0x01]).is_err());
        assert!(Backup::parse(&[0x00, 0x00, 0x00]).is_err());
        // Boolean (MS-XLS 2.5.14) allows only 0x0000 and 0x0001.
        assert!(Backup::parse(&[0x02, 0x00]).is_err());
        assert!(Backup::parse(&[0x00, 0x01]).is_err());
    }
}
