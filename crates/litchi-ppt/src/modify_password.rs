//! Inert modify-password metadata from MS-PPT 2.4.7.

use std::fmt;

use crate::consts::RecordType;

use super::package::{Error, Result};
use super::records::Record;

const MAX_PASSWORD_BYTES: usize = 510;

/// A bounded `PowerPoint` modify-password value.
///
/// The secret is intentionally redacted from `Debug` and is only returned by
/// the explicitly named [`Self::expose_secret`] method. This type does not
/// verify passwords, decrypt files, or grant modification access.
#[derive(Clone, PartialEq, Eq)]
pub struct ModifyPassword {
    value: String,
}

impl ModifyPassword {
    /// Construct a canonical printable Unicode modify password.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let secret = value.into();
        validate_value(&secret)?;
        Ok(Self { value: secret })
    }

    /// Parse a `ModifyPasswordAtom` represented by a `CString` record.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse(record: &Record) -> Result<Self> {
        if record.record_type != RecordType::CString
            || record.version != 0
            || record.instance != 3
            || record.data.len() > MAX_PASSWORD_BYTES
            || !record.data.len().is_multiple_of(2)
        {
            return Err(Error::Corrupted(
                "ModifyPasswordAtom has an invalid record header or size".to_string(),
            ));
        }
        let mut units = Vec::with_capacity(record.data.len() / 2);
        for bytes in record.data.chunks_exact(2) {
            let unit = u16::from_le_bytes([bytes[0], bytes[1]]);
            if unit == 0 {
                break;
            }
            if is_forbidden_printable_unit(unit) {
                return Err(Error::Corrupted(
                    "ModifyPasswordAtom contains a forbidden control character".to_string(),
                ));
            }
            units.push(unit);
        }
        let value = String::from_utf16(&units).map_err(|_err| {
            Error::Corrupted("ModifyPasswordAtom contains invalid UTF-16".to_string())
        })?;
        Ok(Self { value })
    }

    /// Discover the single modify-password atom in the PPT10 document tag.
    pub(crate) fn parse_document(document: &Record) -> Result<Option<Self>> {
        let records = document.versioned_binary_tag_records(10)?;
        let mut matches = records
            .iter()
            .filter(|record| record.record_type == RecordType::CString && record.instance == 3);
        let Some(record) = matches.next() else {
            return Ok(None);
        };
        if matches.next().is_some() {
            return Err(Error::Corrupted(
                "PPT10 document tag contains multiple ModifyPasswordAtom records".to_string(),
            ));
        }
        Self::parse(record).map(Some)
    }

    /// Explicitly expose the secret string.
    #[must_use]
    pub fn expose_secret(&self) -> &str {
        &self.value
    }

    /// Return the number of UTF-16 code units in the secret.
    #[must_use]
    pub fn len_utf16(&self) -> usize {
        self.value.encode_utf16().count()
    }

    /// Return whether the secret is empty.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.value.is_empty()
    }

    /// Encode a canonical `ModifyPasswordAtom` record without a terminator.
    #[allow(
        clippy::cast_possible_truncation,
        reason = "the password is validated to at most 510 bytes, always representable as u32"
    )]
    pub fn to_record(&self) -> Record {
        let data: Vec<u8> = self
            .value
            .encode_utf16()
            .flat_map(u16::to_le_bytes)
            .collect();
        Record {
            record_type: RecordType::CString,
            record_type_raw: 4026,
            version: 0,
            instance: 3,
            data_length: data.len() as u32,
            data,
            children: Vec::new(),
        }
    }
}

impl fmt::Debug for ModifyPassword {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ModifyPassword")
            .field("utf16_units", &self.len_utf16())
            .field("value", &"[REDACTED]")
            .finish()
    }
}

pub(crate) fn validate_value(value: &str) -> Result<()> {
    validate_printable_unicode(value)
}

fn validate_printable_unicode(value: &str) -> Result<()> {
    let mut bytes = 0usize;
    for unit in value.encode_utf16() {
        if is_forbidden_printable_unit(unit) {
            return Err(Error::Corrupted(
                "Modify password contains a forbidden control character".to_string(),
            ));
        }
        bytes = bytes
            .checked_add(2)
            .ok_or_else(|| Error::Corrupted("Modify password length overflow".to_string()))?;
        if bytes > MAX_PASSWORD_BYTES {
            return Err(Error::Corrupted(
                "Modify password exceeds the MS-PPT 510-byte limit".to_string(),
            ));
        }
    }
    Ok(())
}

const fn is_forbidden_printable_unit(unit: u16) -> bool {
    matches!(unit, 0x0000..=0x001f | 0x007f..=0x009f)
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;

    #[test]
    fn round_trips_and_redacts_modify_password() {
        let password = ModifyPassword::new("s3cret\u{1f34b}").unwrap();
        let parsed = ModifyPassword::parse(&password.to_record()).unwrap();
        assert_eq!(parsed.expose_secret(), "s3cret\u{1f34b}");
        let debug = format!("{parsed:?}");
        assert!(debug.contains("[REDACTED]"));
        assert!(!debug.contains("s3cret"));
    }

    #[test]
    fn null_terminates_input_but_writer_is_canonical() {
        let mut record = ModifyPassword::new("visible").unwrap().to_record();
        record.data.extend_from_slice(&0u16.to_le_bytes());
        record.data.extend_from_slice(&('x' as u16).to_le_bytes());
        record.data_length = u32::try_from(record.data.len()).unwrap();
        let parsed = ModifyPassword::parse(&record).unwrap();
        assert_eq!(parsed.expose_secret(), "visible");
        assert_eq!(parsed.to_record().data.len(), 14);
    }

    #[test]
    fn rejects_controls_invalid_utf16_and_oversize_values() {
        assert!(ModifyPassword::new("bad\nvalue").is_err());
        assert!(ModifyPassword::new("x".repeat(256)).is_err());

        let record = Record {
            record_type: RecordType::CString,
            record_type_raw: 4026,
            version: 0,
            instance: 3,
            data_length: 2,
            data: 0xd800u16.to_le_bytes().to_vec(),
            children: Vec::new(),
        };
        assert!(ModifyPassword::parse(&record).is_err());
    }
}
