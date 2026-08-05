//! BIFF8 workbook access-provenance metadata.

use super::{Error, Result};

pub(crate) const WRITE_ACCESS_RECORD_TYPE: u16 = 0x005C;
const WRITE_ACCESS_PAYLOAD_LEN: usize = 112;
const WRITE_ACCESS_HEADER_LEN: usize = 3;
const WRITE_ACCESS_MAX_CHARACTERS: usize = 54;

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidRecord {
        record_type: WRITE_ACCESS_RECORD_TYPE,
        message: message.into(),
    }
}

/// Character storage selected by the `WriteAccess.userName` string options.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum WriteAccessEncoding {
    CompressedUnicode,
    Utf16,
}

/// The user recorded as having last created, opened, or modified a workbook.
///
/// This is inert provenance metadata. Parsing it does not authenticate or impersonate the user.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WriteAccess {
    user_name: String,
    encoding: WriteAccessEncoding,
    unused: Vec<u8>,
}

impl WriteAccess {
    /// Construct a canonical record, using compressed Unicode when possible.
    pub fn try_new(user_name: impl Into<String>) -> Result<Self> {
        let user_name = user_name.into();
        let units = user_name.encode_utf16().collect::<Vec<_>>();
        let encoding = if units.iter().all(|&unit| unit <= 0x00FF) {
            WriteAccessEncoding::CompressedUnicode
        } else {
            WriteAccessEncoding::Utf16
        };
        let byte_count = encoded_byte_count(&units, encoding)?;
        let unused = vec![b' '; WRITE_ACCESS_PAYLOAD_LEN - WRITE_ACCESS_HEADER_LEN - byte_count];
        Self::try_new_with_parts(user_name, encoding, unused)
    }

    /// Construct a record with explicit encoding and ignored bytes.
    pub fn try_new_with_parts(
        user_name: impl Into<String>,
        encoding: WriteAccessEncoding,
        unused: Vec<u8>,
    ) -> Result<Self> {
        let value = Self {
            user_name: user_name.into(),
            encoding,
            unused,
        };
        value.validate()?;
        Ok(value)
    }

    /// Parse the fixed 112-byte BIFF8 record payload.
    pub fn parse_payload(data: &[u8]) -> Result<Self> {
        if data.len() != WRITE_ACCESS_PAYLOAD_LEN {
            return Err(invalid(format!(
                "WriteAccess payload has {} bytes; expected {WRITE_ACCESS_PAYLOAD_LEN}",
                data.len()
            )));
        }
        let character_count = usize::from(u16::from_le_bytes([data[0], data[1]]));
        if character_count > WRITE_ACCESS_MAX_CHARACTERS {
            return Err(invalid(format!(
                "WriteAccess userName has {character_count} characters; maximum is 54"
            )));
        }
        if data[2] & 0xFE != 0 {
            return Err(invalid(
                "WriteAccess userName contains reserved option bits",
            ));
        }
        let encoding = if data[2] & 1 == 0 {
            WriteAccessEncoding::CompressedUnicode
        } else {
            WriteAccessEncoding::Utf16
        };
        let byte_count = character_count
            .checked_mul(match encoding {
                WriteAccessEncoding::CompressedUnicode => 1,
                WriteAccessEncoding::Utf16 => 2,
            })
            .ok_or_else(|| invalid("WriteAccess userName byte length overflows"))?;
        let end = WRITE_ACCESS_HEADER_LEN
            .checked_add(byte_count)
            .ok_or_else(|| invalid("WriteAccess userName range overflows"))?;
        let bytes = data
            .get(WRITE_ACCESS_HEADER_LEN..end)
            .ok_or_else(|| invalid("WriteAccess userName is truncated"))?;
        let units = match encoding {
            WriteAccessEncoding::CompressedUnicode => bytes
                .iter()
                .map(|&byte| u16::from(byte))
                .collect::<Vec<_>>(),
            WriteAccessEncoding::Utf16 => bytes
                .chunks_exact(2)
                .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
                .collect::<Vec<_>>(),
        };
        let user_name = String::from_utf16(&units)
            .map_err(|_| invalid("WriteAccess userName contains invalid UTF-16"))?;
        Self::try_new_with_parts(user_name, encoding, data[end..].to_vec())
    }

    pub fn user_name(&self) -> &str {
        &self.user_name
    }
    pub fn encoding(&self) -> WriteAccessEncoding {
        self.encoding
    }
    pub fn unused_bytes(&self) -> &[u8] {
        &self.unused
    }

    /// Serialize the fixed-size BIFF8 payload, preserving ignored bytes.
    pub fn to_payload(&self) -> Result<[u8; WRITE_ACCESS_PAYLOAD_LEN]> {
        let units = self.validate()?;
        let character_count = u16::try_from(units.len())
            .map_err(|_| invalid("WriteAccess character count exceeds u16"))?;
        let mut data = [0u8; WRITE_ACCESS_PAYLOAD_LEN];
        data[..2].copy_from_slice(&character_count.to_le_bytes());
        data[2] = match self.encoding {
            WriteAccessEncoding::CompressedUnicode => 0,
            WriteAccessEncoding::Utf16 => 1,
        };
        let mut offset = WRITE_ACCESS_HEADER_LEN;
        match self.encoding {
            WriteAccessEncoding::CompressedUnicode => {
                for unit in units {
                    data[offset] = unit as u8;
                    offset += 1;
                }
            },
            WriteAccessEncoding::Utf16 => {
                for unit in units {
                    let bytes = unit.to_le_bytes();
                    data[offset..offset + 2].copy_from_slice(&bytes);
                    offset += 2;
                }
            },
        }
        data[offset..].copy_from_slice(&self.unused);
        Ok(data)
    }

    /// Serialize the complete BIFF record including its four-byte record header.
    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        let mut data = Vec::with_capacity(4 + WRITE_ACCESS_PAYLOAD_LEN);
        data.extend_from_slice(&WRITE_ACCESS_RECORD_TYPE.to_le_bytes());
        data.extend_from_slice(&(WRITE_ACCESS_PAYLOAD_LEN as u16).to_le_bytes());
        data.extend_from_slice(&self.to_payload()?);
        Ok(data)
    }

    fn validate(&self) -> Result<Vec<u16>> {
        let units = self.user_name.encode_utf16().collect::<Vec<_>>();
        if units.len() > WRITE_ACCESS_MAX_CHARACTERS {
            return Err(invalid(format!(
                "WriteAccess userName has {} characters; maximum is 54",
                units.len()
            )));
        }
        let byte_count = encoded_byte_count(&units, self.encoding)?;
        let expected_unused = WRITE_ACCESS_PAYLOAD_LEN
            .checked_sub(WRITE_ACCESS_HEADER_LEN + byte_count)
            .ok_or_else(|| invalid("WriteAccess userName exceeds its fixed envelope"))?;
        if self.unused.len() != expected_unused {
            return Err(invalid(format!(
                "WriteAccess unused field has {} bytes; expected {expected_unused}",
                self.unused.len()
            )));
        }
        Ok(units)
    }
}

fn encoded_byte_count(units: &[u16], encoding: WriteAccessEncoding) -> Result<usize> {
    if units.len() > WRITE_ACCESS_MAX_CHARACTERS {
        return Err(invalid("WriteAccess userName exceeds 54 characters"));
    }
    match encoding {
        WriteAccessEncoding::CompressedUnicode => {
            if units.iter().any(|&unit| unit > 0x00FF) {
                return Err(invalid(
                    "WriteAccess compressed Unicode contains a nonzero high byte",
                ));
            }
            Ok(units.len())
        },
        WriteAccessEncoding::Utf16 => units
            .len()
            .checked_mul(2)
            .ok_or_else(|| invalid("WriteAccess UTF-16 byte length overflows")),
    }
}

pub(crate) struct WriteAccessCollector {
    seen: bool,
    value: Result<Option<WriteAccess>>,
}

impl WriteAccessCollector {
    pub(crate) fn new() -> Self {
        Self {
            seen: false,
            value: Ok(None),
        }
    }

    pub(crate) fn feed_record(&mut self, record_type: u16, data: &[u8]) {
        if record_type != WRITE_ACCESS_RECORD_TYPE {
            return;
        }
        if self.seen {
            self.value = Err(invalid("duplicate WriteAccess record"));
            return;
        }
        self.seen = true;
        self.value = WriteAccess::parse_payload(data).map(Some);
    }

    pub(crate) fn finish(self) -> Result<Option<WriteAccess>> {
        self.value
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn poi_compressed_reference() -> Vec<u8> {
        let mut data = vec![b' '; WRITE_ACCESS_PAYLOAD_LEN];
        data[..2].copy_from_slice(&12u16.to_le_bytes());
        data[2] = 0;
        data[3..15].copy_from_slice(b"Yegor Kozlov");
        data
    }

    fn libreoffice_utf16_reference() -> Vec<u8> {
        let mut data = vec![b' '; WRITE_ACCESS_PAYLOAD_LEN];
        data[..2].copy_from_slice(&6u16.to_le_bytes());
        data[2] = 1;
        for (index, unit) in "TOBIAS".encode_utf16().enumerate() {
            data[3 + index * 2..5 + index * 2].copy_from_slice(&unit.to_le_bytes());
        }
        data[15] = 0;
        data[16] = 0;
        data
    }

    #[test]
    fn parses_and_round_trips_poi_and_libreoffice_references() {
        let poi = poi_compressed_reference();
        let parsed = WriteAccess::parse_payload(&poi).unwrap();
        assert_eq!(parsed.user_name(), "Yegor Kozlov");
        assert_eq!(parsed.encoding(), WriteAccessEncoding::CompressedUnicode);
        assert_eq!(parsed.to_payload().unwrap().as_slice(), poi);

        let libreoffice = libreoffice_utf16_reference();
        let parsed = WriteAccess::parse_payload(&libreoffice).unwrap();
        assert_eq!(parsed.user_name(), "TOBIAS");
        assert_eq!(parsed.encoding(), WriteAccessEncoding::Utf16);
        assert_eq!(&parsed.unused_bytes()[..2], &[0, 0]);
        assert_eq!(parsed.to_payload().unwrap().as_slice(), libreoffice);
    }

    #[test]
    fn constructs_canonical_compressed_and_utf16_records() {
        let latin = WriteAccess::try_new("Andr\u{00e9}").unwrap();
        assert_eq!(latin.encoding(), WriteAccessEncoding::CompressedUnicode);
        assert!(latin.unused_bytes().iter().all(|&byte| byte == b' '));
        let unicode = WriteAccess::try_new("\u{6587}\u{6863}").unwrap();
        assert_eq!(unicode.encoding(), WriteAccessEncoding::Utf16);
        assert_eq!(
            WriteAccess::parse_payload(&unicode.to_payload().unwrap()).unwrap(),
            unicode
        );
        let record = unicode.to_record_bytes().unwrap();
        assert_eq!(&record[..4], &[0x5c, 0, 112, 0]);
    }

    #[test]
    fn rejects_malformed_envelopes_strings_and_duplicates() {
        let reference = poi_compressed_reference();
        assert!(WriteAccess::parse_payload(&reference[..111]).is_err());
        let mut data = reference.clone();
        data[..2].copy_from_slice(&55u16.to_le_bytes());
        assert!(WriteAccess::parse_payload(&data).is_err());
        let mut data = reference.clone();
        data[2] = 2;
        assert!(WriteAccess::parse_payload(&data).is_err());
        let mut data = reference;
        data[..2].copy_from_slice(&1u16.to_le_bytes());
        data[2] = 1;
        data[3..5].copy_from_slice(&0xD800u16.to_le_bytes());
        assert!(WriteAccess::parse_payload(&data).is_err());
        assert!(WriteAccess::try_new("x".repeat(55)).is_err());
        assert!(
            WriteAccess::try_new_with_parts(
                "\u{6587}",
                WriteAccessEncoding::CompressedUnicode,
                vec![b' '; 108],
            )
            .is_err()
        );

        let mut collector = WriteAccessCollector::new();
        let reference = poi_compressed_reference();
        collector.feed_record(WRITE_ACCESS_RECORD_TYPE, &reference);
        collector.feed_record(WRITE_ACCESS_RECORD_TYPE, &reference);
        assert!(collector.finish().is_err());
    }
}
