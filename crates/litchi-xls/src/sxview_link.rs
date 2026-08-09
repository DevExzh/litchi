//! BIFF8 `SXViewLink` record of the chart substream (MS-XLS 2.1): the name
//! of the source PivotTable view associated with a Pivot Chart
//! (MS-XLS 2.4.316).
//!
//! Everything in this module is INERT: the fields are stored verbatim and no
//! PivotTable linkage is resolved.
//!
//! # References
//!
//! - MS-XLS 2.4.316 (SXViewLink), 2.5.296 (XLUnicodeStringNoCch)

use super::{Error, Result};

/// Record type of the `SXViewLink` record (MS-XLS 2.4.316). The `rt` field
/// of the payload MUST repeat this value.
pub(crate) const SX_VIEW_LINK_RECORD_TYPE: u16 = 0x0858;

/// Byte length of the fixed `SXViewLink` prefix: `rt` (2) + `unused` (2) +
/// `reserved` (2) + `cch` (1) (MS-XLS 2.4.316).
const HEADER_LEN: usize = 7;

/// `fHighByte` option bit of an `XLUnicodeStringNoCch` (MS-XLS 2.5.296).
const STRING_HIGH_BYTE: u8 = 0x01;

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidRecord {
        record_type: SX_VIEW_LINK_RECORD_TYPE,
        message: message.into(),
    }
}

/// Read a little-endian `u16` from a fixed offset (length checked by caller).
fn read_u16(data: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes(data[offset..offset + 2].try_into().expect("length checked"))
}

/// Typed `SXViewLink` record content (MS-XLS 2.4.316): the name of the
/// source `PivotTable` view associated with a Pivot Chart.
///
/// The `unused` and `reserved` fields MUST be ignored; they are preserved
/// verbatim so the record round-trips unchanged.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SXViewLink {
    /// Undefined two-byte field (`unused`), preserved verbatim.
    unused: u16,
    /// Reserved two-byte field (`reserved`), preserved verbatim; MUST be
    /// zero (MS-XLS 2.4.316).
    reserved: u16,
    /// Name of the `PivotTable` view (`stPivotTable`).
    pivot_table_name: String,
}

impl SXViewLink {
    /// Parse an `SXViewLink` record payload.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn parse(data: &[u8]) -> Result<Self> {
        if data.len() < HEADER_LEN {
            return Err(Error::InvalidLength {
                expected: HEADER_LEN,
                found: data.len(),
            });
        }
        let rt = read_u16(data, 0);
        if rt != SX_VIEW_LINK_RECORD_TYPE {
            return Err(invalid(format!(
                "SXViewLink rt {rt:#06X} must be {SX_VIEW_LINK_RECORD_TYPE:#06X}"
            )));
        }
        let unused = read_u16(data, 2);
        let reserved = read_u16(data, 4);
        let cch = usize::from(data[6]);

        // stPivotTable: XLUnicodeStringNoCch (MS-XLS 2.5.296).
        let flags = *data.get(HEADER_LEN).ok_or(Error::InvalidLength {
            expected: HEADER_LEN + 1,
            found: data.len(),
        })?;
        if flags & !STRING_HIGH_BYTE != 0 {
            return Err(invalid(
                "SXViewLink stPivotTable has unsupported option flags",
            ));
        }
        let wide = flags & STRING_HIGH_BYTE != 0;
        let char_bytes = cch
            .checked_mul(if wide { 2 } else { 1 })
            .ok_or_else(|| invalid("SXViewLink stPivotTable length overflow"))?;
        let expected_len = HEADER_LEN + 1 + char_bytes;
        if data.len() != expected_len {
            return Err(Error::InvalidLength {
                expected: expected_len,
                found: data.len(),
            });
        }
        let raw = &data[HEADER_LEN + 1..];
        let pivot_table_name = if wide {
            let units: Vec<u16> = raw
                .chunks_exact(2)
                .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
                .collect();
            String::from_utf16(&units)
                .map_err(|error| Error::Encoding(format!("SXViewLink stPivotTable: {error}")))?
        } else {
            // Compressed Unicode supplies an implicit zero high byte.
            raw.iter().map(|&byte| char::from(byte)).collect()
        };
        Ok(Self {
            unused,
            reserved,
            pivot_table_name,
        })
    }

    /// Serialize back to a complete `SXViewLink` record payload.
    #[must_use]
    pub fn to_payload(&self) -> Vec<u8> {
        let mut payload = Vec::with_capacity(HEADER_LEN + 1);
        payload.extend_from_slice(&SX_VIEW_LINK_RECORD_TYPE.to_le_bytes());
        payload.extend_from_slice(&self.unused.to_le_bytes());
        payload.extend_from_slice(&self.reserved.to_le_bytes());
        let units: Vec<u16> = self.pivot_table_name.encode_utf16().collect();
        payload.extend_from_slice(&crate::utils::truncate_usize_to_u8(units.len()).to_le_bytes());
        let wide = units.iter().any(|&unit| unit > 0x00FF);
        payload.push(if wide { STRING_HIGH_BYTE } else { 0 });
        for unit in units {
            if wide {
                payload.extend_from_slice(&unit.to_le_bytes());
            } else {
                payload.push(crate::utils::truncate_u16_to_u8(unit));
            }
        }
        payload
    }

    /// Undefined field, preserved verbatim.
    #[must_use]
    pub fn unused(&self) -> u16 {
        self.unused
    }

    /// Reserved field, preserved verbatim; MUST be zero (MS-XLS 2.4.316).
    #[must_use]
    pub fn reserved(&self) -> u16 {
        self.reserved
    }

    /// Name of the source `PivotTable` view associated with the Pivot Chart.
    #[must_use]
    pub fn pivot_table_name(&self) -> &str {
        &self.pivot_table_name
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Build an `SXViewLink` payload around a pre-encoded `stPivotTable`.
    fn build_payload(cch: u8, string_flags: u8, string_bytes: &[u8]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&SX_VIEW_LINK_RECORD_TYPE.to_le_bytes());
        data.extend_from_slice(&[0x00, 0x00]); // unused
        data.extend_from_slice(&[0x00, 0x00]); // reserved
        data.push(cch);
        data.push(string_flags);
        data.extend_from_slice(string_bytes);
        data
    }

    #[test]
    fn round_trip_compressed() {
        let payload = build_payload(10, 0, b"PivotTable");
        let record = SXViewLink::parse(&payload).unwrap();
        assert_eq!(record.unused(), 0);
        assert_eq!(record.reserved(), 0);
        assert_eq!(record.pivot_table_name(), "PivotTable");
        assert_eq!(record.to_payload(), payload);
    }

    #[test]
    fn round_trip_utf16() {
        // "Сводная": Cyrillic characters that do not fit in one byte.
        let name: Vec<u8> = "Сводная"
            .encode_utf16()
            .flat_map(u16::to_le_bytes)
            .collect();
        let payload = build_payload(7, STRING_HIGH_BYTE, &name);
        let record = SXViewLink::parse(&payload).unwrap();
        assert_eq!(record.pivot_table_name(), "Сводная");
        assert_eq!(record.to_payload(), payload);
    }

    #[test]
    fn round_trip_empty_name() {
        let payload = build_payload(0, 0, &[]);
        let record = SXViewLink::parse(&payload).unwrap();
        assert_eq!(record.pivot_table_name(), "");
        assert_eq!(record.to_payload(), payload);
    }

    #[test]
    fn rejects_wrong_rt() {
        let mut payload = build_payload(1, 0, b"A");
        payload[0..2].copy_from_slice(&0x0857u16.to_le_bytes());
        assert!(SXViewLink::parse(&payload).is_err());
    }

    #[test]
    fn rejects_truncation_and_trailing_garbage() {
        let payload = build_payload(10, 0, b"PivotTable");
        assert!(SXViewLink::parse(&payload[..HEADER_LEN]).is_err());
        assert!(SXViewLink::parse(&payload[..payload.len() - 1]).is_err());
        let mut longer = payload.clone();
        longer.push(0);
        assert!(SXViewLink::parse(&longer).is_err());
    }

    #[test]
    fn rejects_unsupported_string_flags() {
        // Bits other than fHighByte are reserved (MS-XLS 2.5.296).
        let payload = build_payload(10, 0x02, b"PivotTable");
        assert!(SXViewLink::parse(&payload).is_err());
    }

    #[test]
    fn rejects_count_length_mismatch() {
        // cch says 11 characters but only 10 compressed bytes follow.
        let payload = build_payload(11, 0, b"PivotTable");
        assert!(SXViewLink::parse(&payload).is_err());
        // cch says 5 characters but 10 follow.
        let payload = build_payload(5, 0, b"PivotTable");
        assert!(SXViewLink::parse(&payload).is_err());
    }

    #[test]
    fn rejects_invalid_utf16() {
        // Lone surrogate in an otherwise well-formed wide string.
        let payload = build_payload(1, STRING_HIGH_BYTE, &[0x00, 0xD8]);
        assert!(SXViewLink::parse(&payload).is_err());
    }
}
