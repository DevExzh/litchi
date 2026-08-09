//! BIFF8 `CellWatch` record (0x086C, MS-XLS 2.4.41) of the worksheet substream
//! (MS-XLS 2.1): a reference to a watched cell.
//!
//! Everything in this module is INERT: the cell reference is stored verbatim
//! and no watch window is updated.
//!
//! # References
//!
//! - MS-XLS 2.4.41 (CellWatch), 2.5.134 (FrtFlags), 2.5.139 (FrtRefHeaderU),
//!   2.5.209 (Ref8U)

use super::{Error, Result};

/// Record type of the `CellWatch` record (MS-XLS 2.4.41); also the required
/// `frtRefHeaderU.rt` value.
pub(crate) const CELL_WATCH_RECORD_TYPE: u16 = 0x086C;

/// Byte length of a `CellWatch` record payload: `FrtRefHeaderU` (12) +
/// `reserved` (4).
const PAYLOAD_LEN: usize = 16;

/// `FrtFlags.fFrtRef` bit: MUST be 1 in a `CellWatch` record (MS-XLS 2.4.41).
const FRT_FLAG_REF: u16 = 0x0001;
/// `FrtFlags.fFrtAlert` bit: MUST be zero in an `FrtRefHeaderU` (MS-XLS 2.5.139).
const FRT_FLAG_ALERT: u16 = 0x0002;
/// Maximum column index of a `Ref8U` (MS-XLS 2.5.209).
const MAX_COLUMN_INDEX: u16 = 0x00FF;

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidRecord {
        record_type: CELL_WATCH_RECORD_TYPE,
        message: message.into(),
    }
}

/// Typed `CellWatch` record content (MS-XLS 2.4.41): a reference to a
/// watched cell.
///
/// The `grbitFrt` reserved bits and the trailing `reserved` field (MUST be
/// zero, and MUST be ignored) are preserved verbatim so the record
/// round-trips unchanged.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CellWatch {
    /// Raw `frtRefHeaderU.grbitFrt` bitfield. `fFrtRef` is guaranteed set and
    /// `fFrtAlert` guaranteed clear; the undefined reserved bits are preserved.
    flags: u16,
    /// `ref8.rwFirst`: zero-based index of the first row of the watched range.
    row_first: u16,
    /// `ref8.rwLast`: zero-based index of the last row of the watched range.
    row_last: u16,
    /// `ref8.colFirst`: zero-based index of the first column of the watched range.
    column_first: u16,
    /// `ref8.colLast`: zero-based index of the last column of the watched range.
    column_last: u16,
    /// Trailing `reserved` field, preserved verbatim.
    reserved: [u8; 4],
}

impl CellWatch {
    /// Parse a `CellWatch` record payload.
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
        if u16::from_le_bytes([data[0], data[1]]) != CELL_WATCH_RECORD_TYPE {
            return Err(invalid("CellWatch FrtRefHeaderU.rt mismatch"));
        }
        let flags = u16::from_le_bytes([data[2], data[3]]);
        if flags & FRT_FLAG_REF == 0 {
            return Err(invalid("CellWatch FrtRefHeaderU.grbitFrt.fFrtRef is not 1"));
        }
        if flags & FRT_FLAG_ALERT != 0 {
            return Err(invalid(
                "CellWatch FrtRefHeaderU.grbitFrt.fFrtAlert is not 0",
            ));
        }
        let row_first = u16::from_le_bytes([data[4], data[5]]);
        let row_last = u16::from_le_bytes([data[6], data[7]]);
        let column_first = u16::from_le_bytes([data[8], data[9]]);
        let column_last = u16::from_le_bytes([data[10], data[11]]);
        // Ref8U (MS-XLS 2.5.209): row/column bounds and ordering.
        if row_first > row_last {
            return Err(invalid("CellWatch ref8.rwFirst exceeds ref8.rwLast"));
        }
        if column_first > column_last || column_last > MAX_COLUMN_INDEX {
            return Err(invalid(format!(
                "CellWatch ref8 columns {column_first:#06X}..{column_last:#06X} are invalid"
            )));
        }
        Ok(Self {
            flags,
            row_first,
            row_last,
            column_first,
            column_last,
            reserved: data[12..16].try_into().expect("length checked"),
        })
    }

    /// Serialize back to a complete `CellWatch` record payload.
    #[must_use]
    pub fn to_payload(&self) -> Vec<u8> {
        let mut payload = Vec::with_capacity(PAYLOAD_LEN);
        payload.extend_from_slice(&CELL_WATCH_RECORD_TYPE.to_le_bytes());
        payload.extend_from_slice(&self.flags.to_le_bytes());
        payload.extend_from_slice(&self.row_first.to_le_bytes());
        payload.extend_from_slice(&self.row_last.to_le_bytes());
        payload.extend_from_slice(&self.column_first.to_le_bytes());
        payload.extend_from_slice(&self.column_last.to_le_bytes());
        payload.extend_from_slice(&self.reserved);
        payload
    }

    /// Raw `grbitFrt` bitfield (`fFrtRef` set, `fFrtAlert` clear).
    #[must_use]
    pub fn flags(&self) -> u16 {
        self.flags
    }

    /// Zero-based index of the first row of the watched range (`ref8.rwFirst`).
    #[must_use]
    pub fn row_first(&self) -> u16 {
        self.row_first
    }

    /// Zero-based index of the last row of the watched range (`ref8.rwLast`).
    #[must_use]
    pub fn row_last(&self) -> u16 {
        self.row_last
    }

    /// Zero-based index of the first column of the watched range (`ref8.colFirst`).
    #[must_use]
    pub fn column_first(&self) -> u16 {
        self.column_first
    }

    /// Zero-based index of the last column of the watched range (`ref8.colLast`).
    #[must_use]
    pub fn column_last(&self) -> u16 {
        self.column_last
    }

    /// The preserved trailing `reserved` field.
    #[must_use]
    pub fn reserved(&self) -> [u8; 4] {
        self.reserved
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record(flags: u16, range: [u16; 4], reserved: [u8; 4]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&CELL_WATCH_RECORD_TYPE.to_le_bytes());
        data.extend_from_slice(&flags.to_le_bytes());
        for value in range {
            data.extend_from_slice(&value.to_le_bytes());
        }
        data.extend_from_slice(&reserved);
        data
    }

    #[test]
    fn round_trip() {
        let bytes = record(0x0001, [3, 3, 2, 2], [0; 4]);
        let parsed = CellWatch::parse(&bytes).unwrap();
        assert_eq!(parsed.row_first(), 3);
        assert_eq!(parsed.row_last(), 3);
        assert_eq!(parsed.column_first(), 2);
        assert_eq!(parsed.column_last(), 2);
        assert_eq!(parsed.flags(), 0x0001);
        assert_eq!(parsed.to_payload(), bytes);
    }

    #[test]
    fn preserves_reserved_bits_and_field() {
        // The 14 reserved grbitFrt bits and the 4 reserved bytes MUST be
        // ignored but round-trip verbatim.
        let bytes = record(0xFFFD, [0, 10, 0, 0xFF], [0xDE, 0xAD, 0xBE, 0xEF]);
        let parsed = CellWatch::parse(&bytes).unwrap();
        assert_eq!(parsed.flags(), 0xFFFD);
        assert_eq!(parsed.reserved(), [0xDE, 0xAD, 0xBE, 0xEF]);
        assert_eq!(parsed.to_payload(), bytes);
    }

    #[test]
    fn rejects_malformed_records() {
        let bytes = record(0x0001, [3, 3, 2, 2], [0; 4]);
        // Truncated and overlong payloads.
        assert!(CellWatch::parse(&bytes[..15]).is_err());
        assert!(CellWatch::parse(&[bytes.as_slice(), &[0]].concat()).is_err());
        // Wrong FrtRefHeaderU.rt.
        let mut wrong_rt = bytes.clone();
        wrong_rt[0..2].copy_from_slice(&0x0868u16.to_le_bytes());
        assert!(CellWatch::parse(&wrong_rt).is_err());
        // fFrtRef clear / fFrtAlert set.
        assert!(CellWatch::parse(&record(0x0000, [3, 3, 2, 2], [0; 4])).is_err());
        assert!(CellWatch::parse(&record(0x0003, [3, 3, 2, 2], [0; 4])).is_err());
        // Invalid Ref8U ranges.
        assert!(CellWatch::parse(&record(0x0001, [4, 3, 2, 2], [0; 4])).is_err());
        assert!(CellWatch::parse(&record(0x0001, [3, 3, 3, 2], [0; 4])).is_err());
        assert!(CellWatch::parse(&record(0x0001, [3, 3, 2, 0x0100], [0; 4])).is_err());
    }
}
