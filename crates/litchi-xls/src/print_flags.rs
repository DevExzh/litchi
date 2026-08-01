//! BIFF8 worksheet print display flags of the worksheet substream
//! (MS-XLS 2.1):
//!
//! - **PrintRowCol** (0x002A): whether the row and column headers are
//!   printed (MS-XLS 2.4.203).
//! - **GridSet** (0x0082): whether the gridlines are printed
//!   (MS-XLS 2.4.132).
//!
//! Everything in this module is INERT: the flags are stored verbatim and no
//! print layout is applied.
//!
//! # References
//!
//! - MS-XLS 2.4.132 (GridSet), 2.4.203 (PrintRowCol), 2.5.14 (Boolean)

use super::{XlsError, XlsResult};

/// Record type of the `PrintRowCol` record (MS-XLS 2.4.203).
pub(crate) const PRINT_ROW_COL_RECORD_TYPE: u16 = 0x002A;

/// Byte length of a `PrintRowCol` or `GridSet` (record type 0x0082) record
/// payload: a single two-byte field (MS-XLS 2.4.132 / 2.4.203).
const PAYLOAD_LEN: usize = 2;

/// `GridSet` bitfield: `fPrintGrid` (MS-XLS 2.4.132).
const GRID_SET_PRINT_GRID: u16 = 0x0001;

fn invalid(record_type: u16, message: impl Into<String>) -> XlsError {
    XlsError::InvalidRecord {
        record_type,
        message: message.into(),
    }
}

/// Typed `PrintRowCol` record content (MS-XLS 2.4.203): whether the row and
/// column headers are printed.
///
/// The value table in MS-XLS 2.4.203 lists both 0x0000 and 0x0001 as "not
/// printed"; the second entry is a specification typo. This reader applies
/// the `Boolean` semantics of MS-XLS 2.5.14, where 0x0001 means the headers
/// ARE printed.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsPrintRowCol {
    /// Whether the row and column headers are printed (`printRwCol`).
    print_headers: bool,
}

impl XlsPrintRowCol {
    /// Parse a `PrintRowCol` record payload.
    pub fn parse(data: &[u8]) -> XlsResult<Self> {
        if data.len() != PAYLOAD_LEN {
            return Err(XlsError::InvalidLength {
                expected: PAYLOAD_LEN,
                found: data.len(),
            });
        }
        let value = u16::from_le_bytes(data[0..2].try_into().expect("length checked"));
        // Boolean (MS-XLS 2.5.14): only 0x0000 and 0x0001 are legal.
        match value {
            0x0000 => Ok(Self {
                print_headers: false,
            }),
            0x0001 => Ok(Self {
                print_headers: true,
            }),
            other => Err(invalid(
                PRINT_ROW_COL_RECORD_TYPE,
                format!("PrintRowCol printRwCol {other:#06X} is not a Boolean"),
            )),
        }
    }

    /// Serialize back to a complete `PrintRowCol` record payload.
    pub fn to_payload(&self) -> Vec<u8> {
        u16::from(self.print_headers).to_le_bytes().to_vec()
    }

    /// Whether the row and column headers are printed.
    pub fn print_headers(&self) -> bool {
        self.print_headers
    }
}

/// Typed `GridSet` record content (MS-XLS 2.4.132): whether the gridlines
/// are printed.
///
/// The 15 `unused` bits are undefined and MUST be ignored; they are
/// preserved verbatim so the record round-trips unchanged.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsGridSet {
    /// Raw two-byte bitfield, including the undefined `unused` bits.
    flags: u16,
}

impl XlsGridSet {
    /// Parse a `GridSet` record payload.
    pub fn parse(data: &[u8]) -> XlsResult<Self> {
        if data.len() != PAYLOAD_LEN {
            return Err(XlsError::InvalidLength {
                expected: PAYLOAD_LEN,
                found: data.len(),
            });
        }
        let flags = u16::from_le_bytes(data[0..2].try_into().expect("length checked"));
        Ok(Self { flags })
    }

    /// Serialize back to a complete `GridSet` record payload.
    pub fn to_payload(&self) -> Vec<u8> {
        self.flags.to_le_bytes().to_vec()
    }

    /// Whether the gridlines are printed (`fPrintGrid`).
    pub fn print_grid(&self) -> bool {
        self.flags & GRID_SET_PRINT_GRID != 0
    }

    /// Raw bitfield value, including the undefined `unused` bits.
    pub fn flags(&self) -> u16 {
        self.flags
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn print_row_col_round_trip() {
        for (payload, expected) in [([0x00, 0x00], false), ([0x01, 0x00], true)] {
            let record = XlsPrintRowCol::parse(&payload).unwrap();
            assert_eq!(record.print_headers(), expected);
            assert_eq!(record.to_payload(), payload);
        }
    }

    #[test]
    fn print_row_col_rejects_bad_length_and_non_boolean() {
        assert!(XlsPrintRowCol::parse(&[0x01]).is_err());
        assert!(XlsPrintRowCol::parse(&[0x00, 0x00, 0x00]).is_err());
        // Boolean (MS-XLS 2.5.14) allows only 0x0000 and 0x0001.
        assert!(XlsPrintRowCol::parse(&[0x02, 0x00]).is_err());
        assert!(XlsPrintRowCol::parse(&[0x00, 0x01]).is_err());
    }

    #[test]
    fn grid_set_round_trip() {
        let payload = [0x01, 0x00];
        let record = XlsGridSet::parse(&payload).unwrap();
        assert!(record.print_grid());
        assert_eq!(record.flags(), 0x0001);
        assert_eq!(record.to_payload(), payload);

        let payload = [0x00, 0x00];
        let record = XlsGridSet::parse(&payload).unwrap();
        assert!(!record.print_grid());
        assert_eq!(record.to_payload(), payload);
    }

    #[test]
    fn grid_set_preserves_unused_bits() {
        // The 15 unused bits are undefined and MUST be ignored, but the
        // record must round-trip unchanged.
        let payload = [0xFE, 0x7F];
        let record = XlsGridSet::parse(&payload).unwrap();
        assert!(!record.print_grid());
        assert_eq!(record.to_payload(), payload);
    }

    #[test]
    fn grid_set_rejects_bad_length() {
        assert!(XlsGridSet::parse(&[0x01]).is_err());
        assert!(XlsGridSet::parse(&[0x01, 0x00, 0x00]).is_err());
    }
}
