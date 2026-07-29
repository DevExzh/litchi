//! BIFF8 `SerParent` record (0x104A, MS-XLS 2.4.255) of the Chart Sheet
//! substream (MS-XLS 2.1): the series a trendline or error bar corresponds
//! to.
//!
//! Everything in this module is INERT: the index is stored verbatim and no
//! series reference is resolved.
//!
//! # References
//!
//! - MS-XLS 2.4.255 (SerParent)

use super::{XlsError, XlsResult};

/// Record type of the `SerParent` record (MS-XLS 2.4.255).
pub(crate) const SER_PARENT_RECORD_TYPE: u16 = 0x104A;

/// Byte length of a `SerParent` record payload: the `series` field.
const PAYLOAD_LEN: usize = 2;
/// Maximum `series` value (MS-XLS 2.4.255).
const MAX_SERIES_INDEX: u16 = 0x00FE;

/// Typed `SerParent` record content (MS-XLS 2.4.255): the one-based index of
/// the `Series` record associated with the current trendline or error bar.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsSerParent {
    /// One-based index into the Series records, in 0x0001..=0x00FE.
    series: u16,
}

impl XlsSerParent {
    /// Parse a `SerParent` record payload.
    pub fn parse(data: &[u8]) -> XlsResult<Self> {
        if data.len() != PAYLOAD_LEN {
            return Err(XlsError::InvalidLength {
                expected: PAYLOAD_LEN,
                found: data.len(),
            });
        }
        let series = u16::from_le_bytes([data[0], data[1]]);
        // MS-XLS 2.4.255: series MUST be in 0x0001..=0x00FE.
        if series == 0 || series > MAX_SERIES_INDEX {
            return Err(XlsError::InvalidRecord {
                record_type: SER_PARENT_RECORD_TYPE,
                message: format!(
                    "SerParent series {series:#06X} is outside 0x0001..={MAX_SERIES_INDEX:#06X}"
                ),
            });
        }
        Ok(Self { series })
    }

    /// Serialize back to a complete `SerParent` record payload.
    pub fn to_payload(&self) -> Vec<u8> {
        self.series.to_le_bytes().to_vec()
    }

    /// One-based index into the Series records (`series`).
    pub fn series(&self) -> u16 {
        self.series
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn round_trip() {
        for value in [0x0001u16, 0x0002, 0x00FE] {
            let payload = value.to_le_bytes();
            let parsed = XlsSerParent::parse(&payload).unwrap();
            assert_eq!(parsed.series(), value);
            assert_eq!(parsed.to_payload(), payload);
        }
    }

    #[test]
    fn rejects_malformed_records() {
        assert!(XlsSerParent::parse(&[0x01]).is_err());
        assert!(XlsSerParent::parse(&[0x01, 0x00, 0x00]).is_err());
        // series MUST be in 0x0001..=0x00FE.
        assert!(XlsSerParent::parse(&0x0000u16.to_le_bytes()).is_err());
        assert!(XlsSerParent::parse(&0x00FFu16.to_le_bytes()).is_err());
        assert!(XlsSerParent::parse(&0xFFFFu16.to_le_bytes()).is_err());
    }
}
