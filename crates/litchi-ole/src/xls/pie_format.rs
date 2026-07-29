//! BIFF8 `PieFormat` record (0x100B, MS-XLS 2.4.195) of the Chart Sheet
//! substream (MS-XLS 2.1): the distance of a data point or series from the
//! pie center, as a percentage.
//!
//! Everything in this module is INERT: the value is stored verbatim and no
//! chart geometry is computed. The chart-group restrictions of MS-XLS 2.4.195
//! (pie, doughnut, bar of pie, or pie of pie groups only) are cross-record
//! constraints the caller validates.
//!
//! # References
//!
//! - MS-XLS 2.4.195 (PieFormat)

use super::{XlsError, XlsResult};

/// Record type of the `PieFormat` record (MS-XLS 2.4.195).
pub(crate) const PIE_FORMAT_RECORD_TYPE: u16 = 0x100B;

/// Byte length of a `PieFormat` record payload: the `pcExplode` field.
const PAYLOAD_LEN: usize = 2;

/// Typed `PieFormat` record content (MS-XLS 2.4.195): the distance of a data
/// point or series from the pie center, as a percentage (`pcExplode`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsPieFormat {
    /// Distance from the center as a percentage. 0 means as close to the
    /// center as possible; 100 means at the chart-area edge; larger values
    /// scale the whole chart group down (MS-XLS 2.4.195). Guaranteed
    /// non-negative.
    explode_percent: u16,
}

impl XlsPieFormat {
    /// Parse a `PieFormat` record payload.
    pub fn parse(data: &[u8]) -> XlsResult<Self> {
        if data.len() != PAYLOAD_LEN {
            return Err(XlsError::InvalidLength {
                expected: PAYLOAD_LEN,
                found: data.len(),
            });
        }
        // MS-XLS 2.4.195: pcExplode MUST be greater than or equal to 0.
        let value = i16::from_le_bytes([data[0], data[1]]);
        if value < 0 {
            return Err(XlsError::InvalidRecord {
                record_type: PIE_FORMAT_RECORD_TYPE,
                message: format!("PieFormat pcExplode {value} is negative"),
            });
        }
        Ok(Self {
            explode_percent: u16::try_from(value).expect("non-negative"),
        })
    }

    /// Serialize back to a complete `PieFormat` record payload.
    pub fn to_payload(&self) -> Vec<u8> {
        self.explode_percent.to_le_bytes().to_vec()
    }

    /// Distance from the center as a percentage (`pcExplode`).
    pub fn explode_percent(&self) -> u16 {
        self.explode_percent
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn round_trip() {
        for value in [0u16, 25, 100, 400, 0x7FFF] {
            let payload = value.to_le_bytes();
            let parsed = XlsPieFormat::parse(&payload).unwrap();
            assert_eq!(parsed.explode_percent(), value);
            assert_eq!(parsed.to_payload(), payload);
        }
    }

    #[test]
    fn rejects_malformed_records() {
        assert!(XlsPieFormat::parse(&[0x00]).is_err());
        assert!(XlsPieFormat::parse(&[0x00, 0x00, 0x00]).is_err());
        // pcExplode MUST be greater than or equal to 0.
        assert!(XlsPieFormat::parse(&(-1i16).to_le_bytes()).is_err());
        assert!(XlsPieFormat::parse(&0x8000u16.to_le_bytes()).is_err());
    }
}
