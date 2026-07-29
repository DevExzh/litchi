//! BIFF8 `PlotGrowth` record (0x1064, MS-XLS 2.4.198) of the Chart Sheet
//! substream (MS-XLS 2.1): the plot area scale factors for font scaling.
//!
//! Everything in this module is INERT: the scale factors are stored verbatim
//! and the font scaling algorithm of MS-XLS 2.4.109 (`Fbi`) is not applied.
//! The record is unused and MUST be ignored when no `Fbi` record with
//! `scab` 0x0001 exists (MS-XLS 2.4.198); that cross-record constraint is
//! documented here, not enforced by the record reader.
//!
//! # References
//!
//! - MS-XLS 2.4.198 (PlotGrowth), MS-OSHARED 2.2.1.6 (FixedPoint)

use super::{XlsError, XlsResult};

/// Byte length of a `PlotGrowth` record payload.
const PAYLOAD_LEN: usize = 8;

/// A FixedPoint value (MS-OSHARED 2.2.1.6): a signed 16.16 fixed-point
/// number stored as a 32-bit integer.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsFixedPoint {
    /// Raw 16.16 fixed-point value.
    raw: i32,
}

impl XlsFixedPoint {
    /// Wrap a raw 16.16 fixed-point value.
    pub const fn from_raw(raw: i32) -> Self {
        Self { raw }
    }

    /// The raw 16.16 fixed-point value.
    pub const fn raw(self) -> i32 {
        self.raw
    }

    /// The value as a floating-point number (`raw` / 65536).
    pub fn to_f64(self) -> f64 {
        f64::from(self.raw) / 65536.0
    }
}

/// Typed `PlotGrowth` record content (MS-XLS 2.4.198): the horizontal and
/// vertical growth of the plot area for font scaling.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsPlotGrowth {
    /// Horizontal growth of the plot area, in points (`dxPlotGrowth`).
    dx: XlsFixedPoint,
    /// Vertical growth of the plot area, in points (`dyPlotGrowth`).
    dy: XlsFixedPoint,
}

impl XlsPlotGrowth {
    /// Parse a `PlotGrowth` record payload.
    pub fn parse(data: &[u8]) -> XlsResult<Self> {
        if data.len() != PAYLOAD_LEN {
            return Err(XlsError::InvalidLength {
                expected: PAYLOAD_LEN,
                found: data.len(),
            });
        }
        Ok(Self {
            dx: XlsFixedPoint::from_raw(i32::from_le_bytes(data[0..4].try_into().expect("checked"))),
            dy: XlsFixedPoint::from_raw(i32::from_le_bytes(data[4..8].try_into().expect("checked"))),
        })
    }

    /// Serialize back to a complete `PlotGrowth` record payload.
    pub fn to_payload(&self) -> Vec<u8> {
        let mut payload = Vec::with_capacity(PAYLOAD_LEN);
        payload.extend_from_slice(&self.dx.raw().to_le_bytes());
        payload.extend_from_slice(&self.dy.raw().to_le_bytes());
        payload
    }

    /// Horizontal growth of the plot area (`dxPlotGrowth`).
    pub fn dx(&self) -> XlsFixedPoint {
        self.dx
    }

    /// Vertical growth of the plot area (`dyPlotGrowth`).
    pub fn dy(&self) -> XlsFixedPoint {
        self.dy
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn round_trip_and_fixed_point_decode() {
        let mut payload = 0x0001_0000i32.to_le_bytes().to_vec();
        payload.extend_from_slice(&0x0000_8000i32.to_le_bytes());
        let parsed = XlsPlotGrowth::parse(&payload).unwrap();
        assert_eq!(parsed.dx().raw(), 0x0001_0000);
        assert_eq!(parsed.dx().to_f64(), 1.0);
        assert_eq!(parsed.dy().to_f64(), 0.5);
        assert_eq!(parsed.to_payload(), payload);
    }

    #[test]
    fn negative_growth_round_trip() {
        let mut payload = (-65536i32).to_le_bytes().to_vec();
        payload.extend_from_slice(&(-1i32).to_le_bytes());
        let parsed = XlsPlotGrowth::parse(&payload).unwrap();
        assert_eq!(parsed.dx().to_f64(), -1.0);
        assert_eq!(parsed.to_payload(), payload);
    }

    #[test]
    fn rejects_bad_length() {
        assert!(XlsPlotGrowth::parse(&[0; 7]).is_err());
        assert!(XlsPlotGrowth::parse(&[0; 9]).is_err());
    }
}
