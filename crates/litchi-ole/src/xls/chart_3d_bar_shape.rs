//! BIFF8 `Chart3DBarShape` record (0x085F, MS-XLS 2.4.47) of the Chart Sheet
//! substream (MS-XLS 2.1): the shape of the data points in a bar or column
//! chart group.
//!
//! Everything in this module is INERT: the shape values are stored verbatim
//! and no chart geometry is computed. The record only applies to bar/column
//! chart groups and MUST be ignored for all other chart groups and when the
//! substream has no `Chart3d` record (MS-XLS 2.4.47); those are cross-record
//! constraints the caller validates.
//!
//! # References
//!
//! - MS-XLS 2.4.47 (Chart3DBarShape), 2.5.14 (Boolean)

use super::{XlsError, XlsResult};

/// Record type of the `Chart3DBarShape` record (MS-XLS 2.4.47).
pub(crate) const CHART_3D_BAR_SHAPE_RECORD_TYPE: u16 = 0x085F;

/// Byte length of a `Chart3DBarShape` record payload: `riser` + `taper`.
const PAYLOAD_LEN: usize = 2;

fn invalid(message: impl Into<String>) -> XlsError {
    XlsError::InvalidRecord {
        record_type: CHART_3D_BAR_SHAPE_RECORD_TYPE,
        message: message.into(),
    }
}

/// The `riser` base shape of the data points (MS-XLS 2.4.47).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u8)]
pub enum XlsChart3DRiserShape {
    /// 0x00: the base of the data point is a rectangle.
    Rectangle = 0x00,
    /// 0x01: the base of the data point is an ellipse.
    Ellipse = 0x01,
}

impl XlsChart3DRiserShape {
    fn parse(value: u8) -> XlsResult<Self> {
        // Boolean (MS-XLS 2.5.14): only 0x00 and 0x01 are legal.
        match value {
            0x00 => Ok(Self::Rectangle),
            0x01 => Ok(Self::Ellipse),
            other => Err(invalid(format!(
                "Chart3DBarShape riser {other:#04X} is not a Boolean"
            ))),
        }
    }
}

/// The `taper` of the data points from base to tip (MS-XLS 2.4.47).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u8)]
pub enum XlsChart3DTaper {
    /// 0x00: the data points do not taper.
    None = 0x00,
    /// 0x01: the data points taper to a point at their maximum value.
    ToPoint = 0x01,
    /// 0x02: the data points taper towards the projected point at the maximum
    /// value of the chart group, clipped at each data point's value.
    ClippedAtValue = 0x02,
}

impl XlsChart3DTaper {
    fn parse(value: u8) -> XlsResult<Self> {
        match value {
            0x00 => Ok(Self::None),
            0x01 => Ok(Self::ToPoint),
            0x02 => Ok(Self::ClippedAtValue),
            other => Err(invalid(format!(
                "Chart3DBarShape taper {other:#04X} is not a defined taper"
            ))),
        }
    }
}

/// Typed `Chart3DBarShape` record content (MS-XLS 2.4.47): the shape of the
/// data points in a bar or column chart group.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsChart3DBarShape {
    /// The base shape of the data points (`riser`).
    riser: XlsChart3DRiserShape,
    /// The taper of the data points (`taper`).
    taper: XlsChart3DTaper,
}

impl XlsChart3DBarShape {
    /// Parse a `Chart3DBarShape` record payload.
    pub fn parse(data: &[u8]) -> XlsResult<Self> {
        if data.len() != PAYLOAD_LEN {
            return Err(XlsError::InvalidLength {
                expected: PAYLOAD_LEN,
                found: data.len(),
            });
        }
        Ok(Self {
            riser: XlsChart3DRiserShape::parse(data[0])?,
            taper: XlsChart3DTaper::parse(data[1])?,
        })
    }

    /// Serialize back to a complete `Chart3DBarShape` record payload.
    pub fn to_payload(&self) -> Vec<u8> {
        vec![self.riser as u8, self.taper as u8]
    }

    /// The base shape of the data points (`riser`).
    pub fn riser(&self) -> XlsChart3DRiserShape {
        self.riser
    }

    /// The taper of the data points (`taper`).
    pub fn taper(&self) -> XlsChart3DTaper {
        self.taper
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn round_trip_all_combinations() {
        for (riser, riser_value) in [
            (XlsChart3DRiserShape::Rectangle, 0x00),
            (XlsChart3DRiserShape::Ellipse, 0x01),
        ] {
            for (taper, taper_value) in [
                (XlsChart3DTaper::None, 0x00),
                (XlsChart3DTaper::ToPoint, 0x01),
                (XlsChart3DTaper::ClippedAtValue, 0x02),
            ] {
                let payload = [riser_value, taper_value];
                let parsed = XlsChart3DBarShape::parse(&payload).unwrap();
                assert_eq!(parsed.riser(), riser);
                assert_eq!(parsed.taper(), taper);
                assert_eq!(parsed.to_payload(), payload);
            }
        }
    }

    #[test]
    fn rejects_malformed_records() {
        // Bad length.
        assert!(XlsChart3DBarShape::parse(&[0x00]).is_err());
        assert!(XlsChart3DBarShape::parse(&[0x00, 0x00, 0x00]).is_err());
        // riser is a Boolean (MS-XLS 2.5.14).
        assert!(XlsChart3DBarShape::parse(&[0x02, 0x00]).is_err());
        // taper is one of 0x00..=0x02.
        assert!(XlsChart3DBarShape::parse(&[0x00, 0x03]).is_err());
        assert!(XlsChart3DBarShape::parse(&[0x00, 0xFF]).is_err());
    }
}
