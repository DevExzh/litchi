//! BIFF8 chart axis-group records of the Chart Sheet substream (MS-XLS 2.1):
//!
//! - **AxesUsed** (0x1046): the number of axis groups on the chart
//!   (MS-XLS 2.4.10).
//! - **AxisParent** (0x1041): properties of one axis group and the beginning
//!   of its record collection (MS-XLS 2.4.13).
//!
//! Everything in this module is INERT: values are stored verbatim and no
//! chart axes are constructed. The MS-XLS 2.4.13 rule that the first
//! `AxisParent` record is primary and the second secondary, and the MS-XLS
//! 2.4.10 rules tying `cAxes` to the presence of chart groups or a `Chart3d`
//! record, are cross-record constraints the caller validates.
//!
//! # References
//!
//! - MS-XLS 2.4.10 (AxesUsed), 2.4.13 (AxisParent), 2.5.14 (Boolean)

use super::{Error, Result};

/// Record type of the `AxesUsed` record (MS-XLS 2.4.10).
pub(crate) const AXES_USED_RECORD_TYPE: u16 = 0x1046;

/// Record type of the `AxisParent` record (MS-XLS 2.4.13).
pub(crate) const AXIS_PARENT_RECORD_TYPE: u16 = 0x1041;

/// Byte length of an `AxisParent` record payload: `iax` (2) + `unused` (16).
const AXIS_PARENT_LEN: usize = 18;

fn invalid(record_type: u16, message: impl Into<String>) -> Error {
    Error::InvalidRecord {
        record_type,
        message: message.into(),
    }
}

/// The `cAxes` axis-group count of an `AxesUsed` record (MS-XLS 2.4.10).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u16)]
pub enum AxesUsedCount {
    /// 0x0001: a single primary axis group is present.
    PrimaryOnly = 0x0001,
    /// 0x0002: both a primary and a secondary axis group are present.
    PrimaryAndSecondary = 0x0002,
}

impl AxesUsedCount {
    fn parse(value: u16) -> Result<Self> {
        match value {
            0x0001 => Ok(Self::PrimaryOnly),
            0x0002 => Ok(Self::PrimaryAndSecondary),
            other => Err(invalid(
                AXES_USED_RECORD_TYPE,
                format!("AxesUsed cAxes {other:#06X} is not a defined axis-group count"),
            )),
        }
    }
}

/// Typed `AxesUsed` record content (MS-XLS 2.4.10): the number of axis
/// groups on the chart.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct AxesUsed {
    /// The number of axis groups (`cAxes`).
    count: AxesUsedCount,
}

impl AxesUsed {
    /// Parse an `AxesUsed` record payload.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn parse(data: &[u8]) -> Result<Self> {
        if data.len() != 2 {
            return Err(Error::InvalidLength {
                expected: 2,
                found: data.len(),
            });
        }
        Ok(Self {
            count: AxesUsedCount::parse(u16::from_le_bytes([data[0], data[1]]))?,
        })
    }

    /// Serialize back to a complete `AxesUsed` record payload.
    #[must_use]
    pub fn to_payload(&self) -> Vec<u8> {
        (self.count as u16).to_le_bytes().to_vec()
    }

    /// The number of axis groups (`cAxes`).
    #[must_use]
    pub fn count(&self) -> AxesUsedCount {
        self.count
    }
}

/// The `iax` axis-group position of an `AxisParent` record (MS-XLS 2.4.13).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u16)]
pub enum AxisGroupPosition {
    /// 0x0000: the axis group is primary.
    Primary = 0x0000,
    /// 0x0001: the axis group is secondary.
    Secondary = 0x0001,
}

impl AxisGroupPosition {
    fn parse(value: u16) -> Result<Self> {
        // Boolean (MS-XLS 2.5.14): only 0x0000 and 0x0001 are legal.
        match value {
            0x0000 => Ok(Self::Primary),
            0x0001 => Ok(Self::Secondary),
            other => Err(invalid(
                AXIS_PARENT_RECORD_TYPE,
                format!("AxisParent iax {other:#06X} is not a Boolean"),
            )),
        }
    }
}

/// Typed `AxisParent` record content (MS-XLS 2.4.13): properties of an axis
/// group.
///
/// The 16 `unused` bytes are undefined and MUST be ignored; they are
/// preserved verbatim so the record round-trips unchanged.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct AxisParent {
    /// Whether the axis group is primary or secondary (`iax`).
    position: AxisGroupPosition,
    /// The undefined `unused` bytes, preserved verbatim.
    unused: [u8; 16],
}

impl AxisParent {
    /// Parse an `AxisParent` record payload.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    /// # Panics
    ///
    /// Panics only if an internal BIFF invariant has been violated.
    pub fn parse(data: &[u8]) -> Result<Self> {
        if data.len() != AXIS_PARENT_LEN {
            return Err(Error::InvalidLength {
                expected: AXIS_PARENT_LEN,
                found: data.len(),
            });
        }
        Ok(Self {
            position: AxisGroupPosition::parse(u16::from_le_bytes([data[0], data[1]]))?,
            unused: data[2..18].try_into().expect("length checked"),
        })
    }

    /// Serialize back to a complete `AxisParent` record payload.
    #[must_use]
    pub fn to_payload(&self) -> Vec<u8> {
        let mut payload = Vec::with_capacity(AXIS_PARENT_LEN);
        payload.extend_from_slice(&(self.position as u16).to_le_bytes());
        payload.extend_from_slice(&self.unused);
        payload
    }

    /// Whether the axis group is primary or secondary (`iax`).
    #[must_use]
    pub fn position(&self) -> AxisGroupPosition {
        self.position
    }

    /// The preserved undefined `unused` bytes.
    #[must_use]
    pub fn unused(&self) -> [u8; 16] {
        self.unused
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn axes_used_round_trip() {
        for (value, expected) in [
            (0x0001u16, AxesUsedCount::PrimaryOnly),
            (0x0002, AxesUsedCount::PrimaryAndSecondary),
        ] {
            let payload = value.to_le_bytes();
            let parsed = AxesUsed::parse(&payload).unwrap();
            assert_eq!(parsed.count(), expected);
            assert_eq!(parsed.to_payload(), payload);
        }
    }

    #[test]
    fn axes_used_rejects_malformed_records() {
        assert!(AxesUsed::parse(&[0x01]).is_err());
        assert!(AxesUsed::parse(&[0x01, 0x00, 0x00]).is_err());
        assert!(AxesUsed::parse(&0x0000u16.to_le_bytes()).is_err());
        assert!(AxesUsed::parse(&0x0003u16.to_le_bytes()).is_err());
    }

    #[test]
    fn axis_parent_round_trip() {
        for (value, expected) in [
            (0x0000u16, AxisGroupPosition::Primary),
            (0x0001, AxisGroupPosition::Secondary),
        ] {
            let mut payload = value.to_le_bytes().to_vec();
            payload.extend_from_slice(&[0; 16]);
            let parsed = AxisParent::parse(&payload).unwrap();
            assert_eq!(parsed.position(), expected);
            assert_eq!(parsed.unused(), [0; 16]);
            assert_eq!(parsed.to_payload(), payload);
        }
    }

    #[test]
    fn axis_parent_preserves_unused_bytes() {
        // The 16 unused bytes are undefined and MUST be ignored; they
        // round-trip verbatim.
        let mut payload = 0x0001u16.to_le_bytes().to_vec();
        payload.extend_from_slice(&[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16]);
        let parsed = AxisParent::parse(&payload).unwrap();
        assert_eq!(parsed.unused()[15], 16);
        assert_eq!(parsed.to_payload(), payload);
    }

    #[test]
    fn axis_parent_rejects_malformed_records() {
        let mut payload = 0x0000u16.to_le_bytes().to_vec();
        payload.extend_from_slice(&[0; 16]);
        assert!(AxisParent::parse(&payload[..17]).is_err());
        assert!(AxisParent::parse(&[payload.as_slice(), &[0]].concat()).is_err());
        // iax is a Boolean (MS-XLS 2.5.14).
        let mut bad = payload.clone();
        bad[0..2].copy_from_slice(&0x0002u16.to_le_bytes());
        assert!(AxisParent::parse(&bad).is_err());
    }
}
