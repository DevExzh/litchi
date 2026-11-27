//! BIFF8 `SerAuxErrBar` record (0x105B, MS-XLS 2.4.249) of the Chart Sheet
//! substream (MS-XLS 2.1): properties of an error bar.
//!
//! Everything in this module is INERT: values are stored verbatim and no
//! error bars are rendered. The `cnum` count and the value-source rules of
//! MS-XLS 2.4.249 depend on the preceding `BRAI` and `Number` records; those
//! cross-record constraints are documented here, not enforced by the record
//! reader.
//!
//! # References
//!
//! - MS-XLS 2.4.249 (SerAuxErrBar), 2.5.14 (Boolean), 2.5.342 (Xnum)

use super::{Error, Result};

/// Record type of the `SerAuxErrBar` record (MS-XLS 2.4.249).
pub(crate) const SER_AUX_ERR_BAR_RECORD_TYPE: u16 = 0x105B;

/// Byte length of a `SerAuxErrBar` record payload.
const PAYLOAD_LEN: usize = 14;

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidRecord {
        record_type: SER_AUX_ERR_BAR_RECORD_TYPE,
        message: message.into(),
    }
}

/// The `sertm` direction of the error bars (MS-XLS 2.4.249).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u8)]
pub enum ErrorBarDirection {
    /// 0x01: horizontal in the plus direction.
    HorizontalPlus = 0x01,
    /// 0x02: horizontal in the minus direction.
    HorizontalMinus = 0x02,
    /// 0x03: vertical in the plus direction.
    VerticalPlus = 0x03,
    /// 0x04: vertical in the minus direction.
    VerticalMinus = 0x04,
}

impl ErrorBarDirection {
    fn parse(value: u8) -> Result<Self> {
        match value {
            0x01 => Ok(Self::HorizontalPlus),
            0x02 => Ok(Self::HorizontalMinus),
            0x03 => Ok(Self::VerticalPlus),
            0x04 => Ok(Self::VerticalMinus),
            other => Err(invalid(format!(
                "SerAuxErrBar sertm {other:#04X} is not a defined direction"
            ))),
        }
    }
}

/// The `ebsrc` error amount type of the error bars (MS-XLS 2.4.249).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u8)]
pub enum ErrorBarSource {
    /// 0x01: percentage.
    Percentage = 0x01,
    /// 0x02: fixed value.
    FixedValue = 0x02,
    /// 0x03: standard deviation.
    StandardDeviation = 0x03,
    /// 0x04: custom values (array of values or range).
    Custom = 0x04,
    /// 0x05: standard error.
    StandardError = 0x05,
}

impl ErrorBarSource {
    fn parse(value: u8) -> Result<Self> {
        match value {
            0x01 => Ok(Self::Percentage),
            0x02 => Ok(Self::FixedValue),
            0x03 => Ok(Self::StandardDeviation),
            0x04 => Ok(Self::Custom),
            0x05 => Ok(Self::StandardError),
            other => Err(invalid(format!(
                "SerAuxErrBar ebsrc {other:#04X} is not a defined error amount type"
            ))),
        }
    }
}

/// Typed `SerAuxErrBar` record content (MS-XLS 2.4.249): properties of an
/// error bar.
///
/// The `reserved` byte (MUST be 0x01 and MUST be ignored) is preserved
/// verbatim so the record round-trips unchanged. `numValue` is preserved even
/// when `ebsrc` is `Custom` or `StandardError` (where it MUST be ignored),
/// and `cnum` is preserved when `ebsrc` is not `Custom`.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct SerAuxErrBar {
    /// Direction of the error bars (`sertm`).
    direction: ErrorBarDirection,
    /// Error amount type (`ebsrc`).
    source: ErrorBarSource,
    /// Whether the error bars are T-shaped (`fTeeTop`).
    tee_top: bool,
    /// The `reserved` byte, preserved verbatim.
    reserved: u8,
    /// Fixed value, percentage, or number of standard deviations (`numValue`).
    value: f64,
    /// Number of value or cell references for custom error bars (`cnum`).
    value_count: u16,
}

impl SerAuxErrBar {
    /// Parse a `SerAuxErrBar` record payload.
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
        // Boolean (MS-XLS 2.5.14): only 0x00 and 0x01 are legal.
        let tee_top = match data[2] {
            0x00 => false,
            0x01 => true,
            other => {
                return Err(invalid(format!(
                    "SerAuxErrBar fTeeTop {other:#04X} is not a Boolean"
                )));
            },
        };
        Ok(Self {
            direction: ErrorBarDirection::parse(data[0])?,
            source: ErrorBarSource::parse(data[1])?,
            tee_top,
            reserved: data[3],
            value: f64::from_le_bytes(data[4..12].try_into().expect("length checked")),
            value_count: u16::from_le_bytes([data[12], data[13]]),
        })
    }

    /// Serialize back to a complete `SerAuxErrBar` record payload.
    #[must_use]
    pub fn to_payload(&self) -> Vec<u8> {
        let mut payload = Vec::with_capacity(PAYLOAD_LEN);
        payload.push(self.direction as u8);
        payload.push(self.source as u8);
        payload.push(u8::from(self.tee_top));
        payload.push(self.reserved);
        payload.extend_from_slice(&self.value.to_le_bytes());
        payload.extend_from_slice(&self.value_count.to_le_bytes());
        payload
    }

    /// Direction of the error bars (`sertm`).
    #[must_use]
    pub fn direction(&self) -> ErrorBarDirection {
        self.direction
    }

    /// Error amount type (`ebsrc`).
    #[must_use]
    pub fn source(&self) -> ErrorBarSource {
        self.source
    }

    /// Whether the error bars are T-shaped (`fTeeTop`).
    #[must_use]
    pub fn tee_top(&self) -> bool {
        self.tee_top
    }

    /// The preserved `reserved` byte.
    #[must_use]
    pub fn reserved(&self) -> u8 {
        self.reserved
    }

    /// Fixed value, percentage, or number of standard deviations (`numValue`);
    /// preserved verbatim when `ebsrc` is `Custom` or `StandardError`.
    #[must_use]
    pub fn value(&self) -> f64 {
        self.value
    }

    /// Number of value or cell references for custom error bars (`cnum`);
    /// preserved verbatim when `ebsrc` is not `Custom`.
    #[must_use]
    pub fn value_count(&self) -> u16 {
        self.value_count
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record(direction: u8, source: u8, tee_top: u8, value: f64, count: u16) -> Vec<u8> {
        let mut data = vec![direction, source, tee_top, 0x01];
        data.extend_from_slice(&value.to_le_bytes());
        data.extend_from_slice(&count.to_le_bytes());
        data
    }

    #[test]
    fn round_trip_all_enums() {
        for (direction, expected_direction) in [
            (0x01, ErrorBarDirection::HorizontalPlus),
            (0x02, ErrorBarDirection::HorizontalMinus),
            (0x03, ErrorBarDirection::VerticalPlus),
            (0x04, ErrorBarDirection::VerticalMinus),
        ] {
            for (source, expected_source) in [
                (0x01, ErrorBarSource::Percentage),
                (0x02, ErrorBarSource::FixedValue),
                (0x03, ErrorBarSource::StandardDeviation),
                (0x04, ErrorBarSource::Custom),
                (0x05, ErrorBarSource::StandardError),
            ] {
                let bytes = record(direction, source, 0x01, 2.5, 7);
                let parsed = SerAuxErrBar::parse(&bytes).unwrap();
                assert_eq!(parsed.direction(), expected_direction);
                assert_eq!(parsed.source(), expected_source);
                assert!(parsed.tee_top());
                assert_eq!(parsed.value(), 2.5);
                assert_eq!(parsed.value_count(), 7);
                assert_eq!(parsed.to_payload(), bytes);
            }
        }
    }

    #[test]
    fn preserves_reserved_byte_and_ignored_fields() {
        // The reserved byte MUST be 0x01 and MUST be ignored, and numValue /
        // cnum are preserved when ignored for the ebsrc value.
        let mut bytes = record(0x03, 0x05, 0x00, 1.0, 42);
        bytes[3] = 0xAA;
        let parsed = SerAuxErrBar::parse(&bytes).unwrap();
        assert_eq!(parsed.reserved(), 0xAA);
        assert!(!parsed.tee_top());
        assert_eq!(parsed.value(), 1.0);
        assert_eq!(parsed.value_count(), 42);
        assert_eq!(parsed.to_payload(), bytes);
    }

    #[test]
    fn rejects_malformed_records() {
        let bytes = record(0x01, 0x01, 0x01, 0.0, 0);
        // Truncated and overlong payloads.
        assert!(SerAuxErrBar::parse(&bytes[..13]).is_err());
        assert!(SerAuxErrBar::parse(&[bytes.as_slice(), &[0]].concat()).is_err());
        // Undefined sertm / ebsrc / fTeeTop values.
        assert!(SerAuxErrBar::parse(&record(0x00, 0x01, 0x01, 0.0, 0)).is_err());
        assert!(SerAuxErrBar::parse(&record(0x05, 0x01, 0x01, 0.0, 0)).is_err());
        assert!(SerAuxErrBar::parse(&record(0x01, 0x00, 0x01, 0.0, 0)).is_err());
        assert!(SerAuxErrBar::parse(&record(0x01, 0x06, 0x01, 0.0, 0)).is_err());
        assert!(SerAuxErrBar::parse(&record(0x01, 0x01, 0x02, 0.0, 0)).is_err());
    }
}
