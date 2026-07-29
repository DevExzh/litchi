//! BIFF8 `Scl` record (0x00A0, MS-XLS 2.4.247) of the worksheet substream
//! (MS-XLS 2.1): the zoom level of the current view as a fraction
//! `nscl`/`dscl`.
//!
//! Everything in this module is INERT: the fraction is stored verbatim and
//! no view zoom is applied. The record MUST exist when the zoom level is not
//! 1 (MS-XLS 2.4.247); that is a writer-side constraint, not a reader one.
//!
//! # References
//!
//! - MS-XLS 2.4.247 (Scl)

use super::{XlsError, XlsResult};

/// Record type of the `Scl` record (MS-XLS 2.4.247).
pub(crate) const SCL_RECORD_TYPE: u16 = 0x00A0;

/// Byte length of an `Scl` record payload: `nscl` + `dscl`.
const PAYLOAD_LEN: usize = 4;

fn invalid(message: impl Into<String>) -> XlsError {
    XlsError::InvalidRecord {
        record_type: SCL_RECORD_TYPE,
        message: message.into(),
    }
}

/// Typed `Scl` record content (MS-XLS 2.4.247): the zoom level of the
/// current view as the fraction `nscl`/`dscl`, which MUST be in
/// 1/10..=4 (validated exactly with integer arithmetic).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsScl {
    /// Numerator of the zoom fraction (`nscl`), greater than or equal to 1.
    numerator: i16,
    /// Denominator of the zoom fraction (`dscl`), greater than or equal to 1.
    denominator: i16,
}

impl XlsScl {
    /// Parse an `Scl` record payload.
    pub fn parse(data: &[u8]) -> XlsResult<Self> {
        if data.len() != PAYLOAD_LEN {
            return Err(XlsError::InvalidLength {
                expected: PAYLOAD_LEN,
                found: data.len(),
            });
        }
        let numerator = i16::from_le_bytes([data[0], data[1]]);
        let denominator = i16::from_le_bytes([data[2], data[3]]);
        // MS-XLS 2.4.247: both fields MUST be greater than or equal to 1.
        if numerator < 1 {
            return Err(invalid(format!("Scl nscl {numerator} is less than 1")));
        }
        if denominator < 1 {
            return Err(invalid(format!("Scl dscl {denominator} is less than 1")));
        }
        // MS-XLS 2.4.247: the fraction MUST be in 1/10..=4, checked exactly
        // with integer arithmetic (10*nscl >= dscl and nscl <= 4*dscl).
        let n = i32::from(numerator);
        let d = i32::from(denominator);
        if 10 * n < d || n > 4 * d {
            return Err(invalid(format!(
                "Scl fraction {numerator}/{denominator} is outside 1/10..=4"
            )));
        }
        Ok(Self {
            numerator,
            denominator,
        })
    }

    /// Serialize back to a complete `Scl` record payload.
    pub fn to_payload(&self) -> Vec<u8> {
        let mut payload = Vec::with_capacity(PAYLOAD_LEN);
        payload.extend_from_slice(&self.numerator.to_le_bytes());
        payload.extend_from_slice(&self.denominator.to_le_bytes());
        payload
    }

    /// Numerator of the zoom fraction (`nscl`).
    pub fn numerator(&self) -> i16 {
        self.numerator
    }

    /// Denominator of the zoom fraction (`dscl`).
    pub fn denominator(&self) -> i16 {
        self.denominator
    }

    /// The zoom level as a floating-point value (`nscl`/`dscl`).
    pub fn zoom(&self) -> f64 {
        f64::from(self.numerator) / f64::from(self.denominator)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn round_trip() {
        for (n, d) in [(1i16, 1i16), (1, 10), (4, 1), (3, 10), (1, 2)] {
            let mut payload = n.to_le_bytes().to_vec();
            payload.extend_from_slice(&d.to_le_bytes());
            let parsed = XlsScl::parse(&payload).unwrap();
            assert_eq!(parsed.numerator(), n);
            assert_eq!(parsed.denominator(), d);
            assert_eq!(parsed.to_payload(), payload);
        }
        assert_eq!(XlsScl::parse(&[1, 0, 10, 0]).unwrap().zoom(), 0.1);
        assert_eq!(XlsScl::parse(&[4, 0, 1, 0]).unwrap().zoom(), 4.0);
    }

    #[test]
    fn rejects_malformed_records() {
        // Bad length.
        assert!(XlsScl::parse(&[1, 0, 1]).is_err());
        assert!(XlsScl::parse(&[1, 0, 1, 0, 0]).is_err());
        // nscl / dscl less than 1.
        assert!(XlsScl::parse(&[0, 0, 1, 0]).is_err());
        assert!(XlsScl::parse(&[1, 0, 0, 0]).is_err());
        let mut negative = (-1i16).to_le_bytes().to_vec();
        negative.extend_from_slice(&1i16.to_le_bytes());
        assert!(XlsScl::parse(&negative).is_err());
        // Fraction outside 1/10..=4.
        assert!(XlsScl::parse(&[1, 0, 11, 0]).is_err()); // 1/11 < 1/10
        assert!(XlsScl::parse(&[5, 0, 1, 0]).is_err()); // 5/1 > 4
        // Exact boundaries are legal.
        assert!(XlsScl::parse(&[1, 0, 10, 0]).is_ok());
        assert!(XlsScl::parse(&[4, 0, 1, 0]).is_ok());
    }
}
