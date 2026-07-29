//! BIFF8 `SerFmt` record (0x105D, MS-XLS 2.4.251) of the Chart Sheet
//! substream (MS-XLS 2.1): properties of the data points, data markers, or
//! lines of a series.
//!
//! Everything in this module is INERT: the flags are stored verbatim and no
//! series formatting is applied.
//!
//! # References
//!
//! - MS-XLS 2.4.251 (SerFmt)

use super::{XlsError, XlsResult};

/// Byte length of a `SerFmt` record payload: the flags word.
const PAYLOAD_LEN: usize = 2;
/// Flags bit: `fSmoothedLine` (smooth line effect).
const FLAG_SMOOTHED_LINE: u16 = 0x0001;
/// Flags bit: `f3DBubbles` (3-D bubble effect).
const FLAG_3D_BUBBLES: u16 = 0x0002;
/// Flags bit: `fArShadow` (data markers displayed with a shadow).
const FLAG_SHADOW: u16 = 0x0004;

/// Typed `SerFmt` record content (MS-XLS 2.4.251): properties of the
/// associated data points, data markers, or lines of the series.
///
/// The 13 `reserved` bits (MUST be zero, and MUST be ignored) are preserved
/// verbatim so the record round-trips unchanged.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsSerFmt {
    /// Raw flags word: `fSmoothedLine`, `f3DBubbles`, `fArShadow`, and the 13
    /// reserved bits.
    flags: u16,
}

impl XlsSerFmt {
    /// Parse a `SerFmt` record payload.
    pub fn parse(data: &[u8]) -> XlsResult<Self> {
        if data.len() != PAYLOAD_LEN {
            return Err(XlsError::InvalidLength {
                expected: PAYLOAD_LEN,
                found: data.len(),
            });
        }
        Ok(Self {
            flags: u16::from_le_bytes([data[0], data[1]]),
        })
    }

    /// Serialize back to a complete `SerFmt` record payload.
    pub fn to_payload(&self) -> Vec<u8> {
        self.flags.to_le_bytes().to_vec()
    }

    /// Whether the series lines use a smooth line effect (`fSmoothedLine`).
    pub fn smoothed_line(&self) -> bool {
        self.flags & FLAG_SMOOTHED_LINE != 0
    }

    /// Whether the data points use a 3-D effect (`f3DBubbles`).
    pub fn bubbles_3d(&self) -> bool {
        self.flags & FLAG_3D_BUBBLES != 0
    }

    /// Whether the data markers are displayed with a shadow (`fArShadow`).
    pub fn shadow(&self) -> bool {
        self.flags & FLAG_SHADOW != 0
    }

    /// Raw flags word, including the 13 reserved bits.
    pub fn flags(&self) -> u16 {
        self.flags
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn round_trip_all_flags() {
        for (flags, smoothed, bubbles, shadow) in [
            (0x0000u16, false, false, false),
            (0x0001, true, false, false),
            (0x0002, false, true, false),
            (0x0004, false, false, true),
            (0x0007, true, true, true),
        ] {
            let payload = flags.to_le_bytes();
            let parsed = XlsSerFmt::parse(&payload).unwrap();
            assert_eq!(parsed.smoothed_line(), smoothed);
            assert_eq!(parsed.bubbles_3d(), bubbles);
            assert_eq!(parsed.shadow(), shadow);
            assert_eq!(parsed.flags(), flags);
            assert_eq!(parsed.to_payload(), payload);
        }
    }

    #[test]
    fn preserves_reserved_bits() {
        // The 13 reserved bits MUST be zero and MUST be ignored; they
        // round-trip verbatim.
        let payload = 0xFFF8u16.to_le_bytes();
        let parsed = XlsSerFmt::parse(&payload).unwrap();
        assert!(!parsed.smoothed_line());
        assert_eq!(parsed.to_payload(), payload);
    }

    #[test]
    fn rejects_bad_length() {
        assert!(XlsSerFmt::parse(&[0x01]).is_err());
        assert!(XlsSerFmt::parse(&[0x01, 0x00, 0x00]).is_err());
    }
}
