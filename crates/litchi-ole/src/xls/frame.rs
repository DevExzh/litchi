//! BIFF8 `Frame` record (0x1032, MS-XLS 2.4.128) of the Chart Sheet
//! substream (MS-XLS 2.1): the type, size, and position of the frame around
//! a chart element.
//!
//! Everything in this module is INERT: values are stored verbatim and no
//! frame is drawn.
//!
//! # References
//!
//! - MS-XLS 2.4.128 (Frame)

use super::{XlsError, XlsResult};

/// Record type of the `Frame` record (MS-XLS 2.4.128).
pub(crate) const FRAME_RECORD_TYPE: u16 = 0x1032;

/// Byte length of a `Frame` record payload: `frt` + flags.
const PAYLOAD_LEN: usize = 4;
/// Flags bit: `fAutoSize` (the frame size is automatically calculated).
const FLAG_AUTO_SIZE: u16 = 0x0001;
/// Flags bit: `fAutoPosition` (the frame position is automatically
/// calculated).
const FLAG_AUTO_POSITION: u16 = 0x0002;

fn invalid(message: impl Into<String>) -> XlsError {
    XlsError::InvalidRecord {
        record_type: FRAME_RECORD_TYPE,
        message: message.into(),
    }
}

/// The `frt` frame type (MS-XLS 2.4.128).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u16)]
pub enum XlsFrameType {
    /// 0x0000: a frame surrounding the chart element.
    Surrounding = 0x0000,
    /// 0x0004: a frame with a shadow surrounding the chart element.
    Shadowed = 0x0004,
}

impl XlsFrameType {
    fn parse(value: u16) -> XlsResult<Self> {
        match value {
            0x0000 => Ok(Self::Surrounding),
            0x0004 => Ok(Self::Shadowed),
            other => Err(invalid(format!(
                "Frame frt {other:#06X} is not a defined frame type"
            ))),
        }
    }
}

/// Typed `Frame` record content (MS-XLS 2.4.128): the type, size, and
/// position of the frame around a chart element.
///
/// The 14 `reserved` bits (MUST be zero, and MUST be ignored) are preserved
/// verbatim so the record round-trips unchanged.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsFrame {
    /// The frame type (`frt`).
    frame_type: XlsFrameType,
    /// Raw flags word: `fAutoSize`, `fAutoPosition`, and the 14 reserved bits.
    flags: u16,
}

impl XlsFrame {
    /// Parse a `Frame` record payload.
    pub fn parse(data: &[u8]) -> XlsResult<Self> {
        if data.len() != PAYLOAD_LEN {
            return Err(XlsError::InvalidLength {
                expected: PAYLOAD_LEN,
                found: data.len(),
            });
        }
        Ok(Self {
            frame_type: XlsFrameType::parse(u16::from_le_bytes([data[0], data[1]]))?,
            flags: u16::from_le_bytes([data[2], data[3]]),
        })
    }

    /// Serialize back to a complete `Frame` record payload.
    pub fn to_payload(&self) -> Vec<u8> {
        let mut payload = Vec::with_capacity(PAYLOAD_LEN);
        payload.extend_from_slice(&(self.frame_type as u16).to_le_bytes());
        payload.extend_from_slice(&self.flags.to_le_bytes());
        payload
    }

    /// The frame type (`frt`).
    pub fn frame_type(&self) -> XlsFrameType {
        self.frame_type
    }

    /// Whether the frame size is automatically calculated (`fAutoSize`).
    pub fn auto_size(&self) -> bool {
        self.flags & FLAG_AUTO_SIZE != 0
    }

    /// Whether the frame position is automatically calculated
    /// (`fAutoPosition`).
    pub fn auto_position(&self) -> bool {
        self.flags & FLAG_AUTO_POSITION != 0
    }

    /// Raw flags word, including the 14 reserved bits.
    pub fn flags(&self) -> u16 {
        self.flags
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn round_trip_all_frame_types() {
        for (frt, expected) in [
            (0x0000u16, XlsFrameType::Surrounding),
            (0x0004, XlsFrameType::Shadowed),
        ] {
            let mut payload = frt.to_le_bytes().to_vec();
            payload.extend_from_slice(&0x0003u16.to_le_bytes());
            let parsed = XlsFrame::parse(&payload).unwrap();
            assert_eq!(parsed.frame_type(), expected);
            assert!(parsed.auto_size());
            assert!(parsed.auto_position());
            assert_eq!(parsed.to_payload(), payload);
        }
    }

    #[test]
    fn preserves_reserved_bits() {
        // The 14 reserved bits MUST be ignored but round-trip verbatim.
        let mut payload = 0x0000u16.to_le_bytes().to_vec();
        payload.extend_from_slice(&0xFFFCu16.to_le_bytes());
        let parsed = XlsFrame::parse(&payload).unwrap();
        assert_eq!(parsed.flags(), 0xFFFC);
        assert!(!parsed.auto_size());
        assert_eq!(parsed.to_payload(), payload);
    }

    #[test]
    fn rejects_malformed_records() {
        let mut payload = 0x0000u16.to_le_bytes().to_vec();
        payload.extend_from_slice(&[0; 2]);
        // Bad length.
        assert!(XlsFrame::parse(&payload[..3]).is_err());
        assert!(XlsFrame::parse(&[payload.as_slice(), &[0]].concat()).is_err());
        // Undefined frt values.
        let mut bad = payload.clone();
        bad[0..2].copy_from_slice(&0x0001u16.to_le_bytes());
        assert!(XlsFrame::parse(&bad).is_err());
        bad[0..2].copy_from_slice(&0x0003u16.to_le_bytes());
        assert!(XlsFrame::parse(&bad).is_err());
    }
}
