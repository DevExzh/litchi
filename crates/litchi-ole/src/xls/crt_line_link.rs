//! BIFF8 chart `CrtLine` and `CrtLink` records of the Chart Sheet substream
//! (MS-XLS 2.1):
//!
//! - **CrtLine** (0x101C): the presence of drop, high-low, series, or leader
//!   lines on a chart group (MS-XLS 2.4.68).
//! - **CrtLink** (0x1022): written but unused (MS-XLS 2.4.69).
//!
//! Everything in this module is INERT: values are stored verbatim and no
//! chart lines are drawn. The uniqueness and ascending order of `CrtLine`
//! `id` values within a chart group are cross-record constraints documented
//! on the type; a single-record reader cannot enforce them.
//!
//! # References
//!
//! - MS-XLS 2.4.68 (CrtLine), 2.4.69 (CrtLink)

use super::{XlsError, XlsResult};

/// Record type of the `CrtLine` record (MS-XLS 2.4.68).
pub(crate) const CRT_LINE_RECORD_TYPE: u16 = 0x101C;

/// Byte length of a `CrtLine` record payload: the `id` field.
const CRT_LINE_LEN: usize = 2;
/// Byte length of a `CrtLink` record payload: 10 undefined bytes.
const CRT_LINK_LEN: usize = 10;

fn invalid(record_type: u16, message: impl Into<String>) -> XlsError {
    XlsError::InvalidRecord {
        record_type,
        message: message.into(),
    }
}

/// The `id` line type of a `CrtLine` record (MS-XLS 2.4.68). The value MUST
/// be unique and ascending among the `CrtLine` records of a chart group
/// (a cross-record constraint the caller validates).
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
#[repr(u16)]
pub enum XlsCrtLineKind {
    /// 0x0000: drop lines below the data points of line, area, and stock
    /// chart groups.
    DropLines = 0x0000,
    /// 0x0001: high-low lines around the data points of line and stock chart
    /// groups.
    HighLowLines = 0x0001,
    /// 0x0002: series lines of stacked column/bar and bar-of-pie/pie-of-pie
    /// chart groups.
    SeriesLines = 0x0002,
    /// 0x0003: leader lines connecting data labels to data points of pie and
    /// pie-of-pie chart groups.
    LeaderLines = 0x0003,
}

impl XlsCrtLineKind {
    fn parse(value: u16) -> XlsResult<Self> {
        match value {
            0x0000 => Ok(Self::DropLines),
            0x0001 => Ok(Self::HighLowLines),
            0x0002 => Ok(Self::SeriesLines),
            0x0003 => Ok(Self::LeaderLines),
            other => Err(invalid(
                CRT_LINE_RECORD_TYPE,
                format!("CrtLine id {other:#06X} is not a defined line type"),
            )),
        }
    }
}

/// Typed `CrtLine` record content (MS-XLS 2.4.68): the presence of a line
/// type on a chart group.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsCrtLine {
    /// The type of line present (`id`).
    kind: XlsCrtLineKind,
}

impl XlsCrtLine {
    /// Parse a `CrtLine` record payload.
    pub fn parse(data: &[u8]) -> XlsResult<Self> {
        if data.len() != CRT_LINE_LEN {
            return Err(XlsError::InvalidLength {
                expected: CRT_LINE_LEN,
                found: data.len(),
            });
        }
        Ok(Self {
            kind: XlsCrtLineKind::parse(u16::from_le_bytes([data[0], data[1]]))?,
        })
    }

    /// Serialize back to a complete `CrtLine` record payload.
    pub fn to_payload(&self) -> Vec<u8> {
        (self.kind as u16).to_le_bytes().to_vec()
    }

    /// The type of line present (`id`).
    pub fn kind(&self) -> XlsCrtLineKind {
        self.kind
    }
}

/// Typed `CrtLink` record content (MS-XLS 2.4.69): written but unused.
///
/// The 10 `unused` bytes are undefined and MUST be ignored; they are
/// preserved verbatim so the record round-trips unchanged.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsCrtLink {
    /// The undefined `unused` bytes, preserved verbatim.
    unused: [u8; CRT_LINK_LEN],
}

impl XlsCrtLink {
    /// Parse a `CrtLink` record payload.
    pub fn parse(data: &[u8]) -> XlsResult<Self> {
        if data.len() != CRT_LINK_LEN {
            return Err(XlsError::InvalidLength {
                expected: CRT_LINK_LEN,
                found: data.len(),
            });
        }
        Ok(Self {
            unused: data.try_into().expect("length checked"),
        })
    }

    /// Serialize back to a complete `CrtLink` record payload.
    pub fn to_payload(&self) -> Vec<u8> {
        self.unused.to_vec()
    }

    /// The preserved undefined `unused` bytes.
    pub fn unused(&self) -> [u8; CRT_LINK_LEN] {
        self.unused
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn crt_line_round_trip_all_kinds() {
        for (value, expected) in [
            (0x0000u16, XlsCrtLineKind::DropLines),
            (0x0001, XlsCrtLineKind::HighLowLines),
            (0x0002, XlsCrtLineKind::SeriesLines),
            (0x0003, XlsCrtLineKind::LeaderLines),
        ] {
            let payload = value.to_le_bytes();
            let parsed = XlsCrtLine::parse(&payload).unwrap();
            assert_eq!(parsed.kind(), expected);
            assert_eq!(parsed.to_payload(), payload);
        }
    }

    #[test]
    fn crt_line_rejects_malformed_records() {
        assert!(XlsCrtLine::parse(&[0x00]).is_err());
        assert!(XlsCrtLine::parse(&[0x00, 0x00, 0x00]).is_err());
        assert!(XlsCrtLine::parse(&0x0004u16.to_le_bytes()).is_err());
        assert!(XlsCrtLine::parse(&0xFFFFu16.to_le_bytes()).is_err());
    }

    #[test]
    fn crt_link_round_trip_and_rejects() {
        let payload = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10];
        let parsed = XlsCrtLink::parse(&payload).unwrap();
        assert_eq!(parsed.unused(), payload);
        assert_eq!(parsed.to_payload(), payload);
        // The undefined bytes MUST be ignored, including zero payloads.
        let zero = [0; CRT_LINK_LEN];
        assert_eq!(XlsCrtLink::parse(&zero).unwrap().to_payload(), zero);

        assert!(XlsCrtLink::parse(&payload[..9]).is_err());
        assert!(XlsCrtLink::parse(&[payload.as_slice(), &[0]].concat()).is_err());
    }
}
