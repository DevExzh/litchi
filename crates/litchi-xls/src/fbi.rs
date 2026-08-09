//! BIFF8 `Fbi` record (0x1060, MS-XLS 2.4.109) of the Chart Sheet substream
//! (MS-XLS 2.1): font information at the time a scalable font is added to
//! the chart.
//!
//! Everything in this module is INERT: the font scaling inputs are stored
//! verbatim and the font scaling algorithm of MS-XLS 2.4.109 is not applied.
//!
//! # References
//!
//! - MS-XLS 2.4.109 (Fbi), 2.5.14 (Boolean)

use super::{Error, Result};

/// Record type of the `Fbi` record (MS-XLS 2.4.109).
pub(crate) const FBI_RECORD_TYPE: u16 = 0x1060;

/// Byte length of an `Fbi` record payload.
const PAYLOAD_LEN: usize = 10;
/// Maximum `dmixBasis`/`dmiyBasis` value (MS-XLS 2.4.109).
const MAX_BASIS: u16 = 0x7FFF;
/// Minimum `twpHeightBasis` value, in twips (MS-XLS 2.4.109).
const MIN_HEIGHT_BASIS: u16 = 20;
/// Maximum `twpHeightBasis` value, in twips (MS-XLS 2.4.109).
const MAX_HEIGHT_BASIS: u16 = 8180;

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidRecord {
        record_type: FBI_RECORD_TYPE,
        message: message.into(),
    }
}

/// The `scab` scale basis of an `Fbi` record (MS-XLS 2.4.109).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u16)]
pub enum FontScaleBasis {
    /// 0x0000: scale by chart area.
    ChartArea = 0x0000,
    /// 0x0001: scale by plot area.
    PlotArea = 0x0001,
}

impl FontScaleBasis {
    fn parse(value: u16) -> Result<Self> {
        // Boolean (MS-XLS 2.5.14): only 0x0000 and 0x0001 are legal.
        match value {
            0x0000 => Ok(Self::ChartArea),
            0x0001 => Ok(Self::PlotArea),
            other => Err(invalid(format!("Fbi scab {other:#06X} is not a Boolean"))),
        }
    }
}

/// Typed `Fbi` record content (MS-XLS 2.4.109): font information for a
/// scalable chart font.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Fbi {
    /// Font width in twips when the font was first applied (`dmixBasis`).
    width_basis: u16,
    /// Font height in twips when the font was first applied (`dmiyBasis`).
    height_basis: u16,
    /// Default font height in twips (`twpHeightBasis`), in 20..=8180.
    font_height_basis: u16,
    /// The scale basis (`scab`).
    scale: FontScaleBasis,
    /// The font index (`ifnt`), a `FontIndex` structure.
    font_index: u16,
}

impl Fbi {
    /// Parse an `Fbi` record payload.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn parse(data: &[u8]) -> Result<Self> {
        if data.len() != PAYLOAD_LEN {
            return Err(Error::InvalidLength {
                expected: PAYLOAD_LEN,
                found: data.len(),
            });
        }
        let width_basis = u16::from_le_bytes([data[0], data[1]]);
        let height_basis = u16::from_le_bytes([data[2], data[3]]);
        for (field, value) in [("dmixBasis", width_basis), ("dmiyBasis", height_basis)] {
            if value > MAX_BASIS {
                return Err(invalid(format!(
                    "Fbi {field} {value:#06X} exceeds {MAX_BASIS:#06X}"
                )));
            }
        }
        let font_height_basis = u16::from_le_bytes([data[4], data[5]]);
        if !(MIN_HEIGHT_BASIS..=MAX_HEIGHT_BASIS).contains(&font_height_basis) {
            return Err(invalid(format!(
                "Fbi twpHeightBasis {font_height_basis} is outside {MIN_HEIGHT_BASIS}..={MAX_HEIGHT_BASIS}"
            )));
        }
        Ok(Self {
            width_basis,
            height_basis,
            font_height_basis,
            scale: FontScaleBasis::parse(u16::from_le_bytes([data[6], data[7]]))?,
            font_index: u16::from_le_bytes([data[8], data[9]]),
        })
    }

    /// Serialize back to a complete `Fbi` record payload.
    #[must_use]
    pub fn to_payload(&self) -> Vec<u8> {
        let mut payload = Vec::with_capacity(PAYLOAD_LEN);
        payload.extend_from_slice(&self.width_basis.to_le_bytes());
        payload.extend_from_slice(&self.height_basis.to_le_bytes());
        payload.extend_from_slice(&self.font_height_basis.to_le_bytes());
        payload.extend_from_slice(&(self.scale as u16).to_le_bytes());
        payload.extend_from_slice(&self.font_index.to_le_bytes());
        payload
    }

    /// Font width in twips when the font was first applied (`dmixBasis`).
    #[must_use]
    pub fn width_basis(&self) -> u16 {
        self.width_basis
    }

    /// Font height in twips when the font was first applied (`dmiyBasis`).
    #[must_use]
    pub fn height_basis(&self) -> u16 {
        self.height_basis
    }

    /// Default font height in twips (`twpHeightBasis`).
    #[must_use]
    pub fn font_height_basis(&self) -> u16 {
        self.font_height_basis
    }

    /// The scale basis (`scab`).
    #[must_use]
    pub fn scale(&self) -> FontScaleBasis {
        self.scale
    }

    /// The font index (`ifnt`).
    #[must_use]
    pub fn font_index(&self) -> u16 {
        self.font_index
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record(width: u16, height: u16, font_height: u16, scale: u16, font: u16) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&width.to_le_bytes());
        data.extend_from_slice(&height.to_le_bytes());
        data.extend_from_slice(&font_height.to_le_bytes());
        data.extend_from_slice(&scale.to_le_bytes());
        data.extend_from_slice(&font.to_le_bytes());
        data
    }

    #[test]
    fn round_trip_both_scales() {
        for (scale, expected) in [
            (0x0000, FontScaleBasis::ChartArea),
            (0x0001, FontScaleBasis::PlotArea),
        ] {
            let bytes = record(96, 1440, 240, scale, 5);
            let parsed = Fbi::parse(&bytes).unwrap();
            assert_eq!(parsed.width_basis(), 96);
            assert_eq!(parsed.height_basis(), 1440);
            assert_eq!(parsed.font_height_basis(), 240);
            assert_eq!(parsed.scale(), expected);
            assert_eq!(parsed.font_index(), 5);
            assert_eq!(parsed.to_payload(), bytes);
        }
    }

    #[test]
    fn accepts_basis_bounds() {
        assert!(Fbi::parse(&record(0x7FFF, 0x7FFF, 20, 0, 0)).is_ok());
        assert!(Fbi::parse(&record(0, 0, 8180, 0, 0)).is_ok());
    }

    #[test]
    fn rejects_malformed_records() {
        let bytes = record(96, 1440, 240, 0, 5);
        // Truncated and overlong payloads.
        assert!(Fbi::parse(&bytes[..9]).is_err());
        assert!(Fbi::parse(&[bytes.as_slice(), &[0]].concat()).is_err());
        // dmixBasis / dmiyBasis above 0x7FFF.
        assert!(Fbi::parse(&record(0x8000, 0, 240, 0, 0)).is_err());
        assert!(Fbi::parse(&record(0, 0x8000, 240, 0, 0)).is_err());
        // twpHeightBasis outside 20..=8180.
        assert!(Fbi::parse(&record(0, 0, 19, 0, 0)).is_err());
        assert!(Fbi::parse(&record(0, 0, 8181, 0, 0)).is_err());
        // scab is a Boolean.
        assert!(Fbi::parse(&record(0, 0, 240, 0x0002, 0)).is_err());
    }
}
