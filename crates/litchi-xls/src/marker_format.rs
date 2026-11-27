//! BIFF8 `MarkerFormat` record (0x1009, MS-XLS 2.4.160) of the Chart Sheet
//! substream (MS-XLS 2.1): the color, size, and shape of data markers on
//! line, radar, and scatter chart groups.
//!
//! Everything in this module is INERT: values are stored verbatim and no
//! markers are rendered. The `rgbFore`/`rgbBack` values MUST match the
//! `icvFore`/`icvBack` chart colors (MS-XLS 2.4.160); that cross-field
//! constraint needs the workbook color table and is documented here, not
//! enforced by the record reader.
//!
//! # References
//!
//! - MS-XLS 2.4.160 (MarkerFormat), 2.5.162 (IcvChart), 2.5.177 (LongRGB)

use super::{Error, Result};

/// Record type of the `MarkerFormat` record (MS-XLS 2.4.160).
pub(crate) const MARKER_FORMAT_RECORD_TYPE: u16 = 0x1009;

/// Byte length of a `MarkerFormat` record payload.
const PAYLOAD_LEN: usize = 20;
/// Flags bit: `fAuto` (the data marker is automatically generated).
const FLAG_AUTO: u16 = 0x0001;
/// Flags bit: `fNotShowInt` (the data marker interior is not shown).
const FLAG_NOT_SHOW_INTERIOR: u16 = 0x0010;
/// Flags bit: `fNotShowBrd` (the data marker border is not shown).
const FLAG_NOT_SHOW_BORDER: u16 = 0x0020;
/// Minimum `miSize` value, in twips (MS-XLS 2.4.160).
const MIN_MARKER_SIZE: u32 = 40;
/// Maximum `miSize` value, in twips (MS-XLS 2.4.160).
const MAX_MARKER_SIZE: u32 = 1440;
/// Maximum `icv` value in the primary chart color range (MS-XLS 2.5.162).
const MAX_ICV_PRIMARY: u16 = 0x0041;
/// Minimum `icv` value in the extended chart color range (MS-XLS 2.5.162).
const MIN_ICV_EXTENDED: u16 = 0x004D;
/// Maximum `icv` value in the extended chart color range (MS-XLS 2.5.162).
const MAX_ICV_EXTENDED: u16 = 0x004F;

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidRecord {
        record_type: MARKER_FORMAT_RECORD_TYPE,
        message: message.into(),
    }
}

/// Validate an `IcvChart` chart color index (MS-XLS 2.5.162).
fn parse_icv(value: u16, field: &str) -> Result<u16> {
    if value <= MAX_ICV_PRIMARY || (MIN_ICV_EXTENDED..=MAX_ICV_EXTENDED).contains(&value) {
        Ok(value)
    } else {
        Err(invalid(format!(
            "MarkerFormat {field} {value:#06X} is not an IcvChart color"
        )))
    }
}

/// A `LongRGB` color (MS-XLS 2.5.177). The `reserved` byte (MUST be zero,
/// and MUST be ignored) is preserved verbatim so the record round-trips
/// unchanged.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ChartRgb {
    /// Relative intensity of red.
    pub red: u8,
    /// Relative intensity of green.
    pub green: u8,
    /// Relative intensity of blue.
    pub blue: u8,
    /// The `reserved` byte, preserved verbatim.
    pub reserved: u8,
}

impl ChartRgb {
    fn parse(data: &[u8]) -> Self {
        Self {
            red: data[0],
            green: data[1],
            blue: data[2],
            reserved: data[3],
        }
    }

    fn write_payload(&self, output: &mut Vec<u8>) {
        output.extend_from_slice(&[self.red, self.green, self.blue, self.reserved]);
    }
}

/// The `imk` data marker type (MS-XLS 2.4.160).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u16)]
pub enum DataMarkerKind {
    /// 0x0000: no marker.
    None = 0x0000,
    /// 0x0001: square markers.
    Square = 0x0001,
    /// 0x0002: diamond-shaped markers.
    Diamond = 0x0002,
    /// 0x0003: triangular markers.
    Triangle = 0x0003,
    /// 0x0004: square markers with an X.
    SquareX = 0x0004,
    /// 0x0005: square markers with an asterisk.
    SquareAsterisk = 0x0005,
    /// 0x0006: short bar markers.
    ShortBar = 0x0006,
    /// 0x0007: long bar markers.
    LongBar = 0x0007,
    /// 0x0008: circular markers.
    Circle = 0x0008,
    /// 0x0009: square markers with a plus sign.
    SquarePlus = 0x0009,
}

impl DataMarkerKind {
    fn parse(value: u16) -> Result<Self> {
        match value {
            0x0000 => Ok(Self::None),
            0x0001 => Ok(Self::Square),
            0x0002 => Ok(Self::Diamond),
            0x0003 => Ok(Self::Triangle),
            0x0004 => Ok(Self::SquareX),
            0x0005 => Ok(Self::SquareAsterisk),
            0x0006 => Ok(Self::ShortBar),
            0x0007 => Ok(Self::LongBar),
            0x0008 => Ok(Self::Circle),
            0x0009 => Ok(Self::SquarePlus),
            other => Err(invalid(format!(
                "MarkerFormat imk {other:#06X} is not a defined data marker type"
            ))),
        }
    }
}

/// Typed `MarkerFormat` record content (MS-XLS 2.4.160): the color, size,
/// and shape of the associated data markers.
///
/// The 13 reserved flags bits (`reserved1`/`reserved2`, MUST be ignored) are
/// preserved verbatim so the record round-trips unchanged.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct MarkerFormat {
    /// Border color of the data marker (`rgbFore`).
    foreground: ChartRgb,
    /// Interior color of the data marker (`rgbBack`).
    background: ChartRgb,
    /// The data marker type (`imk`).
    kind: DataMarkerKind,
    /// Raw flags word: `fAuto`, `fNotShowInt`, `fNotShowBrd`, and the 13
    /// reserved bits, preserved verbatim.
    flags: u16,
    /// Border chart color (`icvFore`).
    icv_foreground: u16,
    /// Interior chart color (`icvBack`).
    icv_background: u16,
    /// Marker size in twips (`miSize`), in 40..=1440.
    size_twips: u32,
}

impl MarkerFormat {
    /// Parse a `MarkerFormat` record payload.
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
        let size_twips = u32::from_le_bytes(data[16..20].try_into().expect("length checked"));
        if !(MIN_MARKER_SIZE..=MAX_MARKER_SIZE).contains(&size_twips) {
            return Err(invalid(format!(
                "MarkerFormat miSize {size_twips} is outside {MIN_MARKER_SIZE}..={MAX_MARKER_SIZE} twips"
            )));
        }
        Ok(Self {
            foreground: ChartRgb::parse(&data[0..4]),
            background: ChartRgb::parse(&data[4..8]),
            kind: DataMarkerKind::parse(u16::from_le_bytes([data[8], data[9]]))?,
            flags: u16::from_le_bytes([data[10], data[11]]),
            icv_foreground: parse_icv(u16::from_le_bytes([data[12], data[13]]), "icvFore")?,
            icv_background: parse_icv(u16::from_le_bytes([data[14], data[15]]), "icvBack")?,
            size_twips,
        })
    }

    /// Serialize back to a complete `MarkerFormat` record payload.
    #[must_use]
    pub fn to_payload(&self) -> Vec<u8> {
        let mut payload = Vec::with_capacity(PAYLOAD_LEN);
        self.foreground.write_payload(&mut payload);
        self.background.write_payload(&mut payload);
        payload.extend_from_slice(&(self.kind as u16).to_le_bytes());
        payload.extend_from_slice(&self.flags.to_le_bytes());
        payload.extend_from_slice(&self.icv_foreground.to_le_bytes());
        payload.extend_from_slice(&self.icv_background.to_le_bytes());
        payload.extend_from_slice(&self.size_twips.to_le_bytes());
        payload
    }

    /// Border color of the data marker (`rgbFore`).
    #[must_use]
    pub fn foreground(&self) -> ChartRgb {
        self.foreground
    }

    /// Interior color of the data marker (`rgbBack`).
    #[must_use]
    pub fn background(&self) -> ChartRgb {
        self.background
    }

    /// The data marker type (`imk`).
    #[must_use]
    pub fn kind(&self) -> DataMarkerKind {
        self.kind
    }

    /// Whether the data marker is automatically generated (`fAuto`).
    #[must_use]
    pub fn is_auto(&self) -> bool {
        self.flags & FLAG_AUTO != 0
    }

    /// Whether the data marker interior is not shown (`fNotShowInt`).
    #[must_use]
    pub fn hides_interior(&self) -> bool {
        self.flags & FLAG_NOT_SHOW_INTERIOR != 0
    }

    /// Whether the data marker border is not shown (`fNotShowBrd`).
    #[must_use]
    pub fn hides_border(&self) -> bool {
        self.flags & FLAG_NOT_SHOW_BORDER != 0
    }

    /// Raw flags word, including the 13 reserved bits.
    #[must_use]
    pub fn flags(&self) -> u16 {
        self.flags
    }

    /// Border chart color (`icvFore`).
    #[must_use]
    pub fn icv_foreground(&self) -> u16 {
        self.icv_foreground
    }

    /// Interior chart color (`icvBack`).
    #[must_use]
    pub fn icv_background(&self) -> u16 {
        self.icv_background
    }

    /// Marker size in twips (`miSize`).
    #[must_use]
    pub fn size_twips(&self) -> u32 {
        self.size_twips
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record(kind: u16, flags: u16, icv_fore: u16, icv_back: u16, size: u32) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&[0xFF, 0x00, 0x00, 0x00]);
        data.extend_from_slice(&[0xFF, 0xFF, 0xFF, 0x00]);
        data.extend_from_slice(&kind.to_le_bytes());
        data.extend_from_slice(&flags.to_le_bytes());
        data.extend_from_slice(&icv_fore.to_le_bytes());
        data.extend_from_slice(&icv_back.to_le_bytes());
        data.extend_from_slice(&size.to_le_bytes());
        data
    }

    #[test]
    fn round_trip_all_marker_kinds() {
        for (value, expected) in [
            (0x0000, DataMarkerKind::None),
            (0x0001, DataMarkerKind::Square),
            (0x0002, DataMarkerKind::Diamond),
            (0x0003, DataMarkerKind::Triangle),
            (0x0004, DataMarkerKind::SquareX),
            (0x0005, DataMarkerKind::SquareAsterisk),
            (0x0006, DataMarkerKind::ShortBar),
            (0x0007, DataMarkerKind::LongBar),
            (0x0008, DataMarkerKind::Circle),
            (0x0009, DataMarkerKind::SquarePlus),
        ] {
            let bytes = record(value, 0x0031, 0x004D, 0x0041, 100);
            let parsed = MarkerFormat::parse(&bytes).unwrap();
            assert_eq!(parsed.kind(), expected);
            assert!(parsed.is_auto());
            assert!(parsed.hides_interior());
            assert!(parsed.hides_border());
            assert_eq!(parsed.icv_foreground(), 0x004D);
            assert_eq!(parsed.icv_background(), 0x0041);
            assert_eq!(parsed.size_twips(), 100);
            assert_eq!(parsed.foreground().red, 0xFF);
            assert_eq!(parsed.background().green, 0xFF);
            assert_eq!(parsed.to_payload(), bytes);
        }
    }

    #[test]
    fn preserves_reserved_bits_and_rgb_reserved_bytes() {
        // The 13 reserved flags bits and the LongRGB reserved bytes MUST be
        // ignored but round-trip verbatim.
        let mut bytes = record(0x0008, 0xFFCE, 0x0000, 0x004F, 40);
        bytes[3] = 0xAA;
        bytes[7] = 0xBB;
        let parsed = MarkerFormat::parse(&bytes).unwrap();
        assert_eq!(parsed.flags(), 0xFFCE);
        assert!(!parsed.is_auto());
        assert_eq!(parsed.size_twips(), 40);
        assert_eq!(parsed.to_payload(), bytes);
    }

    #[test]
    fn rejects_malformed_records() {
        let bytes = record(0x0001, 0x0001, 0x0000, 0x0000, 100);
        // Truncated and overlong payloads.
        assert!(MarkerFormat::parse(&bytes[..19]).is_err());
        assert!(MarkerFormat::parse(&[bytes.as_slice(), &[0]].concat()).is_err());
        // Undefined imk value.
        assert!(MarkerFormat::parse(&record(0x000A, 0, 0, 0, 100)).is_err());
        // icv values outside the IcvChart ranges.
        assert!(MarkerFormat::parse(&record(0x0001, 0, 0x0042, 0, 100)).is_err());
        assert!(MarkerFormat::parse(&record(0x0001, 0, 0, 0x0050, 100)).is_err());
        // miSize outside 40..=1440 twips.
        assert!(MarkerFormat::parse(&record(0x0001, 0, 0, 0, 39)).is_err());
        assert!(MarkerFormat::parse(&record(0x0001, 0, 0, 0, 1441)).is_err());
        assert!(MarkerFormat::parse(&record(0x0001, 0, 0, 0, 1440)).is_ok());
    }
}
