//! Strict support for MS-PPT OfficeArtClientAnchor records.

use std::io::{self, Write};

use super::package::{PptError, Result};

/// OfficeArt record type for an OfficeArtClientAnchor record.
pub const OFFICE_ART_CLIENT_ANCHOR_RECORD_TYPE: u16 = 0xF010;

const OFFICE_ART_HEADER_LEN: usize = 8;
const SMALL_RECT_LEN: usize = 8;
const RECT_LEN: usize = 16;

/// Resource limits for parsing an OfficeArtClientAnchor record.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PowerPointClientAnchorLimits {
    /// Maximum accepted payload size, excluding the eight-byte OfficeArt header.
    pub max_payload_bytes: usize,
}

impl Default for PowerPointClientAnchorLimits {
    fn default() -> Self {
        Self {
            max_payload_bytes: RECT_LEN,
        }
    }
}

/// The exact rectangle representation carried by an anchor.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PowerPointClientAnchorEncoding {
    /// An eight-byte SmallRectStruct containing four signed 16-bit coordinates.
    SmallRect,
    /// A sixteen-byte RectStruct containing four signed 32-bit coordinates.
    Rect,
}

/// The MS-PPT SmallRectStruct payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PowerPointSmallRect {
    top: i16,
    left: i16,
    right: i16,
    bottom: i16,
}

impl PowerPointSmallRect {
    /// Construct a compact rectangle. Bounds are supplied in geometric order.
    pub fn new(left: i16, top: i16, right: i16, bottom: i16) -> Result<Self> {
        validate_bounds(left as i32, top as i32, right as i32, bottom as i32)?;
        Ok(Self {
            top,
            left,
            right,
            bottom,
        })
    }

    /// Minimum x-coordinate.
    pub fn left(self) -> i16 {
        self.left
    }

    /// Minimum y-coordinate.
    pub fn top(self) -> i16 {
        self.top
    }

    /// Maximum x-coordinate.
    pub fn right(self) -> i16 {
        self.right
    }

    /// Maximum y-coordinate.
    pub fn bottom(self) -> i16 {
        self.bottom
    }
}

/// The MS-PPT RectStruct payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PowerPointRect {
    top: i32,
    left: i32,
    right: i32,
    bottom: i32,
}

impl PowerPointRect {
    /// Construct a full-width rectangle. Bounds are supplied in geometric order.
    pub fn new(left: i32, top: i32, right: i32, bottom: i32) -> Result<Self> {
        validate_bounds(left, top, right, bottom)?;
        Ok(Self {
            top,
            left,
            right,
            bottom,
        })
    }

    /// Minimum x-coordinate.
    pub fn left(self) -> i32 {
        self.left
    }

    /// Minimum y-coordinate.
    pub fn top(self) -> i32 {
        self.top
    }

    /// Maximum x-coordinate.
    pub fn right(self) -> i32 {
        self.right
    }

    /// Maximum y-coordinate.
    pub fn bottom(self) -> i32 {
        self.bottom
    }
}

/// The variable OfficeArtClientAnchorData payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PowerPointClientAnchorData {
    /// The original record used the compact eight-byte form.
    SmallRect(PowerPointSmallRect),
    /// The original record used the full sixteen-byte form.
    Rect(PowerPointRect),
}

impl PowerPointClientAnchorData {
    /// Parse an exact payload, selecting its representation from its length.
    pub fn parse(bytes: &[u8]) -> Result<Self> {
        Self::parse_with_limits(bytes, PowerPointClientAnchorLimits::default())
    }

    /// Parse an exact payload with a caller-supplied allocation/input bound.
    pub fn parse_with_limits(bytes: &[u8], limits: PowerPointClientAnchorLimits) -> Result<Self> {
        if bytes.len() > limits.max_payload_bytes {
            return Err(corrupted(
                "OfficeArtClientAnchorData exceeds the configured limit",
            ));
        }

        match bytes.len() {
            SMALL_RECT_LEN => Ok(Self::SmallRect(PowerPointSmallRect::new(
                i16_at(bytes, 2),
                i16_at(bytes, 0),
                i16_at(bytes, 4),
                i16_at(bytes, 6),
            )?)),
            RECT_LEN => Ok(Self::Rect(PowerPointRect::new(
                i32_at(bytes, 4),
                i32_at(bytes, 0),
                i32_at(bytes, 8),
                i32_at(bytes, 12),
            )?)),
            length => Err(PptError::Corrupted(format!(
                "OfficeArtClientAnchorData length must be 8 or 16 bytes, got {length}"
            ))),
        }
    }

    /// Return the representation used on the wire.
    pub fn encoding(self) -> PowerPointClientAnchorEncoding {
        match self {
            Self::SmallRect(_) => PowerPointClientAnchorEncoding::SmallRect,
            Self::Rect(_) => PowerPointClientAnchorEncoding::Rect,
        }
    }

    /// Exact encoded payload length.
    pub fn encoded_len(self) -> usize {
        match self {
            Self::SmallRect(_) => SMALL_RECT_LEN,
            Self::Rect(_) => RECT_LEN,
        }
    }

    /// Minimum x-coordinate, normalized to 32 bits without changing the encoding.
    pub fn left(self) -> i32 {
        match self {
            Self::SmallRect(rect) => rect.left() as i32,
            Self::Rect(rect) => rect.left(),
        }
    }

    /// Minimum y-coordinate, normalized to 32 bits without changing the encoding.
    pub fn top(self) -> i32 {
        match self {
            Self::SmallRect(rect) => rect.top() as i32,
            Self::Rect(rect) => rect.top(),
        }
    }

    /// Maximum x-coordinate, normalized to 32 bits without changing the encoding.
    pub fn right(self) -> i32 {
        match self {
            Self::SmallRect(rect) => rect.right() as i32,
            Self::Rect(rect) => rect.right(),
        }
    }

    /// Maximum y-coordinate, normalized to 32 bits without changing the encoding.
    pub fn bottom(self) -> i32 {
        match self {
            Self::SmallRect(rect) => rect.bottom() as i32,
            Self::Rect(rect) => rect.bottom(),
        }
    }

    /// Rectangle width, widened so every valid i32 rectangle is representable.
    pub fn width(self) -> i64 {
        i64::from(self.right()) - i64::from(self.left())
    }

    /// Rectangle height, widened so every valid i32 rectangle is representable.
    pub fn height(self) -> i64 {
        i64::from(self.bottom()) - i64::from(self.top())
    }

    /// Serialize only the OfficeArtClientAnchorData payload.
    pub fn to_bytes(self) -> Vec<u8> {
        let mut bytes = Vec::with_capacity(self.encoded_len());
        match self {
            Self::SmallRect(rect) => {
                bytes.extend_from_slice(&rect.top.to_le_bytes());
                bytes.extend_from_slice(&rect.left.to_le_bytes());
                bytes.extend_from_slice(&rect.right.to_le_bytes());
                bytes.extend_from_slice(&rect.bottom.to_le_bytes());
            },
            Self::Rect(rect) => {
                bytes.extend_from_slice(&rect.top.to_le_bytes());
                bytes.extend_from_slice(&rect.left.to_le_bytes());
                bytes.extend_from_slice(&rect.right.to_le_bytes());
                bytes.extend_from_slice(&rect.bottom.to_le_bytes());
            },
        }
        bytes
    }
}

/// A complete OfficeArtClientAnchor record, including its strict header.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PowerPointClientAnchor {
    data: PowerPointClientAnchorData,
}

impl PowerPointClientAnchor {
    /// Construct an anchor from an already validated payload.
    pub fn new(data: PowerPointClientAnchorData) -> Self {
        Self { data }
    }

    /// Construct an eight-byte SmallRectStruct anchor.
    pub fn small(left: i16, top: i16, right: i16, bottom: i16) -> Result<Self> {
        Ok(Self::new(PowerPointClientAnchorData::SmallRect(
            PowerPointSmallRect::new(left, top, right, bottom)?,
        )))
    }

    /// Construct a sixteen-byte RectStruct anchor.
    pub fn rect(left: i32, top: i32, right: i32, bottom: i32) -> Result<Self> {
        Ok(Self::new(PowerPointClientAnchorData::Rect(
            PowerPointRect::new(left, top, right, bottom)?,
        )))
    }

    /// Parse one complete record and reject truncation or trailing bytes.
    pub fn parse(bytes: &[u8]) -> Result<Self> {
        Self::parse_with_limits(bytes, PowerPointClientAnchorLimits::default())
    }

    /// Parse one complete record with a caller-supplied input bound.
    pub fn parse_with_limits(bytes: &[u8], limits: PowerPointClientAnchorLimits) -> Result<Self> {
        if bytes.len() < OFFICE_ART_HEADER_LEN {
            return Err(corrupted("OfficeArtClientAnchor header is truncated"));
        }

        let version_instance = u16::from_le_bytes([bytes[0], bytes[1]]);
        let version = version_instance & 0x000F;
        let instance = version_instance >> 4;
        if version != 0 {
            return Err(PptError::Corrupted(format!(
                "OfficeArtClientAnchor recVer must be 0, got {version}"
            )));
        }
        if instance != 0 {
            return Err(PptError::Corrupted(format!(
                "OfficeArtClientAnchor recInstance must be 0, got {instance}"
            )));
        }

        let record_type = u16::from_le_bytes([bytes[2], bytes[3]]);
        if record_type != OFFICE_ART_CLIENT_ANCHOR_RECORD_TYPE {
            return Err(PptError::Corrupted(format!(
                "OfficeArtClientAnchor recType must be 0xF010, got 0x{record_type:04X}"
            )));
        }

        let payload_len = u32::from_le_bytes([bytes[4], bytes[5], bytes[6], bytes[7]]) as usize;
        if payload_len != SMALL_RECT_LEN && payload_len != RECT_LEN {
            return Err(PptError::Corrupted(format!(
                "OfficeArtClientAnchor recLen must be 8 or 16, got {payload_len}"
            )));
        }
        if payload_len > limits.max_payload_bytes {
            return Err(corrupted(
                "OfficeArtClientAnchor payload exceeds the configured limit",
            ));
        }

        let expected_len = OFFICE_ART_HEADER_LEN + payload_len;
        if bytes.len() != expected_len {
            return Err(PptError::Corrupted(format!(
                "OfficeArtClientAnchor record length is {0}, expected {expected_len}",
                bytes.len()
            )));
        }

        Ok(Self::new(PowerPointClientAnchorData::parse_with_limits(
            &bytes[OFFICE_ART_HEADER_LEN..],
            limits,
        )?))
    }

    /// The exact variable payload.
    pub fn data(self) -> PowerPointClientAnchorData {
        self.data
    }

    /// Return the original compact or full-width encoding.
    pub fn encoding(self) -> PowerPointClientAnchorEncoding {
        self.data.encoding()
    }

    /// Minimum x-coordinate.
    pub fn left(self) -> i32 {
        self.data.left()
    }

    /// Minimum y-coordinate.
    pub fn top(self) -> i32 {
        self.data.top()
    }

    /// Maximum x-coordinate.
    pub fn right(self) -> i32 {
        self.data.right()
    }

    /// Maximum y-coordinate.
    pub fn bottom(self) -> i32 {
        self.data.bottom()
    }

    /// Rectangle width as a non-overflowing signed value.
    pub fn width(self) -> i64 {
        self.data.width()
    }

    /// Rectangle height as a non-overflowing signed value.
    pub fn height(self) -> i64 {
        self.data.height()
    }

    /// Exact complete record length, including the OfficeArt header.
    pub fn encoded_len(self) -> usize {
        OFFICE_ART_HEADER_LEN + self.data.encoded_len()
    }

    /// Serialize the complete record to a newly allocated byte vector.
    pub fn to_bytes(self) -> Vec<u8> {
        let mut bytes = Vec::with_capacity(self.encoded_len());
        bytes.extend_from_slice(&0u16.to_le_bytes());
        bytes.extend_from_slice(&OFFICE_ART_CLIENT_ANCHOR_RECORD_TYPE.to_le_bytes());
        bytes.extend_from_slice(&(self.data.encoded_len() as u32).to_le_bytes());
        bytes.extend_from_slice(&self.data.to_bytes());
        bytes
    }

    /// Write the complete record without changing its compact/full representation.
    pub fn write_to<W: Write>(self, writer: &mut W) -> io::Result<()> {
        writer.write_all(&0u16.to_le_bytes())?;
        writer.write_all(&OFFICE_ART_CLIENT_ANCHOR_RECORD_TYPE.to_le_bytes())?;
        writer.write_all(&(self.data.encoded_len() as u32).to_le_bytes())?;
        match self.data {
            PowerPointClientAnchorData::SmallRect(rect) => {
                writer.write_all(&rect.top.to_le_bytes())?;
                writer.write_all(&rect.left.to_le_bytes())?;
                writer.write_all(&rect.right.to_le_bytes())?;
                writer.write_all(&rect.bottom.to_le_bytes())?;
            },
            PowerPointClientAnchorData::Rect(rect) => {
                writer.write_all(&rect.top.to_le_bytes())?;
                writer.write_all(&rect.left.to_le_bytes())?;
                writer.write_all(&rect.right.to_le_bytes())?;
                writer.write_all(&rect.bottom.to_le_bytes())?;
            },
        }
        Ok(())
    }
}

fn validate_bounds(left: i32, top: i32, right: i32, bottom: i32) -> Result<()> {
    if left > right {
        return Err(corrupted("OfficeArtClientAnchor left exceeds right"));
    }
    if top > bottom {
        return Err(corrupted("OfficeArtClientAnchor top exceeds bottom"));
    }
    Ok(())
}

fn i16_at(bytes: &[u8], offset: usize) -> i16 {
    i16::from_le_bytes([bytes[offset], bytes[offset + 1]])
}

fn i32_at(bytes: &[u8], offset: usize) -> i32 {
    i32::from_le_bytes([
        bytes[offset],
        bytes[offset + 1],
        bytes[offset + 2],
        bytes[offset + 3],
    ])
}

fn corrupted(message: &str) -> PptError {
    PptError::Corrupted(message.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record(payload: &[u8]) -> Vec<u8> {
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&0u16.to_le_bytes());
        bytes.extend_from_slice(&OFFICE_ART_CLIENT_ANCHOR_RECORD_TYPE.to_le_bytes());
        bytes.extend_from_slice(&(payload.len() as u32).to_le_bytes());
        bytes.extend_from_slice(payload);
        bytes
    }

    #[test]
    fn small_rect_round_trips_with_normative_field_order() {
        let payload = [
            0x9C, 0xFF, // top = -100
            0x38, 0xFF, // left = -200
            0x2C, 0x01, // right = 300
            0x90, 0x01, // bottom = 400
        ];
        let bytes = record(&payload);
        let anchor = PowerPointClientAnchor::parse(&bytes).unwrap();

        assert_eq!(anchor.encoding(), PowerPointClientAnchorEncoding::SmallRect);
        assert_eq!((anchor.left(), anchor.top()), (-200, -100));
        assert_eq!((anchor.right(), anchor.bottom()), (300, 400));
        assert_eq!((anchor.width(), anchor.height()), (500, 500));
        assert_eq!(anchor.to_bytes(), bytes);
    }

    #[test]
    fn full_rect_round_trips_extreme_coordinates_without_geometry_overflow() {
        let anchor = PowerPointClientAnchor::rect(i32::MIN, -7, i32::MAX, 9).unwrap();
        let bytes = anchor.to_bytes();
        let parsed = PowerPointClientAnchor::parse(&bytes).unwrap();

        assert_eq!(parsed.encoding(), PowerPointClientAnchorEncoding::Rect);
        assert_eq!(parsed.width(), u32::MAX as i64);
        assert_eq!(parsed.height(), 16);
        assert_eq!(parsed.to_bytes(), bytes);
    }

    #[test]
    fn write_to_is_byte_exact_for_both_encodings() {
        for anchor in [
            PowerPointClientAnchor::small(-2, -1, 4, 8).unwrap(),
            PowerPointClientAnchor::rect(-200_000, -100_000, 400_000, 800_000).unwrap(),
        ] {
            let expected = anchor.to_bytes();
            let mut actual = Vec::new();
            anchor.write_to(&mut actual).unwrap();
            assert_eq!(actual, expected);
        }
    }

    #[test]
    fn rejects_invalid_header_fields_lengths_and_bounds() {
        let valid = PowerPointClientAnchor::small(1, 2, 3, 4)
            .unwrap()
            .to_bytes();

        let mut bad_version = valid.clone();
        bad_version[0] = 1;
        assert!(PowerPointClientAnchor::parse(&bad_version).is_err());

        let mut bad_instance = valid.clone();
        bad_instance[1] = 1;
        assert!(PowerPointClientAnchor::parse(&bad_instance).is_err());

        let mut bad_type = valid.clone();
        bad_type[2] ^= 1;
        assert!(PowerPointClientAnchor::parse(&bad_type).is_err());

        let mut bad_length = valid.clone();
        bad_length[4..8].copy_from_slice(&12u32.to_le_bytes());
        assert!(PowerPointClientAnchor::parse(&bad_length).is_err());

        let mut trailing = valid.clone();
        trailing.push(0);
        assert!(PowerPointClientAnchor::parse(&trailing).is_err());
        assert!(PowerPointClientAnchor::parse(&valid[..valid.len() - 1]).is_err());
        assert!(PowerPointClientAnchor::small(4, 2, 3, 5).is_err());
        assert!(PowerPointClientAnchor::rect(1, 8, 3, 4).is_err());
    }

    #[test]
    fn enforces_payload_limit_before_parsing() {
        let bytes = PowerPointClientAnchor::rect(1, 2, 3, 4).unwrap().to_bytes();
        let limits = PowerPointClientAnchorLimits {
            max_payload_bytes: 8,
        };
        assert!(PowerPointClientAnchor::parse_with_limits(&bytes, limits).is_err());
    }
}
