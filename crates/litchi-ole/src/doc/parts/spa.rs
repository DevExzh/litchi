//! Shape Positioning Attributes (SPA) for floating shapes.
//!
//! A **PlcfSpa** ([MS-DOC] 2.8.27) maps floating-shape anchor character
//! positions to **Spa** records ([MS-DOC] 2.9.253) that carry the shape's
//! position rectangle, position origin, and text-wrapping style. The Main
//! Document table is referenced by `fcPlcSpaMom` and the Header Document
//! table by `fcPlcSpaHdr` in the FIB.

use super::super::package::{DocError, Result};

/// Size of one Spa structure in bytes ([MS-DOC] 2.9.253).
pub const SPA_LEN: usize = 26;

/// FIB index (into `FileInformationBlock::get_table_pointer`) of the Main
/// Document `fcPlcSpaMom`/`lcbPlcSpaMom` pair.
pub(crate) const FIB_INDEX_PLC_SPA_MOM: usize = 40;
/// FIB index of the Header Document `fcPlcSpaHdr`/`lcbPlcSpaHdr` pair.
pub(crate) const FIB_INDEX_PLC_SPA_HDR: usize = 41;

/// Horizontal position origin of a floating shape (Spa `bx`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ShapeHorizontalOrigin {
    /// Anchored at the leading margin of the page.
    Margin = 0,
    /// Anchored at the leading edge of the page.
    Page = 1,
    /// Anchored at the leading edge of the column.
    Column = 2,
}

impl ShapeHorizontalOrigin {
    fn from_bits(value: u8) -> Result<Self> {
        Ok(match value {
            0 => Self::Margin,
            1 => Self::Page,
            2 => Self::Column,
            other => {
                return Err(DocError::InvalidFormat(format!(
                    "Invalid SPA horizontal origin: {other}"
                )))
            },
        })
    }
}

/// Vertical position origin of a floating shape (Spa `by`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ShapeVerticalOrigin {
    /// Anchored at the top margin of the page.
    Margin = 0,
    /// Anchored at the top edge of the page.
    Page = 1,
    /// Anchored at the top edge of the paragraph.
    Paragraph = 2,
}

impl ShapeVerticalOrigin {
    fn from_bits(value: u8) -> Result<Self> {
        Ok(match value {
            0 => Self::Margin,
            1 => Self::Page,
            2 => Self::Paragraph,
            other => {
                return Err(DocError::InvalidFormat(format!(
                    "Invalid SPA vertical origin: {other}"
                )))
            },
        })
    }
}

/// Text-wrapping style around a floating shape (Spa `wr`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ShapeTextWrap {
    /// Wrap text around the object.
    Around = 0,
    /// No text on either side of the shape (top and bottom wrapping).
    TopAndBottom = 1,
    /// Wrap text around an absolutely positioned object (square wrapping).
    Square = 2,
    /// Display as if the shape is not there (front of or behind text).
    None = 3,
    /// Wrap text tightly around the shape's contour on the left and right.
    Tight = 4,
    /// Wrap text tightly around the shape's contour on all sides.
    Through = 5,
}

impl ShapeTextWrap {
    fn from_bits(value: u8) -> Result<Self> {
        Ok(match value {
            0 => Self::Around,
            1 => Self::TopAndBottom,
            2 => Self::Square,
            3 => Self::None,
            4 => Self::Tight,
            5 => Self::Through,
            other => {
                return Err(DocError::InvalidFormat(format!(
                    "Invalid SPA text wrap style: {other}"
                )))
            },
        })
    }
}

/// Which sides of a floating shape allow wrapped text (Spa `wrk`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ShapeWrapSide {
    /// Allow text wrapping on both sides of the shape.
    Both = 0,
    /// Allow text wrapping only on the left side of the shape.
    Left = 1,
    /// Allow text wrapping only on the right side of the shape.
    Right = 2,
    /// Allow text wrapping only on the largest side of the shape.
    Largest = 3,
}

impl ShapeWrapSide {
    fn from_bits(value: u8) -> Result<Self> {
        Ok(match value {
            0 => Self::Both,
            1 => Self::Left,
            2 => Self::Right,
            3 => Self::Largest,
            other => {
                return Err(DocError::InvalidFormat(format!(
                    "Invalid SPA wrap side: {other}"
                )))
            },
        })
    }
}

/// Shape Positioning Attributes of one floating shape ([MS-DOC] 2.9.253).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Spa {
    /// Shape identifier; matches the `spid` of the shape's OfficeArtFSP.
    pub shape_id: u32,
    /// Left edge of the position rectangle in twips.
    pub left: i32,
    /// Top edge of the position rectangle in twips.
    pub top: i32,
    /// Right edge of the position rectangle in twips.
    pub right: i32,
    /// Bottom edge of the position rectangle in twips.
    pub bottom: i32,
    /// Horizontal position origin (`bx`).
    pub horizontal_origin: ShapeHorizontalOrigin,
    /// Vertical position origin (`by`).
    pub vertical_origin: ShapeVerticalOrigin,
    /// Text-wrapping style (`wr`).
    pub wrap: ShapeTextWrap,
    /// Wrap side restriction (`wrk`).
    pub wrap_side: ShapeWrapSide,
    /// Whether the shape appears behind the text (`fBelowText`).
    pub below_text: bool,
    /// Whether the anchor is locked to its paragraph (`fAnchorLock`).
    pub anchor_locked: bool,
}

impl Spa {
    /// Bit offset and width constants for the packed Spa flags field.
    pub(crate) const HORIZONTAL_ORIGIN_SHIFT: u16 = 1;
    pub(crate) const VERTICAL_ORIGIN_SHIFT: u16 = 3;
    pub(crate) const WRAP_SHIFT: u16 = 5;
    pub(crate) const WRAP_SIDE_SHIFT: u16 = 9;
    pub(crate) const BELOW_TEXT_BIT: u16 = 14;
    pub(crate) const ANCHOR_LOCK_BIT: u16 = 15;

    /// Parse one 26-byte Spa structure.
    pub fn parse(data: &[u8]) -> Result<Self> {
        if data.len() < SPA_LEN {
            return Err(DocError::InvalidFormat("Spa structure too short".to_string()));
        }
        let read_i32 = |offset: usize| -> i32 {
            i32::from_le_bytes(data[offset..offset + 4].try_into().unwrap_or([0; 4]))
        };
        let shape_id = read_i32(0) as u32;
        let left = read_i32(4);
        let top = read_i32(8);
        let right = read_i32(12);
        let bottom = read_i32(16);
        let flags = u16::from_le_bytes(data[20..22].try_into().unwrap_or([0; 2]));

        let masked = |shift: u16, width: u16| -> u8 {
            ((flags >> shift) & ((1u16 << width) - 1)) as u8
        };

        Ok(Self {
            shape_id,
            left,
            top,
            right,
            bottom,
            horizontal_origin: ShapeHorizontalOrigin::from_bits(masked(
                Self::HORIZONTAL_ORIGIN_SHIFT,
                2,
            ))?,
            vertical_origin: ShapeVerticalOrigin::from_bits(masked(
                Self::VERTICAL_ORIGIN_SHIFT,
                2,
            ))?,
            wrap: ShapeTextWrap::from_bits(masked(Self::WRAP_SHIFT, 4))?,
            wrap_side: ShapeWrapSide::from_bits(masked(Self::WRAP_SIDE_SHIFT, 4))?,
            below_text: flags & (1 << Self::BELOW_TEXT_BIT) != 0,
            anchor_locked: flags & (1 << Self::ANCHOR_LOCK_BIT) != 0,
        })
    }

    /// Width of the position rectangle in twips.
    pub fn width(&self) -> i32 {
        self.right - self.left
    }

    /// Height of the position rectangle in twips.
    pub fn height(&self) -> i32 {
        self.bottom - self.top
    }

    /// Serialize to the 26-byte Spa structure.
    pub fn to_bytes(&self) -> [u8; SPA_LEN] {
        let mut data = [0u8; SPA_LEN];
        data[0..4].copy_from_slice(&self.shape_id.to_le_bytes());
        data[4..8].copy_from_slice(&self.left.to_le_bytes());
        data[8..12].copy_from_slice(&self.top.to_le_bytes());
        data[12..16].copy_from_slice(&self.right.to_le_bytes());
        data[16..20].copy_from_slice(&self.bottom.to_le_bytes());
        let mut flags: u16 = 0;
        flags |= (self.horizontal_origin as u16) << Self::HORIZONTAL_ORIGIN_SHIFT;
        flags |= (self.vertical_origin as u16) << Self::VERTICAL_ORIGIN_SHIFT;
        flags |= (self.wrap as u16) << Self::WRAP_SHIFT;
        flags |= (self.wrap_side as u16) << Self::WRAP_SIDE_SHIFT;
        if self.below_text {
            flags |= 1 << Self::BELOW_TEXT_BIT;
        }
        if self.anchor_locked {
            flags |= 1 << Self::ANCHOR_LOCK_BIT;
        }
        data[20..22].copy_from_slice(&flags.to_le_bytes());
        // cTxbx (bytes 22..26) is undefined and stays zero.
        data
    }
}

/// One floating-shape anchor: the anchor character position and its
/// positioning attributes.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ShapeAnchor {
    /// Character position of the 0x0008 anchor character, relative to the
    /// start of the story that owns the PlcfSpa.
    pub cp: u32,
    /// Positioning attributes of the anchored shape.
    pub spa: Spa,
}

/// Parse a PlcfSpa ([MS-DOC] 2.8.27): `n + 1` CPs followed by `n` Spa records.
pub fn parse_plcf_spa(data: &[u8]) -> Result<Vec<ShapeAnchor>> {
    // lcb = 4 * (n + 1) + SPA_LEN * n  =>  n = (lcb - 4) / (4 + SPA_LEN)
    let stride = 4 + SPA_LEN;
    if data.len() < 4 || (data.len() - 4) % stride != 0 {
        return Err(DocError::InvalidFormat(
            "Invalid PlcfSpa length".to_string(),
        ));
    }
    let count = (data.len() - 4) / stride;
    let mut anchors = Vec::with_capacity(count);
    for index in 0..count {
        let cp = u32::from_le_bytes(
            data[index * 4..index * 4 + 4]
                .try_into()
                .unwrap_or([0; 4]),
        );
        let spa_start = (count + 1) * 4 + index * SPA_LEN;
        let spa = Spa::parse(&data[spa_start..spa_start + SPA_LEN])?;
        anchors.push(ShapeAnchor { cp, spa });
    }
    Ok(anchors)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn sample_spa_bytes() -> [u8; SPA_LEN] {
        let mut data = [0u8; SPA_LEN];
        data[0..4].copy_from_slice(&1029u32.to_le_bytes());
        data[4..8].copy_from_slice(&3060i32.to_le_bytes());
        data[8..12].copy_from_slice(&720i32.to_le_bytes());
        data[12..16].copy_from_slice(&4560i32.to_le_bytes());
        data[16..20].copy_from_slice(&1845i32.to_le_bytes());
        // bx=2 (column), by=2 (paragraph), wr=2 (square), fAnchorLock=1
        let flags: u16 = (2 << Spa::HORIZONTAL_ORIGIN_SHIFT)
            | (2 << Spa::VERTICAL_ORIGIN_SHIFT)
            | (2 << Spa::WRAP_SHIFT)
            | (1 << Spa::ANCHOR_LOCK_BIT);
        data[20..22].copy_from_slice(&flags.to_le_bytes());
        data
    }

    #[test]
    fn spa_parse_reads_all_fields() {
        let spa = Spa::parse(&sample_spa_bytes()).unwrap();
        assert_eq!(spa.shape_id, 1029);
        assert_eq!((spa.left, spa.top, spa.right, spa.bottom), (3060, 720, 4560, 1845));
        assert_eq!(spa.width(), 1500);
        assert_eq!(spa.height(), 1125);
        assert_eq!(spa.horizontal_origin, ShapeHorizontalOrigin::Column);
        assert_eq!(spa.vertical_origin, ShapeVerticalOrigin::Paragraph);
        assert_eq!(spa.wrap, ShapeTextWrap::Square);
        assert_eq!(spa.wrap_side, ShapeWrapSide::Both);
        assert!(!spa.below_text);
        assert!(spa.anchor_locked);
    }

    #[test]
    fn spa_to_bytes_round_trips_through_parse() {
        let spa = Spa::parse(&sample_spa_bytes()).unwrap();
        let bytes = spa.to_bytes();
        assert_eq!(&bytes[..], &sample_spa_bytes()[..]);
        let reparsed = Spa::parse(&bytes).unwrap();
        assert_eq!(reparsed, spa);
    }

    #[test]
    fn spa_parse_rejects_short_data() {
        assert!(Spa::parse(&[0u8; 10]).is_err());
    }

    #[test]
    fn plcf_spa_round_trip_count() {
        let mut plcf = Vec::new();
        plcf.extend_from_slice(&100u32.to_le_bytes());
        plcf.extend_from_slice(&200u32.to_le_bytes());
        plcf.extend_from_slice(&300u32.to_le_bytes());
        plcf.extend_from_slice(&sample_spa_bytes());
        plcf.extend_from_slice(&sample_spa_bytes());

        let anchors = parse_plcf_spa(&plcf).unwrap();
        assert_eq!(anchors.len(), 2);
        assert_eq!(anchors[0].cp, 100);
        assert_eq!(anchors[1].cp, 200);
        assert_eq!(anchors[1].spa.shape_id, 1029);

        assert!(parse_plcf_spa(&plcf[..plcf.len() - 1]).is_err());
        assert!(parse_plcf_spa(&[]).is_err());
    }
}
