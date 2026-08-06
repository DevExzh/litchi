//! Typed values for an MS-PPT `OfficeArtClientAnchor` record.

use crate::package::Result;

use super::validation;

pub(super) const SMALL_RECT_LEN: usize = 8;
pub(super) const RECT_LEN: usize = 16;

/// Resource limits for parsing an anchor record.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    /// Maximum accepted payload size, excluding the eight-byte OfficeArt header.
    pub max_payload_bytes: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_payload_bytes: RECT_LEN,
        }
    }
}

/// Exact rectangle representation carried on the wire.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Encoding {
    /// Eight-byte `SmallRectStruct` with signed 16-bit coordinates.
    Small,
    /// Sixteen-byte `RectStruct` with signed 32-bit coordinates.
    Full,
}

/// Compact `SmallRectStruct` coordinates in PowerPoint master units.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct SmallRect {
    pub(super) top: i16,
    pub(super) left: i16,
    pub(super) right: i16,
    pub(super) bottom: i16,
}

impl SmallRect {
    /// Construct checked bounds in geometric order.
    pub fn new(left: i16, top: i16, right: i16, bottom: i16) -> Result<Self> {
        validation::bounds(left.into(), top.into(), right.into(), bottom.into())?;
        Ok(Self {
            top,
            left,
            right,
            bottom,
        })
    }

    pub const fn left(self) -> i16 {
        self.left
    }

    pub const fn top(self) -> i16 {
        self.top
    }

    pub const fn right(self) -> i16 {
        self.right
    }

    pub const fn bottom(self) -> i16 {
        self.bottom
    }
}

/// Full-width `RectStruct` coordinates in PowerPoint master units.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Rect {
    pub(super) top: i32,
    pub(super) left: i32,
    pub(super) right: i32,
    pub(super) bottom: i32,
}

impl Rect {
    /// Construct checked bounds in geometric order.
    pub fn new(left: i32, top: i32, right: i32, bottom: i32) -> Result<Self> {
        validation::bounds(left, top, right, bottom)?;
        Ok(Self {
            top,
            left,
            right,
            bottom,
        })
    }

    pub const fn left(self) -> i32 {
        self.left
    }

    pub const fn top(self) -> i32 {
        self.top
    }

    pub const fn right(self) -> i32 {
        self.right
    }

    pub const fn bottom(self) -> i32 {
        self.bottom
    }
}

/// Variable MS-PPT anchor payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Data {
    Small(SmallRect),
    Full(Rect),
}

impl Data {
    pub const fn encoding(self) -> Encoding {
        match self {
            Self::Small(_) => Encoding::Small,
            Self::Full(_) => Encoding::Full,
        }
    }

    pub const fn encoded_len(self) -> usize {
        match self {
            Self::Small(_) => SMALL_RECT_LEN,
            Self::Full(_) => RECT_LEN,
        }
    }

    pub const fn left(self) -> i32 {
        match self {
            Self::Small(value) => value.left() as i32,
            Self::Full(value) => value.left(),
        }
    }

    pub const fn top(self) -> i32 {
        match self {
            Self::Small(value) => value.top() as i32,
            Self::Full(value) => value.top(),
        }
    }

    pub const fn right(self) -> i32 {
        match self {
            Self::Small(value) => value.right() as i32,
            Self::Full(value) => value.right(),
        }
    }

    pub const fn bottom(self) -> i32 {
        match self {
            Self::Small(value) => value.bottom() as i32,
            Self::Full(value) => value.bottom(),
        }
    }

    pub fn width(self) -> i64 {
        i64::from(self.right()) - i64::from(self.left())
    }

    pub fn height(self) -> i64 {
        i64::from(self.bottom()) - i64::from(self.top())
    }
}

/// One complete typed `OfficeArtClientAnchor` record.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Anchor {
    data: Data,
}

impl Anchor {
    pub const fn new(data: Data) -> Self {
        Self { data }
    }

    pub fn small(left: i16, top: i16, right: i16, bottom: i16) -> Result<Self> {
        Ok(Self::new(Data::Small(SmallRect::new(
            left, top, right, bottom,
        )?)))
    }

    pub fn full(left: i32, top: i32, right: i32, bottom: i32) -> Result<Self> {
        Ok(Self::new(Data::Full(Rect::new(left, top, right, bottom)?)))
    }

    pub const fn data(self) -> Data {
        self.data
    }

    pub const fn encoding(self) -> Encoding {
        self.data.encoding()
    }

    pub const fn left(self) -> i32 {
        self.data.left()
    }

    pub const fn top(self) -> i32 {
        self.data.top()
    }

    pub const fn right(self) -> i32 {
        self.data.right()
    }

    pub const fn bottom(self) -> i32 {
        self.data.bottom()
    }

    pub fn width(self) -> i64 {
        self.data.width()
    }

    pub fn height(self) -> i64 {
        self.data.height()
    }

    pub const fn encoded_len(self) -> usize {
        super::codec::HEADER_LEN + self.data.encoded_len()
    }
}
