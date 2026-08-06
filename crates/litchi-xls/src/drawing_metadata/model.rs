//! Typed worksheet-anchor values.

use std::io;

/// How a worksheet drawing responds when its underlying cells move or resize.
///
/// This is the two-bit `fMove`/`fSize` state from
/// `[MS-XLS] OfficeArtClientAnchorSheet`.  The reserved bit patterns are not
/// represented and are rejected by the decoder.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum AnchorBehavior {
    /// The drawing stays fixed while cells move or resize.
    Fixed,
    /// The drawing resizes with cells but does not move with them.
    Size,
    /// The drawing moves and resizes with cells.
    MoveAndSize,
}

impl AnchorBehavior {
    pub(crate) const fn wire_flags(self) -> u16 {
        match self {
            Self::Fixed => 0,
            Self::Size => 0b10,
            Self::MoveAndSize => 0b11,
        }
    }

    pub(crate) fn from_wire_flags(value: u16) -> io::Result<Self> {
        match value {
            0 => Ok(Self::Fixed),
            0b10 => Ok(Self::Size),
            0b11 => Ok(Self::MoveAndSize),
            _ => Err(super::validation::invalid(
                "worksheet anchor has reserved behavior flags",
            )),
        }
    }
}

/// One endpoint of an XLS cell-relative drawing anchor.
///
/// `x` is measured in 1/1024ths of the column width and `y` in 1/256ths of
/// the row height.  The signed wire representation is retained exactly,
/// including producer-specific negative offsets.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct AnchorPoint {
    column: u16,
    row: u16,
    x: i16,
    y: i16,
}

impl AnchorPoint {
    /// Creates a checked point in the BIFF8 worksheet grid.
    pub fn new(column: u16, row: u16, x: i16, y: i16) -> io::Result<Self> {
        let value = Self { column, row, x, y };
        value.validate()?;
        Ok(value)
    }

    /// Returns the zero-based worksheet column.
    pub const fn column(self) -> u16 {
        self.column
    }

    /// Returns the zero-based worksheet row.
    pub const fn row(self) -> u16 {
        self.row
    }

    /// Returns the signed horizontal offset in 1/1024ths of a column.
    pub const fn x(self) -> i16 {
        self.x
    }

    /// Returns the signed vertical offset in 1/256ths of a row.
    pub const fn y(self) -> i16 {
        self.y
    }

    pub(crate) fn validate(self) -> io::Result<()> {
        if self.column > 0x00FF {
            return Err(super::validation::invalid(
                "worksheet anchor column exceeds the BIFF8 grid",
            ));
        }
        Ok(())
    }

    pub(crate) const fn wire_fields(self) -> (u16, i16, u16, i16) {
        (self.column, self.x, self.row, self.y)
    }
}

/// The XLS `OfficeArtClientAnchorSheet` metadata attached to one drawing.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct SheetAnchor {
    behavior: AnchorBehavior,
    top_left: AnchorPoint,
    bottom_right: AnchorPoint,
}

impl SheetAnchor {
    /// Creates an anchor whose bounding-rectangle endpoints are ordered.
    pub fn new(
        top_left: AnchorPoint,
        bottom_right: AnchorPoint,
        behavior: AnchorBehavior,
    ) -> io::Result<Self> {
        let value = Self {
            behavior,
            top_left,
            bottom_right,
        };
        value.validate()?;
        Ok(value)
    }

    /// Returns the cell-change behavior.
    pub const fn behavior(self) -> AnchorBehavior {
        self.behavior
    }

    /// Returns the top-left bounding-rectangle endpoint.
    pub const fn top_left(self) -> AnchorPoint {
        self.top_left
    }

    /// Returns the bottom-right bounding-rectangle endpoint.
    pub const fn bottom_right(self) -> AnchorPoint {
        self.bottom_right
    }

    pub(crate) fn validate(self) -> io::Result<()> {
        self.top_left.validate()?;
        self.bottom_right.validate()?;
        let horizontal = (self.top_left.column, self.top_left.x)
            < (self.bottom_right.column, self.bottom_right.x);
        let vertical =
            (self.top_left.row, self.top_left.y) < (self.bottom_right.row, self.bottom_right.y);
        if !horizontal || !vertical {
            return Err(super::validation::invalid(
                "worksheet anchor endpoints are not strictly ordered",
            ));
        }
        Ok(())
    }
}
