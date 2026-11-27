//! Typed worksheet-anchor values.

use std::io;

/// How a worksheet drawing responds when its underlying cells move or resize.
///
/// This is the two-bit `fMove`/`fSize` state from
/// `[MS-XLS]` 2.5.193 `OfficeArtClientAnchorSheet`. Each flag is set when the
/// shape is *kept intact* as cells move (`fMove`, bit 0) or resize (`fSize`,
/// bit 1), and `fSize` MUST be set when `fMove` is. The variants follow
/// Apache POI's `ClientAnchor.AnchorType` (0, 2 and 3) and ECMA-376
/// `ST_EditAs` (`twoCell`, `oneCell`, `absolute`). The reserved bit patterns
/// are not represented and are rejected by the decoder.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum AnchorBehavior {
    /// The drawing moves and resizes with its cells (both flags clear).
    MoveAndResize,
    /// The drawing moves with its cells but keeps its size (`fSize` only).
    MoveWithoutResize,
    /// The drawing neither moves nor resizes with its cells (both flags set).
    Absolute,
}

impl AnchorBehavior {
    pub(crate) const fn wire_flags(self) -> u16 {
        match self {
            Self::MoveAndResize => 0,
            Self::MoveWithoutResize => 0b10,
            Self::Absolute => 0b11,
        }
    }

    pub(crate) fn from_wire_flags(value: u16) -> io::Result<Self> {
        match value {
            0 => Ok(Self::MoveAndResize),
            0b10 => Ok(Self::MoveWithoutResize),
            0b11 => Ok(Self::Absolute),
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
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn new(column: u16, row: u16, x: i16, y: i16) -> io::Result<Self> {
        let value = Self { column, row, x, y };
        value.validate()?;
        Ok(value)
    }

    /// Returns the zero-based worksheet column.
    #[must_use]
    pub const fn column(self) -> u16 {
        self.column
    }

    /// Returns the zero-based worksheet row.
    #[must_use]
    pub const fn row(self) -> u16 {
        self.row
    }

    /// Returns the signed horizontal offset in 1/1024ths of a column.
    #[must_use]
    pub const fn x(self) -> i16 {
        self.x
    }

    /// Returns the signed vertical offset in 1/256ths of a row.
    #[must_use]
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
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
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
    #[must_use]
    pub const fn behavior(self) -> AnchorBehavior {
        self.behavior
    }

    /// Returns the top-left bounding-rectangle endpoint.
    #[must_use]
    pub const fn top_left(self) -> AnchorPoint {
        self.top_left
    }

    /// Returns the bottom-right bounding-rectangle endpoint.
    #[must_use]
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
