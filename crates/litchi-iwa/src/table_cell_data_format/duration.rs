//! Strongly typed settings for native Duration table cells.

use super::TableCellDataFormat;
use crate::{Error, Result};

/// Presentation style used by iWork's Duration formatter.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum TableCellDurationStyle {
    /// Display the selected units as colon-separated values.
    Colon,
    /// Display compact unit symbols such as `1h 2m 3s`.
    #[default]
    Abbreviated,
    /// Display complete unit names such as `1 hour 2 minutes`.
    FullNames,
}

/// A unit supported by iWork's Duration formatter.
///
/// The discriminants match the native unit bit values stored in iWork
/// archives. They are intentionally not exposed as untyped integers.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
#[repr(u32)]
pub enum TableCellDurationUnit {
    Weeks = 1,
    Days = 2,
    Hours = 4,
    Minutes = 8,
    Seconds = 16,
    Milliseconds = 32,
}

/// Inclusive, contiguous range of units displayed by a Duration formatter.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct TableCellDurationUnitRange {
    largest: TableCellDurationUnit,
    smallest: TableCellDurationUnit,
}

impl TableCellDurationUnitRange {
    /// Construct a range from its largest unit through its smallest unit.
    pub fn new(largest: TableCellDurationUnit, smallest: TableCellDurationUnit) -> Result<Self> {
        if largest > smallest {
            return Err(Error::InvalidFormat(
                "Duration largest unit must not be smaller than its smallest unit".to_owned(),
            ));
        }
        Ok(Self { largest, smallest })
    }

    /// Construct the complete native range from weeks through milliseconds.
    pub const fn all() -> Self {
        Self {
            largest: TableCellDurationUnit::Weeks,
            smallest: TableCellDurationUnit::Milliseconds,
        }
    }

    /// Construct the common range from hours through milliseconds.
    pub const fn hours_to_milliseconds() -> Self {
        Self {
            largest: TableCellDurationUnit::Hours,
            smallest: TableCellDurationUnit::Milliseconds,
        }
    }

    /// Largest unit included in the range.
    pub const fn largest(self) -> TableCellDurationUnit {
        self.largest
    }

    /// Smallest unit included in the range.
    pub const fn smallest(self) -> TableCellDurationUnit {
        self.smallest
    }
}

impl Default for TableCellDurationUnitRange {
    fn default() -> Self {
        Self::all()
    }
}

/// Whether iWork chooses visible Duration units or uses a fixed range.
///
/// Automatic native formats still persist their most recently inferred range,
/// so both variants retain a range for lossless archive round-tripping.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum TableCellDurationUnits {
    Automatic(TableCellDurationUnitRange),
    Custom(TableCellDurationUnitRange),
}

impl TableCellDurationUnits {
    /// Inclusive unit range persisted by the native formatter.
    pub const fn range(self) -> TableCellDurationUnitRange {
        match self {
            Self::Automatic(range) | Self::Custom(range) => range,
        }
    }
}

impl Default for TableCellDurationUnits {
    fn default() -> Self {
        Self::Automatic(TableCellDurationUnitRange::all())
    }
}

/// Native Duration table-cell format.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub struct TableCellDurationFormat {
    style: TableCellDurationStyle,
    units: TableCellDurationUnits,
}

impl TableCellDurationFormat {
    /// Construct a Duration format from a presentation style and unit mode.
    pub const fn new(style: TableCellDurationStyle, units: TableCellDurationUnits) -> Self {
        Self { style, units }
    }

    /// Construct an automatic-unit Duration format.
    pub const fn automatic(style: TableCellDurationStyle) -> Self {
        Self::new(
            style,
            TableCellDurationUnits::Automatic(TableCellDurationUnitRange::all()),
        )
    }

    /// Construct a fixed-unit Duration format.
    pub const fn custom(style: TableCellDurationStyle, range: TableCellDurationUnitRange) -> Self {
        Self::new(style, TableCellDurationUnits::Custom(range))
    }

    /// Presentation style.
    pub const fn style(self) -> TableCellDurationStyle {
        self.style
    }

    /// Automatic or fixed unit selection, including its persisted range.
    pub const fn units(self) -> TableCellDurationUnits {
        self.units
    }
}

impl From<TableCellDurationFormat> for TableCellDataFormat {
    fn from(value: TableCellDurationFormat) -> Self {
        Self::Duration(value)
    }
}
