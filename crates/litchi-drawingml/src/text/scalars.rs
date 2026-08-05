//! Checked numeric and coordinate domains used by DrawingML text.

use std::fmt;
use std::str::FromStr;

use crate::coordinate::{Coordinate, Unit};

use super::codec::ParseError;

/// A checked `ST_Coordinate32` value.
///
/// Unqualified values are i32 EMUs. Unit-bearing values retain their exact
/// decimal spelling and do not pass through floating point.
#[derive(Clone, PartialEq, Eq, Hash)]
#[must_use]
pub struct Coordinate32(Coordinate);

impl Coordinate32 {
    /// Zero EMUs.
    pub const ZERO: Self = Self(Coordinate::ZERO);

    /// Construct an unqualified EMU value. Every i32 is schema-valid.
    #[inline]
    pub fn emu(value: i32) -> Self {
        Self(Coordinate::from(value))
    }

    /// Construct an exact physical measurement.
    pub fn measure(number: &str, unit: Unit) -> Result<Self, ParseError> {
        Ok(Self(Coordinate::measure(number, unit)?))
    }

    /// Return the unqualified EMU value, or `None` for a physical measure.
    #[must_use]
    pub fn as_emu(&self) -> Option<i32> {
        self.0.as_emu().and_then(|value| i32::try_from(value).ok())
    }

    /// Borrow the complete coordinate representation.
    pub const fn as_coordinate(&self) -> &Coordinate {
        &self.0
    }

    fn checked(coordinate: Coordinate) -> Result<Self, ParseError> {
        if let Some(value) = coordinate.as_emu()
            && i32::try_from(value).is_err()
        {
            return Err(ParseError::OutOfRange {
                domain: "DrawingML Coordinate32",
                value,
                min: i64::from(i32::MIN),
                max: i64::from(i32::MAX),
            });
        }
        Ok(Self(coordinate))
    }
}

impl Default for Coordinate32 {
    fn default() -> Self {
        Self::ZERO
    }
}

impl fmt::Debug for Coordinate32 {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_tuple("Coordinate32")
            .field(&format_args!("{self}"))
            .finish()
    }
}

impl fmt::Display for Coordinate32 {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        self.0.fmt(formatter)
    }
}

impl FromStr for Coordinate32 {
    type Err = ParseError;

    fn from_str(value: &str) -> Result<Self, Self::Err> {
        Self::checked(value.parse()?)
    }
}

impl TryFrom<String> for Coordinate32 {
    type Error = ParseError;

    fn try_from(value: String) -> Result<Self, Self::Error> {
        Self::checked(Coordinate::try_from(value)?)
    }
}

impl TryFrom<&str> for Coordinate32 {
    type Error = ParseError;

    fn try_from(value: &str) -> Result<Self, Self::Error> {
        value.parse()
    }
}

impl TryFrom<i64> for Coordinate32 {
    type Error = ParseError;

    fn try_from(value: i64) -> Result<Self, Self::Error> {
        let narrowed = i32::try_from(value).map_err(|_| ParseError::OutOfRange {
            domain: "DrawingML Coordinate32",
            value,
            min: i64::from(i32::MIN),
            max: i64::from(i32::MAX),
        })?;
        Ok(Self::emu(narrowed))
    }
}

impl From<i32> for Coordinate32 {
    fn from(value: i32) -> Self {
        Self::emu(value)
    }
}

impl From<Coordinate32> for Coordinate {
    fn from(value: Coordinate32) -> Self {
        value.0
    }
}

/// Checked text-column count (`ST_TextColumnCount`, 1..=16).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
#[repr(transparent)]
#[must_use]
pub struct Columns(u8);

impl Columns {
    pub const MIN: u8 = 1;
    pub const MAX: u8 = 16;
    pub const ONE: Self = Self(Self::MIN);

    /// Construct a checked column count.
    pub const fn new(value: u8) -> Result<Self, ParseError> {
        if value < Self::MIN || value > Self::MAX {
            Err(ParseError::OutOfRange {
                domain: "DrawingML text column count",
                value: value as i64,
                min: Self::MIN as i64,
                max: Self::MAX as i64,
            })
        } else {
            Ok(Self(value))
        }
    }

    /// Return the checked count.
    #[must_use]
    pub const fn get(self) -> u8 {
        self.0
    }
}

impl Default for Columns {
    fn default() -> Self {
        Self::ONE
    }
}

impl fmt::Display for Columns {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        self.0.fmt(formatter)
    }
}

impl FromStr for Columns {
    type Err = ParseError;

    fn from_str(value: &str) -> Result<Self, Self::Err> {
        let parsed = value
            .trim_matches(super::codec::xsd_whitespace)
            .parse::<u32>()
            .map_err(|_| ParseError::InvalidNumber("DrawingML text column count"))?;
        Self::try_from(parsed)
    }
}

impl TryFrom<u32> for Columns {
    type Error = ParseError;

    fn try_from(value: u32) -> Result<Self, Self::Error> {
        let narrowed = u8::try_from(value).map_err(|_| ParseError::OutOfRange {
            domain: "DrawingML text column count",
            value: i64::from(value),
            min: i64::from(Self::MIN),
            max: i64::from(Self::MAX),
        })?;
        Self::new(narrowed)
    }
}

/// Font size in hundredths of a point (`ST_TextFontSize`, 100..=400000).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
#[repr(transparent)]
#[must_use]
pub struct TextSize(u32);

impl TextSize {
    pub const MIN: u32 = 100;
    pub const MAX: u32 = 400_000;

    /// Construct a checked font size in hundredths of a point.
    pub const fn new(value: u32) -> Result<Self, ParseError> {
        if value < Self::MIN || value > Self::MAX {
            Err(ParseError::OutOfRange {
                domain: "DrawingML text size",
                value: value as i64,
                min: Self::MIN as i64,
                max: Self::MAX as i64,
            })
        } else {
            Ok(Self(value))
        }
    }

    /// Return the size in hundredths of a point.
    #[must_use]
    pub const fn get(self) -> u32 {
        self.0
    }
}

impl fmt::Display for TextSize {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        self.0.fmt(formatter)
    }
}

impl FromStr for TextSize {
    type Err = ParseError;

    fn from_str(value: &str) -> Result<Self, Self::Err> {
        let parsed = value
            .trim_matches(super::codec::xsd_whitespace)
            .parse::<u32>()
            .map_err(|_| ParseError::InvalidNumber("DrawingML text size"))?;
        Self::new(parsed)
    }
}

impl TryFrom<u32> for TextSize {
    type Error = ParseError;

    fn try_from(value: u32) -> Result<Self, Self::Error> {
        Self::new(value)
    }
}
