//! Semantic `DrawingML` coordinate values.

use super::codec::ParseError;

/// Smallest unqualified `a:ST_Coordinate` value.
pub const MIN_EMU: i64 = -27_273_042_329_600;
/// Largest unqualified `a:ST_Coordinate` value.
pub const MAX_EMU: i64 = 27_273_042_316_900;
/// Maximum accepted coordinate spelling in bytes.
///
/// The schema's decimal grammar is unbounded. This finite limit prevents an
/// untrusted XML attribute from becoming an unbounded allocation while still
/// leaving ample precision for producer-authored physical measurements.
pub const MAX_BYTES: usize = 128;

/// Unit suffix accepted by `s:ST_UniversalMeasure`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum Unit {
    Mm,
    Cm,
    Inch,
    Pt,
    Pc,
    Pi,
}

impl Unit {
    /// Return the exact OOXML suffix.
    #[inline]
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Mm => "mm",
            Self::Cm => "cm",
            Self::Inch => "in",
            Self::Pt => "pt",
            Self::Pc => "pc",
            Self::Pi => "pi",
        }
    }
}

#[derive(Clone, PartialEq, Eq, Hash)]
pub(super) enum Repr {
    Emu(i64),
    Measure {
        lexical: Box<str>,
        number_len: usize,
        unit: Unit,
    },
}

/// An exact `a:ST_Coordinate` value.
///
/// Unqualified coordinates are EMUs and are range-checked. Unit-bearing
/// coordinates retain an exact canonical decimal rather than passing through
/// floating point.
#[derive(Clone, PartialEq, Eq, Hash)]
#[must_use]
pub struct Coordinate(pub(super) Repr);

impl Coordinate {
    /// Zero EMUs.
    pub const ZERO: Self = Self(Repr::Emu(0));

    /// Construct a checked unqualified EMU coordinate.
    pub const fn emu(value: i64) -> Result<Self, ParseError> {
        if value < MIN_EMU || value > MAX_EMU {
            Err(ParseError::OutOfRange { value })
        } else {
            Ok(Self(Repr::Emu(value)))
        }
    }

    /// Return the EMU value for an unqualified coordinate.
    #[inline]
    #[must_use]
    pub const fn as_emu(&self) -> Option<i64> {
        match &self.0 {
            Repr::Emu(value) => Some(*value),
            Repr::Measure { .. } => None,
        }
    }

    /// Return the exact decimal component for a physical measurement.
    #[must_use]
    pub fn number(&self) -> Option<&str> {
        match &self.0 {
            Repr::Emu(_) => None,
            Repr::Measure {
                lexical,
                number_len,
                ..
            } => lexical.get(..*number_len),
        }
    }

    /// Return the unit for a physical measurement.
    #[inline]
    #[must_use]
    pub const fn unit(&self) -> Option<Unit> {
        match &self.0 {
            Repr::Emu(_) => None,
            Repr::Measure { unit, .. } => Some(*unit),
        }
    }
}

impl Default for Coordinate {
    fn default() -> Self {
        Self::ZERO
    }
}

impl From<i32> for Coordinate {
    fn from(value: i32) -> Self {
        Self(Repr::Emu(i64::from(value)))
    }
}

/// An exact `a:ST_PositiveCoordinate` suitable for `DrawingML` extents.
///
/// Despite its name, the XSD type is an integer restriction with an inclusive
/// lower bound of zero. It does not accept unit-bearing measurements.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
#[repr(transparent)]
#[must_use]
pub struct Extent(i64);

impl Extent {
    /// Zero EMUs, which is valid for `a:ST_PositiveCoordinate`.
    pub const ZERO: Self = Self(0);

    /// Construct a checked extent in EMUs.
    pub const fn emu(value: i64) -> Result<Self, ParseError> {
        if value < 0 || value > MAX_EMU {
            Err(ParseError::ExtentOutOfRange { value })
        } else {
            Ok(Self(value))
        }
    }

    /// Return the checked EMU value.
    #[inline]
    #[must_use]
    pub const fn as_emu(self) -> i64 {
        self.0
    }
}

impl Default for Extent {
    fn default() -> Self {
        Self::ZERO
    }
}

impl From<u32> for Extent {
    fn from(value: u32) -> Self {
        Self(i64::from(value))
    }
}

impl From<Extent> for i64 {
    fn from(value: Extent) -> Self {
        value.0
    }
}
