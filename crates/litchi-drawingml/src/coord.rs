//! Exact, checked DrawingML coordinates.
//!
//! [`Coordinate`] models the complete `a:ST_Coordinate` union: either a
//! bounded unqualified EMU integer or an exact `s:ST_UniversalMeasure` value.
//! [`Extent`] models the integer-only `a:ST_PositiveCoordinate` restriction.

use std::fmt;
use std::str::FromStr;

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

    fn from_suffix(value: &str) -> Result<Self, ParseError> {
        match value {
            "mm" => Ok(Self::Mm),
            "cm" => Ok(Self::Cm),
            "in" => Ok(Self::Inch),
            "pt" => Ok(Self::Pt),
            "pc" => Ok(Self::Pc),
            "pi" => Ok(Self::Pi),
            _ => Err(ParseError::InvalidUnit),
        }
    }
}

impl fmt::Display for Unit {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}

impl FromStr for Unit {
    type Err = ParseError;

    fn from_str(value: &str) -> Result<Self, Self::Err> {
        Self::from_suffix(value)
    }
}

/// Failure to construct or parse a DrawingML coordinate.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ParseError {
    Empty,
    TooLong { len: usize, max: usize },
    InvalidNumber,
    InvalidExtent,
    InvalidUnit,
    ExtentOutOfRange { value: i64 },
    OutOfRange { value: i64 },
}

impl fmt::Display for ParseError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Empty => formatter.write_str("DrawingML coordinate is empty"),
            Self::TooLong { len, max } => {
                write!(
                    formatter,
                    "DrawingML coordinate uses {len} bytes; the limit is {max}"
                )
            },
            Self::InvalidNumber => formatter.write_str(
                "DrawingML coordinate must be a bounded integer or an exact decimal measure",
            ),
            Self::InvalidExtent => write!(
                formatter,
                "DrawingML extent must be an integer between 0 and {MAX_EMU}"
            ),
            Self::InvalidUnit => {
                formatter.write_str("DrawingML coordinate unit must be mm, cm, in, pt, pc, or pi")
            },
            Self::ExtentOutOfRange { value } => write!(
                formatter,
                "DrawingML extent {value} is outside 0..={MAX_EMU}"
            ),
            Self::OutOfRange { value } => write!(
                formatter,
                "DrawingML coordinate {value} is outside {MIN_EMU}..={MAX_EMU}"
            ),
        }
    }
}

impl std::error::Error for ParseError {}

#[derive(Clone, PartialEq, Eq, Hash)]
enum Repr {
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
pub struct Coordinate(Repr);

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

    /// Construct an exact decimal physical measurement.
    pub fn measure(number: &str, unit: Unit) -> Result<Self, ParseError> {
        if number.is_empty() {
            return Err(ParseError::Empty);
        }
        let len = number
            .len()
            .checked_add(unit.as_str().len())
            .ok_or(ParseError::TooLong {
                len: usize::MAX,
                max: MAX_BYTES,
            })?;
        if len > MAX_BYTES {
            return Err(ParseError::TooLong {
                len,
                max: MAX_BYTES,
            });
        }
        let mut lexical = String::with_capacity(len);
        lexical.push_str(number);
        normalize_measure(lexical, unit)
    }

    /// Parse either member of the `a:ST_Coordinate` union.
    #[inline]
    pub fn parse(value: &str) -> Result<Self, ParseError> {
        if value.len() > MAX_BYTES {
            return Err(ParseError::TooLong {
                len: value.len(),
                max: MAX_BYTES,
            });
        }
        Self::parse_owned(value.to_owned())
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

    fn parse_owned(mut value: String) -> Result<Self, ParseError> {
        if value.is_empty() {
            return Err(ParseError::Empty);
        }
        if value.len() > MAX_BYTES {
            return Err(ParseError::TooLong {
                len: value.len(),
                max: MAX_BYTES,
            });
        }

        if let Some(unit_start) = value.find(char::is_alphabetic) {
            let unit = Unit::from_suffix(&value[unit_start..])?;
            value.truncate(unit_start);
            return normalize_measure(value, unit);
        }

        let value = value.trim_matches(xsd_whitespace);
        if value.is_empty() {
            return Err(ParseError::InvalidNumber);
        }
        let parsed = value
            .parse::<i64>()
            .map_err(|_| ParseError::InvalidNumber)?;
        Self::emu(parsed)
    }
}

impl Default for Coordinate {
    fn default() -> Self {
        Self::ZERO
    }
}

impl fmt::Debug for Coordinate {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_tuple("Coordinate")
            .field(&format_args!("{self}"))
            .finish()
    }
}

impl fmt::Display for Coordinate {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match &self.0 {
            Repr::Emu(value) => value.fmt(formatter),
            Repr::Measure { lexical, .. } => formatter.write_str(lexical),
        }
    }
}

impl FromStr for Coordinate {
    type Err = ParseError;

    fn from_str(value: &str) -> Result<Self, Self::Err> {
        Self::parse(value)
    }
}

impl TryFrom<String> for Coordinate {
    type Error = ParseError;

    fn try_from(value: String) -> Result<Self, Self::Error> {
        Self::parse_owned(value)
    }
}

impl TryFrom<&str> for Coordinate {
    type Error = ParseError;

    fn try_from(value: &str) -> Result<Self, Self::Error> {
        Self::parse(value)
    }
}

impl TryFrom<i64> for Coordinate {
    type Error = ParseError;

    fn try_from(value: i64) -> Result<Self, Self::Error> {
        Self::emu(value)
    }
}

impl From<i32> for Coordinate {
    fn from(value: i32) -> Self {
        Self(Repr::Emu(i64::from(value)))
    }
}

/// An exact `a:ST_PositiveCoordinate` suitable for DrawingML extents.
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

    /// Parse the integer lexical space of `a:ST_PositiveCoordinate`.
    #[inline]
    pub fn parse(value: &str) -> Result<Self, ParseError> {
        if value.len() > MAX_BYTES {
            return Err(ParseError::TooLong {
                len: value.len(),
                max: MAX_BYTES,
            });
        }
        let value = value.trim_matches(xsd_whitespace);
        if value.is_empty() {
            return Err(ParseError::InvalidExtent);
        }
        let value = value
            .parse::<i64>()
            .map_err(|_| ParseError::InvalidExtent)?;
        Self::emu(value)
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

impl fmt::Display for Extent {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        self.0.fmt(formatter)
    }
}

impl FromStr for Extent {
    type Err = ParseError;

    fn from_str(value: &str) -> Result<Self, Self::Err> {
        Self::parse(value)
    }
}

impl TryFrom<String> for Extent {
    type Error = ParseError;

    fn try_from(value: String) -> Result<Self, Self::Error> {
        Self::parse(&value)
    }
}

impl TryFrom<&str> for Extent {
    type Error = ParseError;

    fn try_from(value: &str) -> Result<Self, Self::Error> {
        Self::parse(value)
    }
}

impl TryFrom<i64> for Extent {
    type Error = ParseError;

    fn try_from(value: i64) -> Result<Self, Self::Error> {
        Self::emu(value)
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

fn normalize_measure(mut number: String, unit: Unit) -> Result<Coordinate, ParseError> {
    let negative = number.starts_with('-');
    let digits_start = usize::from(negative);
    if number.len() == digits_start || number.starts_with('+') {
        return Err(ParseError::InvalidNumber);
    }

    let mut decimal_index = None;
    for (index, byte) in number.as_bytes()[digits_start..]
        .iter()
        .copied()
        .enumerate()
    {
        match byte {
            b'0'..=b'9' => {},
            b'.' if decimal_index.is_none() => decimal_index = Some(index),
            _ => return Err(ParseError::InvalidNumber),
        }
    }
    let unsigned_len = number.len() - digits_start;
    if decimal_index == Some(0) || decimal_index == unsigned_len.checked_sub(1) {
        return Err(ParseError::InvalidNumber);
    }

    if decimal_index.is_some() {
        while number.ends_with('0') {
            number.pop();
        }
        if number.ends_with('.') {
            number.pop();
        }
    }

    let whole_end = number[digits_start..]
        .find('.')
        .map_or(number.len(), |index| digits_start + index);
    let leading_zeroes = number.as_bytes()[digits_start..whole_end]
        .iter()
        .take_while(|byte| **byte == b'0')
        .count();
    let whole_len = whole_end - digits_start;
    let remove = leading_zeroes.min(whole_len.saturating_sub(1));
    if remove != 0 {
        number.drain(digits_start..digits_start + remove);
    }
    if number == "-0" {
        number.remove(0);
    }

    let number_len = number.len();
    number.push_str(unit.as_str());
    if number.len() > MAX_BYTES {
        return Err(ParseError::TooLong {
            len: number.len(),
            max: MAX_BYTES,
        });
    }
    Ok(Coordinate(Repr::Measure {
        lexical: number.into_boxed_str(),
        number_len,
        unit,
    }))
}

const fn xsd_whitespace(character: char) -> bool {
    matches!(character, ' ' | '\t' | '\n' | '\r')
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::mem::size_of;

    #[test]
    fn unqualified_coordinates_are_canonical_and_bounded() {
        assert_eq!(Coordinate::parse(" +00042 ").unwrap().as_emu(), Some(42));
        assert_eq!(Coordinate::parse("-0").unwrap(), Coordinate::ZERO);
        assert_eq!(
            Coordinate::emu(MIN_EMU).unwrap().to_string(),
            MIN_EMU.to_string()
        );
        assert_eq!(
            Coordinate::emu(MAX_EMU).unwrap().to_string(),
            MAX_EMU.to_string()
        );
        assert!(matches!(
            Coordinate::emu(MIN_EMU - 1),
            Err(ParseError::OutOfRange { .. })
        ));
        assert!(matches!(
            Coordinate::emu(MAX_EMU + 1),
            Err(ParseError::OutOfRange { .. })
        ));
    }

    #[test]
    fn universal_measures_cover_every_unit_without_floating_point() {
        for (source, canonical, unit) in [
            ("001.2500mm", "1.25mm", Unit::Mm),
            ("-2cm", "-2cm", Unit::Cm),
            ("3in", "3in", Unit::Inch),
            ("4pt", "4pt", Unit::Pt),
            ("5pc", "5pc", Unit::Pc),
            ("6pi", "6pi", Unit::Pi),
        ] {
            let coordinate = Coordinate::parse(source).unwrap();
            assert_eq!(coordinate.to_string(), canonical);
            assert_eq!(coordinate.unit(), Some(unit));
        }
        assert_eq!(Coordinate::parse("-0.000mm").unwrap().to_string(), "0mm");
        assert_eq!(
            Coordinate::measure("1.25", Unit::Cm).unwrap().number(),
            Some("1.25")
        );
    }

    #[test]
    fn malformed_or_unbounded_spellings_are_rejected() {
        for invalid in [
            "", ".1mm", "1.mm", "+1mm", "1e2mm", "1 px", "1MM", "--1", "1.0",
        ] {
            assert!(Coordinate::parse(invalid).is_err(), "accepted {invalid:?}");
        }
        assert!(matches!(
            Coordinate::parse(&format!("{}mm", "1".repeat(MAX_BYTES))),
            Err(ParseError::TooLong { .. })
        ));
        let oversized = "1".repeat(MAX_BYTES + 1);
        assert_eq!(
            oversized.parse::<Coordinate>(),
            Err(ParseError::TooLong {
                len: MAX_BYTES + 1,
                max: MAX_BYTES,
            })
        );
        assert_eq!(
            Coordinate::try_from(oversized),
            Err(ParseError::TooLong {
                len: MAX_BYTES + 1,
                max: MAX_BYTES,
            })
        );
        assert!("".parse::<Unit>().is_err());
    }

    #[test]
    fn extents_are_nonnegative_bounded_and_representation_compact() {
        assert_eq!(size_of::<Extent>(), size_of::<i64>());
        assert_eq!(Extent::ZERO.as_emu(), 0);
        assert_eq!(Extent::emu(0).unwrap(), Extent::ZERO);
        assert_eq!(Extent::emu(1).unwrap().as_emu(), 1);
        assert_eq!(
            Extent::emu(MAX_EMU).unwrap().to_string(),
            MAX_EMU.to_string()
        );
        assert_eq!(
            Extent::emu(-1),
            Err(ParseError::ExtentOutOfRange { value: -1 })
        );
        assert!(matches!(
            Extent::emu(MAX_EMU + 1),
            Err(ParseError::ExtentOutOfRange { .. })
        ));
    }

    #[test]
    fn extents_accept_only_the_exact_integer_lexical_space() {
        assert_eq!(Extent::parse(" +00042 ").unwrap().as_emu(), 42);
        assert_eq!(Extent::parse("-0").unwrap(), Extent::ZERO);
        assert_eq!(Extent::from(7_u32).as_emu(), 7);
        assert_eq!(i64::from(Extent::try_from(9_i64).unwrap()), 9);
        assert_eq!(Extent::parse("1cm"), Err(ParseError::InvalidExtent));
        assert_eq!(
            ParseError::InvalidExtent.to_string(),
            format!("DrawingML extent must be an integer between 0 and {MAX_EMU}")
        );
        for invalid in ["-1", "0mm", "1.25cm", "1.0", "1e2"] {
            assert!(Extent::parse(invalid).is_err(), "accepted {invalid:?}");
        }
    }
}
