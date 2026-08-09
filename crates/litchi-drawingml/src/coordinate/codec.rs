//! Exact lexical codecs for `DrawingML` coordinate domains.

use std::fmt;
use std::str::FromStr;

use super::model::{Coordinate, Extent, MAX_BYTES, MAX_EMU, MIN_EMU, Repr, Unit};

/// Failure to construct or parse a `DrawingML` coordinate.
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

impl Unit {
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

impl Coordinate {
    /// Construct an exact decimal physical measurement.
    /// # Errors
    ///
    /// Returns an error when input violates DrawingML constraints, exceeds a configured
    /// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
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
    /// # Errors
    ///
    /// Returns an error when input violates DrawingML constraints, exceeds a configured
    /// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
    pub fn parse(value: &str) -> Result<Self, ParseError> {
        if value.len() > MAX_BYTES {
            return Err(ParseError::TooLong {
                len: value.len(),
                max: MAX_BYTES,
            });
        }
        Self::parse_owned(value.to_owned())
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
            .map_err(|_error| ParseError::InvalidNumber)?;
        Self::emu(parsed)
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

impl fmt::Display for Extent {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        self.as_emu().fmt(formatter)
    }
}

impl Extent {
    /// Parse the integer lexical space of `a:ST_PositiveCoordinate`.
    #[inline]
    /// # Errors
    ///
    /// Returns an error when input violates DrawingML constraints, exceeds a configured
    /// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
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
            .map_err(|_error| ParseError::InvalidExtent)?;
        Self::emu(value)
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
