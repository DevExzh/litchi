//! Checked, format-neutral DrawingML text primitives.
//!
//! The types in this module model closed schema domains directly.  Host
//! formats can therefore share one compact public vocabulary while retaining
//! dialect-exact codecs where WordprocessingML and DrawingML spell the same
//! semantic value differently.

use std::fmt;
use std::str::FromStr;

use crate::coord::{Coordinate, Unit};

pub mod body;

/// Failure to parse or construct a DrawingML text primitive.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum ParseError {
    /// A closed-domain token is not a schema member.
    InvalidToken(&'static str),
    /// A boolean is not an XML Schema boolean.
    InvalidBool,
    /// A numeric value has an invalid lexical form.
    InvalidNumber(&'static str),
    /// A numeric value is outside its schema bounds.
    OutOfRange {
        /// Name of the bounded domain.
        domain: &'static str,
        /// Parsed value.
        value: i64,
        /// Inclusive lower bound.
        min: i64,
        /// Inclusive upper bound.
        max: i64,
    },
    /// The underlying coordinate spelling is invalid.
    Coordinate(crate::coord::ParseError),
}

impl fmt::Display for ParseError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::InvalidToken(domain) => write!(formatter, "invalid {domain} token"),
            Self::InvalidBool => formatter.write_str("invalid XML Schema boolean"),
            Self::InvalidNumber(domain) => write!(formatter, "invalid {domain} number"),
            Self::OutOfRange {
                domain,
                value,
                min,
                max,
            } => write!(formatter, "{domain} value {value} is outside {min}..={max}"),
            Self::Coordinate(error) => error.fmt(formatter),
        }
    }
}

impl std::error::Error for ParseError {}

impl From<crate::coord::ParseError> for ParseError {
    fn from(error: crate::coord::ParseError) -> Self {
        Self::Coordinate(error)
    }
}

/// Vertical anchoring within a text body (`ST_TextAnchoringType`).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum Anchor {
    /// Start at the top inset (`t`).
    #[default]
    Top,
    /// Center vertically (`ctr`).
    Center,
    /// End at the bottom inset (`b`).
    Bottom,
    /// Spread lines to fill the body (`just`).
    Justified,
    /// Spread words to fill the body (`dist`).
    Distributed,
}

impl Anchor {
    /// Return the exact DrawingML token.
    #[must_use]
    pub const fn token(self) -> &'static str {
        match self {
            Self::Top => "t",
            Self::Center => "ctr",
            Self::Bottom => "b",
            Self::Justified => "just",
            Self::Distributed => "dist",
        }
    }
}

impl FromStr for Anchor {
    type Err = ParseError;

    fn from_str(value: &str) -> Result<Self, Self::Err> {
        match value {
            "t" => Ok(Self::Top),
            "ctr" => Ok(Self::Center),
            "b" => Ok(Self::Bottom),
            "just" => Ok(Self::Justified),
            "dist" => Ok(Self::Distributed),
            _ => Err(ParseError::InvalidToken("text anchor")),
        }
    }
}

impl fmt::Display for Anchor {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.token())
    }
}

/// Direction of text within a shape (`ST_TextVerticalType`).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum Direction {
    /// Horizontal text (`horz`).
    #[default]
    Horizontal,
    /// Lines rotated 90 degrees (`vert`).
    Vertical,
    /// Lines rotated 270 degrees (`vert270`).
    Vertical270,
    /// Upright, stacked WordArt letters (`wordArtVert`).
    WordArtVertical,
    /// East Asian vertical text (`eaVert`).
    EastAsianVertical,
    /// Mongolian vertical text (`mongolianVert`).
    MongolianVertical,
    /// Right-to-left WordArt vertical text (`wordArtVertRtl`).
    WordArtVerticalRtl,
}

impl Direction {
    /// Return the exact DrawingML token.
    #[must_use]
    pub const fn token(self) -> &'static str {
        match self {
            Self::Horizontal => "horz",
            Self::Vertical => "vert",
            Self::Vertical270 => "vert270",
            Self::WordArtVertical => "wordArtVert",
            Self::EastAsianVertical => "eaVert",
            Self::MongolianVertical => "mongolianVert",
            Self::WordArtVerticalRtl => "wordArtVertRtl",
        }
    }
}

impl FromStr for Direction {
    type Err = ParseError;

    fn from_str(value: &str) -> Result<Self, Self::Err> {
        match value {
            "horz" => Ok(Self::Horizontal),
            "vert" => Ok(Self::Vertical),
            "vert270" => Ok(Self::Vertical270),
            "wordArtVert" => Ok(Self::WordArtVertical),
            "eaVert" => Ok(Self::EastAsianVertical),
            "mongolianVert" => Ok(Self::MongolianVertical),
            "wordArtVertRtl" => Ok(Self::WordArtVerticalRtl),
            _ => Err(ParseError::InvalidToken("text direction")),
        }
    }
}

impl fmt::Display for Direction {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.token())
    }
}

/// Whether text wraps inside the shape extents (`ST_TextWrappingType`).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum Wrap {
    /// Wrap inside the bounding rectangle (`square`).
    #[default]
    Square,
    /// Do not wrap (`none`).
    None,
}

impl Wrap {
    /// Return the exact DrawingML token.
    #[must_use]
    pub const fn token(self) -> &'static str {
        match self {
            Self::Square => "square",
            Self::None => "none",
        }
    }
}

impl FromStr for Wrap {
    type Err = ParseError;

    fn from_str(value: &str) -> Result<Self, Self::Err> {
        match value {
            "square" => Ok(Self::Square),
            "none" => Ok(Self::None),
            _ => Err(ParseError::InvalidToken("text wrap")),
        }
    }
}

impl fmt::Display for Wrap {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.token())
    }
}

/// Autofit behavior selected by a text-body child element.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum Autofit {
    /// `a:noAutofit`.
    #[default]
    None,
    /// `a:spAutoFit`.
    Shape,
    /// `a:normAutofit`.
    Normal,
}

/// Lossless underline style shared by DrawingML and WordprocessingML.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum Underline {
    #[default]
    None,
    Words,
    Single,
    Double,
    Heavy,
    Dotted,
    DottedHeavy,
    Dash,
    DashHeavy,
    DashLong,
    DashLongHeavy,
    DotDash,
    DotDashHeavy,
    DotDotDash,
    DotDotDashHeavy,
    Wavy,
    WavyHeavy,
    WavyDouble,
}

impl Underline {
    /// Parse an `a:rPr@u` DrawingML token.
    pub fn from_dml(value: &str) -> Result<Self, ParseError> {
        match value {
            "none" => Ok(Self::None),
            "words" => Ok(Self::Words),
            "sng" => Ok(Self::Single),
            "dbl" => Ok(Self::Double),
            "heavy" => Ok(Self::Heavy),
            "dotted" => Ok(Self::Dotted),
            "dottedHeavy" => Ok(Self::DottedHeavy),
            "dash" => Ok(Self::Dash),
            "dashHeavy" => Ok(Self::DashHeavy),
            "dashLong" => Ok(Self::DashLong),
            "dashLongHeavy" => Ok(Self::DashLongHeavy),
            "dotDash" => Ok(Self::DotDash),
            "dotDashHeavy" => Ok(Self::DotDashHeavy),
            "dotDotDash" => Ok(Self::DotDotDash),
            "dotDotDashHeavy" => Ok(Self::DotDotDashHeavy),
            "wavy" => Ok(Self::Wavy),
            "wavyHeavy" => Ok(Self::WavyHeavy),
            "wavyDbl" => Ok(Self::WavyDouble),
            _ => Err(ParseError::InvalidToken("DrawingML underline")),
        }
    }

    /// Return the exact `a:rPr@u` DrawingML token.
    #[must_use]
    pub const fn dml(self) -> &'static str {
        match self {
            Self::None => "none",
            Self::Words => "words",
            Self::Single => "sng",
            Self::Double => "dbl",
            Self::Heavy => "heavy",
            Self::Dotted => "dotted",
            Self::DottedHeavy => "dottedHeavy",
            Self::Dash => "dash",
            Self::DashHeavy => "dashHeavy",
            Self::DashLong => "dashLong",
            Self::DashLongHeavy => "dashLongHeavy",
            Self::DotDash => "dotDash",
            Self::DotDashHeavy => "dotDashHeavy",
            Self::DotDotDash => "dotDotDash",
            Self::DotDotDashHeavy => "dotDotDashHeavy",
            Self::Wavy => "wavy",
            Self::WavyHeavy => "wavyHeavy",
            Self::WavyDouble => "wavyDbl",
        }
    }

    /// Parse a `w:u@w:val` WordprocessingML token.
    pub fn from_wml(value: &str) -> Result<Self, ParseError> {
        match value {
            "none" => Ok(Self::None),
            "words" => Ok(Self::Words),
            "single" => Ok(Self::Single),
            "double" => Ok(Self::Double),
            "thick" => Ok(Self::Heavy),
            "dotted" => Ok(Self::Dotted),
            "dottedHeavy" => Ok(Self::DottedHeavy),
            "dash" => Ok(Self::Dash),
            "dashedHeavy" => Ok(Self::DashHeavy),
            "dashLong" => Ok(Self::DashLong),
            "dashLongHeavy" => Ok(Self::DashLongHeavy),
            "dotDash" => Ok(Self::DotDash),
            "dashDotHeavy" => Ok(Self::DotDashHeavy),
            "dotDotDash" => Ok(Self::DotDotDash),
            "dashDotDotHeavy" => Ok(Self::DotDotDashHeavy),
            "wave" => Ok(Self::Wavy),
            "wavyHeavy" => Ok(Self::WavyHeavy),
            "wavyDouble" => Ok(Self::WavyDouble),
            _ => Err(ParseError::InvalidToken("WordprocessingML underline")),
        }
    }

    /// Return the exact `w:u@w:val` WordprocessingML token.
    #[must_use]
    pub const fn wml(self) -> &'static str {
        match self {
            Self::None => "none",
            Self::Words => "words",
            Self::Single => "single",
            Self::Double => "double",
            Self::Heavy => "thick",
            Self::Dotted => "dotted",
            Self::DottedHeavy => "dottedHeavy",
            Self::Dash => "dash",
            Self::DashHeavy => "dashedHeavy",
            Self::DashLong => "dashLong",
            Self::DashLongHeavy => "dashLongHeavy",
            Self::DotDash => "dotDash",
            Self::DotDashHeavy => "dashDotHeavy",
            Self::DotDotDash => "dotDotDash",
            Self::DotDotDashHeavy => "dashDotDotHeavy",
            Self::Wavy => "wave",
            Self::WavyHeavy => "wavyHeavy",
            Self::WavyDouble => "wavyDouble",
        }
    }
}

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
            .trim_matches(xsd_whitespace)
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
            .trim_matches(xsd_whitespace)
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

/// Parse an XML Schema boolean (`true`, `false`, `1`, or `0`).
pub fn parse_bool(value: &str) -> Result<bool, ParseError> {
    match value.trim_matches(xsd_whitespace) {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(ParseError::InvalidBool),
    }
}

/// Parse WordprocessingML `ST_OnOff`, including transitional `on`/`off`.
pub fn parse_on_off(value: &str) -> Result<bool, ParseError> {
    match value {
        "on" => Ok(true),
        "off" => Ok(false),
        _ => parse_bool(value),
    }
}

const fn xsd_whitespace(character: char) -> bool {
    matches!(character, ' ' | '\t' | '\n' | '\r')
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::mem::size_of;

    #[test]
    fn closed_domains_round_trip_and_reject_unknown_tokens() {
        for value in [
            Anchor::Top,
            Anchor::Center,
            Anchor::Bottom,
            Anchor::Justified,
            Anchor::Distributed,
        ] {
            assert_eq!(value.token().parse::<Anchor>().unwrap(), value);
        }
        for value in [
            Direction::Horizontal,
            Direction::Vertical,
            Direction::Vertical270,
            Direction::WordArtVertical,
            Direction::EastAsianVertical,
            Direction::MongolianVertical,
            Direction::WordArtVerticalRtl,
        ] {
            assert_eq!(value.token().parse::<Direction>().unwrap(), value);
        }
        for value in [Wrap::Square, Wrap::None] {
            assert_eq!(value.token().parse::<Wrap>().unwrap(), value);
        }
        assert!("middle".parse::<Anchor>().is_err());
        assert!("diagonal".parse::<Direction>().is_err());
        assert!("tight".parse::<Wrap>().is_err());
    }

    #[test]
    fn underline_codecs_are_lossless_in_both_dialects() {
        let values = [
            Underline::None,
            Underline::Words,
            Underline::Single,
            Underline::Double,
            Underline::Heavy,
            Underline::Dotted,
            Underline::DottedHeavy,
            Underline::Dash,
            Underline::DashHeavy,
            Underline::DashLong,
            Underline::DashLongHeavy,
            Underline::DotDash,
            Underline::DotDashHeavy,
            Underline::DotDotDash,
            Underline::DotDotDashHeavy,
            Underline::Wavy,
            Underline::WavyHeavy,
            Underline::WavyDouble,
        ];
        for value in values {
            assert_eq!(Underline::from_dml(value.dml()).unwrap(), value);
            assert_eq!(Underline::from_wml(value.wml()).unwrap(), value);
        }
        assert!(Underline::from_dml("single").is_err());
        assert!(Underline::from_wml("sng").is_err());
    }

    #[test]
    fn bounded_values_reject_invalid_authoring_and_xml() {
        assert_eq!(Columns::new(16).unwrap().get(), 16);
        assert!(Columns::new(0).is_err());
        assert!("17".parse::<Columns>().is_err());
        assert_eq!(TextSize::new(100).unwrap().get(), 100);
        assert_eq!(TextSize::new(400_000).unwrap().get(), 400_000);
        assert!(TextSize::new(99).is_err());
        assert!("400001".parse::<TextSize>().is_err());
        assert_eq!(
            Coordinate32::try_from(i64::from(i32::MIN))
                .unwrap()
                .as_emu(),
            Some(i32::MIN)
        );
        assert!(Coordinate32::try_from(i64::from(i32::MAX) + 1).is_err());
        assert_eq!(
            "1.25cm".parse::<Coordinate32>().unwrap().to_string(),
            "1.25cm"
        );
    }

    #[test]
    fn booleans_are_dialect_exact() {
        for (token, expected) in [("1", true), ("true", true), ("0", false), ("false", false)] {
            assert_eq!(parse_bool(token).unwrap(), expected);
        }
        assert!(parse_bool("on").is_err());
        assert!(parse_on_off("on").unwrap());
        assert!(!parse_on_off("off").unwrap());
        assert!(parse_on_off("yes").is_err());
    }

    #[test]
    fn common_values_remain_cache_friendly() {
        assert_eq!(size_of::<Anchor>(), 1);
        assert_eq!(size_of::<Direction>(), 1);
        assert_eq!(size_of::<Wrap>(), 1);
        assert_eq!(size_of::<Autofit>(), 1);
        assert_eq!(size_of::<Underline>(), 1);
        assert_eq!(size_of::<Columns>(), 1);
        assert_eq!(size_of::<TextSize>(), 4);
    }
}
