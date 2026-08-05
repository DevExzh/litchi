//! Exact lexical codecs shared by DrawingML text primitives.

use std::fmt;
use std::str::FromStr;

use super::model::{Anchor, Direction, Underline, Wrap};

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

pub(super) const fn xsd_whitespace(character: char) -> bool {
    matches!(character, ' ' | '\t' | '\n' | '\r')
}
