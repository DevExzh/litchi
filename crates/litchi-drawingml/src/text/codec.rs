//! Exact lexical codecs shared by `DrawingML` text primitives.

use std::fmt;
use std::str::FromStr;

use super::model::{
    Alignment, Anchor, AutoNumber, Caps, Direction, FontAlignment, Strike, TabAlignment, Underline,
    Wrap,
};

/// Failure to parse or construct a `DrawingML` text primitive.
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
    Coordinate(crate::coordinate::ParseError),
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

impl From<crate::coordinate::ParseError> for ParseError {
    fn from(error: crate::coordinate::ParseError) -> Self {
        Self::Coordinate(error)
    }
}

impl Anchor {
    /// Return the exact `DrawingML` token.
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
    /// Return the exact `DrawingML` token.
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
    /// Return the exact `DrawingML` token.
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
    /// Parse an `a:rPr@u` `DrawingML` token.
    /// # Errors
    ///
    /// Returns an error when input violates DrawingML constraints, exceeds a configured
    /// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
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

    /// Return the exact `a:rPr@u` `DrawingML` token.
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

    /// Parse a `w:u@w:val` `WordprocessingML` token.
    /// # Errors
    ///
    /// Returns an error when input violates DrawingML constraints, exceeds a configured
    /// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
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

    /// Return the exact `w:u@w:val` `WordprocessingML` token.
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
/// # Errors
///
/// Returns an error when input violates DrawingML constraints, exceeds a configured
/// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
pub fn parse_bool(value: &str) -> Result<bool, ParseError> {
    match value.trim_matches(xsd_whitespace) {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(ParseError::InvalidBool),
    }
}

/// Parse `WordprocessingML` `ST_OnOff`, including transitional `on`/`off`.
/// # Errors
///
/// Returns an error when input violates DrawingML constraints, exceeds a configured
/// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
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

/// Implement `ALL`, `token`, `FromStr` and `Display` for a closed `DrawingML`
/// token domain. Parsing is exact: callers apply the XML Schema `token`
/// whitespace facet first when the attribute's type requires it.
macro_rules! closed_tokens {
    ($name:ident, $domain:literal, { $($variant:ident => $token:literal),+ $(,)? }) => {
        impl $name {
            /// Every value, in schema order.
            pub const ALL: &'static [Self] = &[$(Self::$variant),+];

            /// Return the exact `DrawingML` token.
            #[must_use]
            pub const fn token(self) -> &'static str {
                match self {
                    $(Self::$variant => $token),+
                }
            }
        }

        impl FromStr for $name {
            type Err = ParseError;

            fn from_str(value: &str) -> Result<Self, Self::Err> {
                match value {
                    $($token => Ok(Self::$variant),)+
                    _ => Err(ParseError::InvalidToken($domain)),
                }
            }
        }

        impl fmt::Display for $name {
            fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
                formatter.write_str(self.token())
            }
        }
    };
}

closed_tokens!(Strike, "text strike", {
    None => "noStrike",
    Single => "sngStrike",
    Double => "dblStrike",
});

closed_tokens!(Caps, "text caps", {
    None => "none",
    Small => "small",
    All => "all",
});

closed_tokens!(Alignment, "paragraph alignment", {
    Left => "l",
    Center => "ctr",
    Right => "r",
    Justified => "just",
    JustifiedLow => "justLow",
    Distributed => "dist",
    ThaiDistributed => "thaiDist",
});

closed_tokens!(FontAlignment, "font alignment", {
    Auto => "auto",
    Top => "t",
    Center => "ctr",
    Baseline => "base",
    Bottom => "b",
});

closed_tokens!(TabAlignment, "tab alignment", {
    Left => "l",
    Center => "ctr",
    Right => "r",
    Decimal => "dec",
});

closed_tokens!(AutoNumber, "auto-number scheme", {
    AlphaLowerParenBoth => "alphaLcParenBoth",
    AlphaUpperParenBoth => "alphaUcParenBoth",
    AlphaLowerParenRight => "alphaLcParenR",
    AlphaUpperParenRight => "alphaUcParenR",
    AlphaLowerPeriod => "alphaLcPeriod",
    AlphaUpperPeriod => "alphaUcPeriod",
    ArabicParenBoth => "arabicParenBoth",
    ArabicParenRight => "arabicParenR",
    ArabicPeriod => "arabicPeriod",
    ArabicPlain => "arabicPlain",
    RomanLowerParenBoth => "romanLcParenBoth",
    RomanUpperParenBoth => "romanUcParenBoth",
    RomanLowerParenRight => "romanLcParenR",
    RomanUpperParenRight => "romanUcParenR",
    RomanLowerPeriod => "romanLcPeriod",
    RomanUpperPeriod => "romanUcPeriod",
    CircleNumberDoubleBytePlain => "circleNumDbPlain",
    CircleNumberWingdingsBlackPlain => "circleNumWdBlackPlain",
    CircleNumberWingdingsWhitePlain => "circleNumWdWhitePlain",
    ArabicDoubleBytePeriod => "arabicDbPeriod",
    ArabicDoubleBytePlain => "arabicDbPlain",
    EastAsianSimplifiedChinesePeriod => "ea1ChsPeriod",
    EastAsianSimplifiedChinesePlain => "ea1ChsPlain",
    EastAsianTraditionalChinesePeriod => "ea1ChtPeriod",
    EastAsianTraditionalChinesePlain => "ea1ChtPlain",
    EastAsianJapaneseChineseDoubleBytePeriod => "ea1JpnChsDbPeriod",
    EastAsianJapaneseKoreanPlain => "ea1JpnKorPlain",
    EastAsianJapaneseKoreanPeriod => "ea1JpnKorPeriod",
    Arabic1Minus => "arabic1Minus",
    Arabic2Minus => "arabic2Minus",
    Hebrew2Minus => "hebrew2Minus",
    ThaiAlphaPeriod => "thaiAlphaPeriod",
    ThaiAlphaParenRight => "thaiAlphaParenR",
    ThaiAlphaParenBoth => "thaiAlphaParenBoth",
    ThaiNumberPeriod => "thaiNumPeriod",
    ThaiNumberParenRight => "thaiNumParenR",
    ThaiNumberParenBoth => "thaiNumParenBoth",
    HindiAlphaPeriod => "hindiAlphaPeriod",
    HindiNumberPeriod => "hindiNumPeriod",
    HindiNumberParenRight => "hindiNumParenR",
    HindiAlpha1Period => "hindiAlpha1Period",
});
