//! Typed SpreadsheetML cell alignment values.

use std::error::Error;
use std::fmt::{Display, Formatter};
use std::str::FromStr;

/// Horizontal placement of cell content.
///
/// This is the complete `ST_HorizontalAlignment` value set from
/// SpreadsheetML. Keeping the wire tokens behind an enum prevents authored
/// workbooks from containing misspelled or unsupported values.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum Horizontal {
    General,
    Left,
    Center,
    Right,
    Fill,
    Justify,
    CenterContinuous,
    Distributed,
}

impl Horizontal {
    /// Return the SpreadsheetML token.
    #[inline]
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::General => "general",
            Self::Left => "left",
            Self::Center => "center",
            Self::Right => "right",
            Self::Fill => "fill",
            Self::Justify => "justify",
            Self::CenterContinuous => "centerContinuous",
            Self::Distributed => "distributed",
        }
    }

    fn parse(value: &str) -> Option<Self> {
        match value {
            "general" => Some(Self::General),
            "left" => Some(Self::Left),
            "center" => Some(Self::Center),
            "right" => Some(Self::Right),
            "fill" => Some(Self::Fill),
            "justify" => Some(Self::Justify),
            "centerContinuous" => Some(Self::CenterContinuous),
            "distributed" => Some(Self::Distributed),
            _ => None,
        }
    }
}

impl Display for Horizontal {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.as_str())
    }
}

impl FromStr for Horizontal {
    type Err = ParseError;

    fn from_str(value: &str) -> Result<Self, Self::Err> {
        Self::parse(value).ok_or(ParseError("horizontal alignment"))
    }
}

/// Vertical placement of cell content.
///
/// This is the complete `ST_VerticalAlignment` value set from
/// SpreadsheetML.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum Vertical {
    Top,
    Center,
    Bottom,
    Justify,
    Distributed,
}

impl Vertical {
    /// Return the SpreadsheetML token.
    #[inline]
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Top => "top",
            Self::Center => "center",
            Self::Bottom => "bottom",
            Self::Justify => "justify",
            Self::Distributed => "distributed",
        }
    }

    fn parse(value: &str) -> Option<Self> {
        match value {
            "top" => Some(Self::Top),
            "center" => Some(Self::Center),
            "bottom" => Some(Self::Bottom),
            "justify" => Some(Self::Justify),
            "distributed" => Some(Self::Distributed),
            _ => None,
        }
    }
}

impl Display for Vertical {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.as_str())
    }
}

impl FromStr for Vertical {
    type Err = ParseError;

    fn from_str(value: &str) -> Result<Self, Self::Err> {
        Self::parse(value).ok_or(ParseError("vertical alignment"))
    }
}

/// Error returned when an alignment token is outside its fixed domain.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ParseError(&'static str);

impl Display for ParseError {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        write!(formatter, "invalid SpreadsheetML {}", self.0)
    }
}

impl Error for ParseError {}

/// A SpreadsheetML text rotation.
///
/// Values `0..=180` are rotation degrees, `254` is context-dependent rotation,
/// and `255` means stacked vertical text. Context-dependent rotation is an
/// Office Transitional extension and is rejected by the Strict writer. The
/// private representation makes every value constructible through this API
/// valid by design.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(transparent)]
pub struct Rotation(u8);

impl Rotation {
    const CONTEXTUAL: u8 = 254;
    const STACKED: u8 = 255;

    /// Construct a degree rotation in the inclusive `0..=180` range.
    pub const fn degrees(value: u8) -> Result<Self, InvalidRotation> {
        if value <= 180 {
            Ok(Self(value))
        } else {
            Err(InvalidRotation(value as u32))
        }
    }

    /// Construct Office's context-dependent rotation, represented by `254`.
    #[inline]
    #[must_use]
    pub const fn contextual() -> Self {
        Self(Self::CONTEXTUAL)
    }

    /// Construct stacked vertical text, represented by `255` on the wire.
    #[inline]
    #[must_use]
    pub const fn stacked() -> Self {
        Self(Self::STACKED)
    }

    /// Return the SpreadsheetML numeric value.
    #[inline]
    #[must_use]
    pub const fn get(self) -> u8 {
        self.0
    }

    /// Whether this is Office's context-dependent rotation value.
    #[inline]
    #[must_use]
    pub const fn is_contextual(self) -> bool {
        self.0 == Self::CONTEXTUAL
    }

    /// Whether this value requests stacked vertical text.
    #[inline]
    #[must_use]
    pub const fn is_stacked(self) -> bool {
        self.0 == Self::STACKED
    }
}

impl TryFrom<u32> for Rotation {
    type Error = InvalidRotation;

    fn try_from(value: u32) -> Result<Self, Self::Error> {
        match value {
            0..=180 => Ok(Self(value as u8)),
            254 => Ok(Self::contextual()),
            255 => Ok(Self::stacked()),
            _ => Err(InvalidRotation(value)),
        }
    }
}

/// Error returned when a number is not a SpreadsheetML text rotation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct InvalidRotation(u32);

impl InvalidRotation {
    /// Return the rejected numeric value.
    #[inline]
    #[must_use]
    pub const fn value(self) -> u32 {
        self.0
    }
}

impl Display for InvalidRotation {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        write!(
            formatter,
            "text rotation {} is outside 0..=180 and is neither contextual value 254 nor stacked value 255",
            self.0
        )
    }
}

impl Error for InvalidRotation {}

/// Direction used to lay out cell text.
///
/// SpreadsheetML's schema exposes an unsigned integer, while Microsoft Excel
/// restricts interoperable values to these three states.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum Reading {
    Context = 0,
    LeftToRight = 1,
    RightToLeft = 2,
}

impl Reading {
    /// Return the SpreadsheetML numeric value.
    #[inline]
    #[must_use]
    pub const fn get(self) -> u8 {
        self as u8
    }
}

impl TryFrom<u32> for Reading {
    type Error = InvalidReading;

    fn try_from(value: u32) -> Result<Self, Self::Error> {
        match value {
            0 => Ok(Self::Context),
            1 => Ok(Self::LeftToRight),
            2 => Ok(Self::RightToLeft),
            _ => Err(InvalidReading(value)),
        }
    }
}

/// Error returned when a number is not an Excel-compatible reading order.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct InvalidReading(u32);

impl InvalidReading {
    /// Return the rejected numeric value.
    #[inline]
    #[must_use]
    pub const fn value(self) -> u32 {
        self.0
    }
}

impl Display for InvalidReading {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        write!(formatter, "reading order {} is not 0, 1, or 2", self.0)
    }
}

impl Error for InvalidReading {}

/// Excel-compatible cell indentation in the inclusive `0..=255` range.
///
/// SpreadsheetML uses an unsigned integer on the wire, but Microsoft Excel's
/// documented interoperability limit is 255. Storing a `u8` makes that bound
/// a type invariant without runtime overhead.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(transparent)]
pub struct Indent(u8);

impl Indent {
    /// Construct an indentation level. Every `u8` is valid.
    #[inline]
    #[must_use]
    pub const fn new(value: u8) -> Self {
        Self(value)
    }

    /// Return the SpreadsheetML numeric value.
    #[inline]
    #[must_use]
    pub const fn get(self) -> u8 {
        self.0
    }
}

impl TryFrom<u32> for Indent {
    type Error = InvalidIndent;

    fn try_from(value: u32) -> Result<Self, Self::Error> {
        u8::try_from(value)
            .map(Self)
            .map_err(|_| InvalidIndent(value))
    }
}

/// Error returned when indentation exceeds Excel's interoperable range.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct InvalidIndent(u32);

impl InvalidIndent {
    /// Return the rejected numeric value.
    #[inline]
    #[must_use]
    pub const fn value(self) -> u32 {
        self.0
    }
}

impl Display for InvalidIndent {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        write!(
            formatter,
            "alignment indent {} exceeds Excel's maximum of 255",
            self.0
        )
    }
}

impl Error for InvalidIndent {}

/// Alignment information for cell content.
///
/// Public fields make struct-update syntax ergonomic, while all fixed wire
/// domains use enums or checked values.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub struct Alignment {
    /// Horizontal placement.
    pub horizontal: Option<Horizontal>,
    /// Vertical placement.
    pub vertical: Option<Vertical>,
    /// Text rotation or stacked vertical text.
    pub text_rotation: Option<Rotation>,
    /// Wrap text within the cell.
    pub wrap_text: bool,
    /// Excel-compatible indent level.
    pub indent: Option<Indent>,
    /// Signed relative indent.
    pub relative_indent: Option<i32>,
    /// Justify the final line of text.
    pub justify_last_line: bool,
    /// Reduce the font size to fit the cell.
    pub shrink_to_fit: bool,
    /// Text reading direction.
    pub reading_order: Option<Reading>,
}

impl Alignment {
    /// Create an alignment with no explicitly authored settings.
    #[inline]
    #[must_use]
    pub const fn new() -> Self {
        Self {
            horizontal: None,
            vertical: None,
            text_rotation: None,
            wrap_text: false,
            indent: None,
            relative_indent: None,
            justify_last_line: false,
            shrink_to_fit: false,
            reading_order: None,
        }
    }

    /// Create an alignment with horizontal placement.
    #[inline]
    #[must_use]
    pub const fn horizontal(value: Horizontal) -> Self {
        Self {
            horizontal: Some(value),
            ..Self::new()
        }
    }

    /// Create an alignment with vertical placement.
    #[inline]
    #[must_use]
    pub const fn vertical(value: Vertical) -> Self {
        Self {
            vertical: Some(value),
            ..Self::new()
        }
    }

    /// Create an alignment with both placement directions.
    #[inline]
    #[must_use]
    pub const fn both(horizontal: Horizontal, vertical: Vertical) -> Self {
        Self {
            horizontal: Some(horizontal),
            vertical: Some(vertical),
            ..Self::new()
        }
    }

    /// Whether any alignment setting is explicitly present.
    #[inline]
    #[must_use]
    pub const fn has_settings(&self) -> bool {
        self.horizontal.is_some()
            || self.vertical.is_some()
            || self.text_rotation.is_some()
            || self.wrap_text
            || self.indent.is_some()
            || self.relative_indent.is_some()
            || self.justify_last_line
            || self.shrink_to_fit
            || self.reading_order.is_some()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn fixed_tokens_round_trip_without_string_fallbacks() {
        let horizontal = [
            Horizontal::General,
            Horizontal::Left,
            Horizontal::Center,
            Horizontal::Right,
            Horizontal::Fill,
            Horizontal::Justify,
            Horizontal::CenterContinuous,
            Horizontal::Distributed,
        ];
        for value in horizontal {
            assert_eq!(value.as_str().parse::<Horizontal>(), Ok(value));
            assert_eq!(value.to_string(), value.as_str());
        }

        let vertical = [
            Vertical::Top,
            Vertical::Center,
            Vertical::Bottom,
            Vertical::Justify,
            Vertical::Distributed,
        ];
        for value in vertical {
            assert_eq!(value.as_str().parse::<Vertical>(), Ok(value));
            assert_eq!(value.to_string(), value.as_str());
        }
        assert!("middle".parse::<Horizontal>().is_err());
        assert!("middle".parse::<Vertical>().is_err());
    }

    #[test]
    fn rotation_is_valid_by_construction() {
        assert_eq!(Rotation::degrees(0).unwrap().get(), 0);
        assert_eq!(Rotation::degrees(180).unwrap().get(), 180);
        assert!(Rotation::degrees(181).is_err());
        assert!(Rotation::contextual().is_contextual());
        assert_eq!(Rotation::try_from(254), Ok(Rotation::contextual()));
        assert!(Rotation::stacked().is_stacked());
        assert_eq!(Rotation::try_from(255), Ok(Rotation::stacked()));
        assert!(Rotation::try_from(u32::MAX).is_err());
        assert_eq!(Indent::new(255).get(), 255);
        assert_eq!(Indent::try_from(255), Ok(Indent::new(255)));
        assert!(Indent::try_from(256).is_err());
    }

    #[test]
    fn struct_update_syntax_keeps_alignment_concise() {
        let alignment = Alignment {
            horizontal: Some(Horizontal::Center),
            text_rotation: Some(Rotation::degrees(45).unwrap()),
            wrap_text: true,
            reading_order: Some(Reading::RightToLeft),
            ..Alignment::new()
        };
        assert!(alignment.has_settings());
        assert_eq!(alignment.vertical, None);
    }
}
