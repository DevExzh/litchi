//! Font information and definitions.

use std::fmt;
use std::str::FromStr;

use super::border::Tint;
use crate::color::Rgb;

/// Error returned when a font property is not an exact SpreadsheetML token.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ParseError {
    kind: &'static str,
    value: Box<str>,
}

impl ParseError {
    fn new(kind: &'static str, value: &str) -> Self {
        Self {
            kind,
            value: value.into(),
        }
    }

    /// Closed font-property domain that rejected the token.
    #[must_use]
    pub const fn kind(&self) -> &'static str {
        self.kind
    }

    /// Rejected token.
    #[must_use]
    pub fn value(&self) -> &str {
        &self.value
    }
}

impl fmt::Display for ParseError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            formatter,
            "invalid SpreadsheetML font {} token '{}'",
            self.kind, self.value
        )
    }
}

impl std::error::Error for ParseError {}

/// The semantic base value of a SpreadsheetML font color.
///
/// Unlike the wire representation, this type cannot contain competing color
/// bases. The parser keeps the original lexical start tag separately on
/// [`FontColor`] so callers can retain producer-specific attributes without
/// exposing an unvalidated string authoring path.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum FontColorKind {
    /// No explicit base; the consumer default is used.
    Default,
    /// A checked four-byte ARGB value.
    Rgb(Rgb),
    /// A theme color index.
    Theme(u32),
    /// An indexed palette color.
    Indexed(u32),
    /// An explicit automatic-color flag.
    Auto(bool),
}

/// A validated SpreadsheetML font color.
///
/// Construction uses typed base values and checked [`Tint`] values, making
/// conflicting `rgb`/`theme`/`indexed`/`auto` states unrepresentable. Parsed
/// colors additionally retain their original compact lexical start tag for
/// diagnostics and lossless source inspection; the stylesheet writer emits a
/// canonical compact color element.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct FontColor {
    kind: FontColorKind,
    tint: Option<Tint>,
    source_lexical: Option<Box<str>>,
}

impl FontColor {
    /// Create a color using the consumer default base.
    #[must_use]
    pub const fn default_base() -> Self {
        Self {
            kind: FontColorKind::Default,
            tint: None,
            source_lexical: None,
        }
    }

    /// Create a checked ARGB font color.
    #[must_use]
    pub const fn rgb(value: Rgb) -> Self {
        Self {
            kind: FontColorKind::Rgb(value),
            tint: None,
            source_lexical: None,
        }
    }

    /// Create a theme-indexed font color.
    #[must_use]
    pub const fn theme(index: u32) -> Self {
        Self {
            kind: FontColorKind::Theme(index),
            tint: None,
            source_lexical: None,
        }
    }

    /// Create a palette-indexed font color.
    #[must_use]
    pub const fn indexed(index: u32) -> Self {
        Self {
            kind: FontColorKind::Indexed(index),
            tint: None,
            source_lexical: None,
        }
    }

    /// Create an explicit automatic font color.
    #[must_use]
    pub const fn automatic() -> Self {
        Self {
            kind: FontColorKind::Auto(true),
            tint: None,
            source_lexical: None,
        }
    }

    /// Attach a checked tint to this color.
    #[must_use]
    pub fn with_tint(mut self, tint: Tint) -> Self {
        self.tint = Some(tint);
        self.source_lexical = None;
        self
    }

    /// Semantic base value.
    #[must_use]
    pub const fn kind(&self) -> FontColorKind {
        self.kind
    }

    /// Optional checked tint.
    #[must_use]
    pub const fn tint(&self) -> Option<Tint> {
        self.tint
    }

    /// Original lexical start tag, when this value came from a stylesheet.
    ///
    /// This is intentionally inspection-only: writers use the typed value and
    /// produce compact XML rather than replaying producer formatting.
    #[must_use]
    pub fn source_lexical(&self) -> Option<&str> {
        self.source_lexical.as_deref()
    }

    pub(crate) fn parsed(
        kind: FontColorKind,
        tint: Option<Tint>,
        source_lexical: Box<str>,
    ) -> Self {
        Self {
            kind,
            tint,
            source_lexical: Some(source_lexical),
        }
    }
}

/// Font underline style (`ST_UnderlineValues`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum Underline {
    Single,
    Double,
    SingleAccounting,
    DoubleAccounting,
    None,
}

impl Underline {
    /// Return the exact SpreadsheetML token.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Single => "single",
            Self::Double => "double",
            Self::SingleAccounting => "singleAccounting",
            Self::DoubleAccounting => "doubleAccounting",
            Self::None => "none",
        }
    }

    /// Whether this value visibly underlines the font.
    #[must_use]
    pub const fn is_enabled(self) -> bool {
        !matches!(self, Self::None)
    }
}

impl FromStr for Underline {
    type Err = ParseError;

    fn from_str(value: &str) -> Result<Self, Self::Err> {
        match value {
            "single" => Ok(Self::Single),
            "double" => Ok(Self::Double),
            "singleAccounting" => Ok(Self::SingleAccounting),
            "doubleAccounting" => Ok(Self::DoubleAccounting),
            "none" => Ok(Self::None),
            _ => Err(ParseError::new("underline", value)),
        }
    }
}

impl fmt::Display for Underline {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}

/// Theme font scheme (`ST_FontScheme`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum Scheme {
    None,
    Major,
    Minor,
}

impl Scheme {
    /// Return the exact SpreadsheetML token.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::None => "none",
            Self::Major => "major",
            Self::Minor => "minor",
        }
    }
}

impl FromStr for Scheme {
    type Err = ParseError;

    fn from_str(value: &str) -> Result<Self, Self::Err> {
        match value {
            "none" => Ok(Self::None),
            "major" => Ok(Self::Major),
            "minor" => Ok(Self::Minor),
            _ => Err(ParseError::new("scheme", value)),
        }
    }
}

impl fmt::Display for Scheme {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}

/// Baseline placement for cell font text (`ST_VerticalAlignRun`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum Script {
    Baseline,
    Superscript,
    Subscript,
}

impl Script {
    /// Return the exact SpreadsheetML token.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Baseline => "baseline",
            Self::Superscript => "superscript",
            Self::Subscript => "subscript",
        }
    }

    /// Whether this value moves text away from its normal baseline.
    #[must_use]
    pub const fn is_shifted(self) -> bool {
        !matches!(self, Self::Baseline)
    }
}

impl FromStr for Script {
    type Err = ParseError;

    fn from_str(value: &str) -> Result<Self, Self::Err> {
        match value {
            "baseline" => Ok(Self::Baseline),
            "superscript" => Ok(Self::Superscript),
            "subscript" => Ok(Self::Subscript),
            _ => Err(ParseError::new("script", value)),
        }
    }
}

impl fmt::Display for Script {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}

/// Font information.
///
/// Defines the visual appearance of text in cells including
/// typeface, size, color, and text decoration.
#[derive(Debug, Clone, Default)]
pub struct Font {
    /// Font name/family (e.g., "Calibri", "Arial")
    pub name: Option<String>,
    /// Font size in points
    pub size: Option<f64>,
    /// Bold flag
    pub bold: bool,
    /// Italic flag
    pub italic: bool,
    /// Underline style, including an explicitly authored `none`.
    pub underline: Option<Underline>,
    /// Strike-through flag
    pub strike: bool,
    /// Typed font color.
    pub color: Option<FontColor>,
    /// Font charset
    pub charset: Option<u32>,
    /// Font family (1=Roman, 2=Swiss, 3=Modern, 4=Script, 5=Decorative)
    pub family: Option<u32>,
    /// Theme font scheme.
    pub scheme: Option<Scheme>,
    /// Baseline placement.
    pub script: Option<Script>,
}

impl Font {
    /// Create a new default font.
    #[inline]
    pub fn new() -> Self {
        Self::default()
    }

    /// Check if the font has any special formatting.
    #[inline]
    pub fn has_formatting(&self) -> bool {
        self.bold
            || self.italic
            || self.strike
            || self.underline.is_some_and(Underline::is_enabled)
            || self.script.is_some_and(Script::is_shifted)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::mem::size_of;

    #[test]
    fn test_font_default() {
        let font = Font::default();
        assert!(font.name.is_none());
        assert!(font.size.is_none());
        assert!(!font.bold);
        assert!(!font.italic);
        assert!(font.underline.is_none());
        assert!(!font.strike);
        assert!(font.color.is_none());
        assert!(font.charset.is_none());
        assert!(font.family.is_none());
        assert!(font.scheme.is_none());
        assert!(font.script.is_none());
    }

    #[test]
    fn test_font_new() {
        let font = Font::new();
        assert!(font.name.is_none());
        assert!(!font.has_formatting());
    }

    #[test]
    fn test_font_has_formatting() {
        // Default font has no formatting
        let font = Font::default();
        assert!(!font.has_formatting());

        // Bold
        let font = Font {
            bold: true,
            ..Default::default()
        };
        assert!(font.has_formatting());

        // Italic
        let font = Font {
            italic: true,
            ..Default::default()
        };
        assert!(font.has_formatting());

        // Strike
        let font = Font {
            strike: true,
            ..Default::default()
        };
        assert!(font.has_formatting());

        // Underline
        let font = Font {
            underline: Some(Underline::Single),
            ..Default::default()
        };
        assert!(font.has_formatting());

        // Explicit no-underline and baseline do not add visible formatting.
        let font = Font {
            underline: Some(Underline::None),
            script: Some(Script::Baseline),
            ..Default::default()
        };
        assert!(!font.has_formatting());
    }

    #[test]
    fn test_font_full() {
        let font = Font {
            name: Some("Arial".to_string()),
            size: Some(12.0),
            bold: true,
            italic: false,
            underline: Some(Underline::Single),
            strike: false,
            color: Some(FontColor::rgb(Rgb::new(0xFF, 0, 0))),
            charset: Some(1),
            family: Some(2),
            scheme: Some(Scheme::Minor),
            script: Some(Script::Superscript),
        };
        assert_eq!(font.name, Some("Arial".to_string()));
        assert_eq!(font.size, Some(12.0));
        assert!(font.bold);
        assert!(!font.italic);
        assert_eq!(font.underline, Some(Underline::Single));
        assert!(!font.strike);
        assert_eq!(
            font.color.as_ref().map(|color| color.kind()),
            Some(FontColorKind::Rgb(Rgb::new(0xFF, 0, 0)))
        );
        assert_eq!(font.charset, Some(1));
        assert_eq!(font.family, Some(2));
        assert_eq!(font.scheme, Some(Scheme::Minor));
        assert_eq!(font.script, Some(Script::Superscript));
        assert!(font.has_formatting());
    }

    #[test]
    fn test_font_clone() {
        let font = Font {
            name: Some("Calibri".to_string()),
            size: Some(11.0),
            bold: false,
            italic: true,
            ..Default::default()
        };
        let font2 = font.clone();
        assert_eq!(font.name, font2.name);
        assert_eq!(font.italic, font2.italic);
    }

    #[test]
    fn fixed_font_domains_are_exhaustive_compact_and_strict() {
        let underlines = [
            (Underline::Single, "single"),
            (Underline::Double, "double"),
            (Underline::SingleAccounting, "singleAccounting"),
            (Underline::DoubleAccounting, "doubleAccounting"),
            (Underline::None, "none"),
        ];
        let schemes = [
            (Scheme::None, "none"),
            (Scheme::Major, "major"),
            (Scheme::Minor, "minor"),
        ];
        let scripts = [
            (Script::Baseline, "baseline"),
            (Script::Superscript, "superscript"),
            (Script::Subscript, "subscript"),
        ];

        for (value, token) in underlines {
            assert_eq!(value.as_str(), token);
            assert_eq!(value.to_string(), token);
            assert_eq!(token.parse::<Underline>(), Ok(value));
        }
        for (value, token) in schemes {
            assert_eq!(value.as_str(), token);
            assert_eq!(value.to_string(), token);
            assert_eq!(token.parse::<Scheme>(), Ok(value));
        }
        for (value, token) in scripts {
            assert_eq!(value.as_str(), token);
            assert_eq!(value.to_string(), token);
            assert_eq!(token.parse::<Script>(), Ok(value));
        }

        let error = "Single".parse::<Underline>().unwrap_err();
        assert_eq!(error.kind(), "underline");
        assert_eq!(error.value(), "Single");
        assert!("body".parse::<Scheme>().is_err());
        assert!("raised".parse::<Script>().is_err());
        assert_eq!(size_of::<Underline>(), 1);
        assert_eq!(size_of::<Scheme>(), 1);
        assert_eq!(size_of::<Script>(), 1);
    }
}
