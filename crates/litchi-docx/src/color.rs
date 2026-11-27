//! Shared strict `WordprocessingML` color vocabulary.

use std::str::FromStr;

use crate::{Error, Result};

/// A theme-color slot used by `WordprocessingML` formatting and web metadata.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Theme {
    Dark1,
    Light1,
    Dark2,
    Light2,
    Accent1,
    Accent2,
    Accent3,
    Accent4,
    Accent5,
    Accent6,
    Hyperlink,
    FollowedHyperlink,
    None,
    Background1,
    Text1,
    Background2,
    Text2,
}

impl Theme {
    /// Every value in the `WordprocessingML` `ST_ThemeColor` domain.
    ///
    /// The array is ordered according to the checked-in `WordprocessingML`
    /// vocabulary and is allocation-free to iterate. It deliberately does
    /// not include DrawingML-only `ST_SchemeColorVal` tokens such as `dk1` or
    /// `hlink`.
    pub const ALL: [Self; 17] = [
        Self::Dark1,
        Self::Light1,
        Self::Dark2,
        Self::Light2,
        Self::Accent1,
        Self::Accent2,
        Self::Accent3,
        Self::Accent4,
        Self::Accent5,
        Self::Accent6,
        Self::Hyperlink,
        Self::FollowedHyperlink,
        Self::None,
        Self::Background1,
        Self::Text1,
        Self::Background2,
        Self::Text2,
    ];

    /// Parse the exact `ST_ThemeColor` lexical token.
    ///
    /// This compatibility parser retains the original `Option` API. New
    /// callers should use [`Self::parse_str`] or `str::parse`, which preserve
    /// the distinction between an invalid token and an absent optional
    /// attribute.
    #[must_use]
    pub fn parse(value: &str) -> Option<Self> {
        match value {
            "dark1" => Some(Self::Dark1),
            "light1" => Some(Self::Light1),
            "dark2" => Some(Self::Dark2),
            "light2" => Some(Self::Light2),
            "accent1" => Some(Self::Accent1),
            "accent2" => Some(Self::Accent2),
            "accent3" => Some(Self::Accent3),
            "accent4" => Some(Self::Accent4),
            "accent5" => Some(Self::Accent5),
            "accent6" => Some(Self::Accent6),
            "hyperlink" => Some(Self::Hyperlink),
            "followedHyperlink" => Some(Self::FollowedHyperlink),
            "none" => Some(Self::None),
            "background1" => Some(Self::Background1),
            "text1" => Some(Self::Text1),
            "background2" => Some(Self::Background2),
            "text2" => Some(Self::Text2),
            _ => None,
        }
    }

    /// Parse an exact `ST_ThemeColor` token with a typed DOCX error.
    ///
    /// `WordprocessingML` theme colors are case-sensitive. The parser rejects
    /// similarly named `DrawingML` scheme-color tokens instead of silently
    /// widening this format-specific domain.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn parse_str(value: &str) -> Result<Self> {
        Self::parse(value).ok_or_else(|| Error::Invalid(format!("invalid theme color '{value}'")))
    }

    /// Return the exact `ST_ThemeColor` lexical token.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Dark1 => "dark1",
            Self::Light1 => "light1",
            Self::Dark2 => "dark2",
            Self::Light2 => "light2",
            Self::Accent1 => "accent1",
            Self::Accent2 => "accent2",
            Self::Accent3 => "accent3",
            Self::Accent4 => "accent4",
            Self::Accent5 => "accent5",
            Self::Accent6 => "accent6",
            Self::Hyperlink => "hyperlink",
            Self::FollowedHyperlink => "followedHyperlink",
            Self::None => "none",
            Self::Background1 => "background1",
            Self::Text1 => "text1",
            Self::Background2 => "background2",
            Self::Text2 => "text2",
        }
    }
}

impl FromStr for Theme {
    type Err = Error;

    fn from_str(value: &str) -> Result<Self> {
        Self::parse_str(value)
    }
}

#[cfg(test)]
mod tests {
    use super::Theme;

    #[test]
    fn every_theme_color_round_trips() {
        for value in Theme::ALL {
            assert_eq!(Theme::parse(value.as_str()), Some(value));
            assert!(matches!(
                Theme::parse_str(value.as_str()),
                Ok(parsed) if parsed == value
            ));
            assert!(matches!(
                value.as_str().parse::<Theme>(),
                Ok(parsed) if parsed == value
            ));
        }
        assert_eq!(Theme::parse("Accent1"), None);
    }

    #[test]
    fn strict_parser_rejects_invalid_and_other_color_domains() {
        for value in ["", "Accent1", "accent7", "dk1", "hlink", "phClr"] {
            assert!(Theme::parse_str(value).is_err(), "{value}");
            assert!(value.parse::<Theme>().is_err(), "{value}");
            assert_eq!(Theme::parse(value), None, "{value}");
        }

        assert!(matches!(
            Theme::parse_str("accent7"),
            Err(crate::Error::Invalid(message)) if message == "invalid theme color 'accent7'"
        ));
    }
}
