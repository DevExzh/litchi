//! Shared strict WordprocessingML color vocabulary.

/// A theme-color slot used by WordprocessingML formatting and web metadata.
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
    /// Parse the exact `ST_ThemeColor` lexical token.
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

    /// Return the exact `ST_ThemeColor` lexical token.
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

#[cfg(test)]
mod tests {
    use super::Theme;

    #[test]
    fn every_theme_color_round_trips() {
        for value in [
            Theme::Dark1,
            Theme::Light1,
            Theme::Dark2,
            Theme::Light2,
            Theme::Accent1,
            Theme::Accent2,
            Theme::Accent3,
            Theme::Accent4,
            Theme::Accent5,
            Theme::Accent6,
            Theme::Hyperlink,
            Theme::FollowedHyperlink,
            Theme::None,
            Theme::Background1,
            Theme::Text1,
            Theme::Background2,
            Theme::Text2,
        ] {
            assert_eq!(Theme::parse(value.as_str()), Some(value));
        }
        assert_eq!(Theme::parse("Accent1"), None);
    }
}
