//! Semantic iWork application families.

use std::{fmt, str::FromStr};

/// Application family for an iWork document or native message namespace.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Application {
    /// Apple Pages.
    Pages,
    /// Apple Keynote.
    Keynote,
    /// Apple Numbers.
    Numbers,
    /// Common/shared native content.
    Common,
}

impl Application {
    /// Return the stable lowercase name used by configuration and diagnostics.
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Pages => "pages",
            Self::Keynote => "keynote",
            Self::Numbers => "numbers",
            Self::Common => "common",
        }
    }

    /// Return whether this is one of the three concrete iWork applications.
    pub const fn is_concrete(self) -> bool {
        !matches!(self, Self::Common)
    }
}

impl fmt::Display for Application {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}

/// Error returned when a string does not name a supported iWork application.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ParseError;

impl fmt::Display for ParseError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("unknown iWork application")
    }
}

impl std::error::Error for ParseError {}

impl FromStr for Application {
    type Err = ParseError;

    fn from_str(value: &str) -> Result<Self, Self::Err> {
        if value.eq_ignore_ascii_case("pages") {
            Ok(Self::Pages)
        } else if value.eq_ignore_ascii_case("keynote") {
            Ok(Self::Keynote)
        } else if value.eq_ignore_ascii_case("numbers") {
            Ok(Self::Numbers)
        } else if value.eq_ignore_ascii_case("common") {
            Ok(Self::Common)
        } else {
            Err(ParseError)
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_case_insensitively_and_formats_canonically() {
        assert_eq!("PaGeS".parse(), Ok(Application::Pages));
        assert_eq!("numbers".parse(), Ok(Application::Numbers));
        assert_eq!(Application::Keynote.to_string(), "keynote");
        assert!(Application::Pages.is_concrete());
        assert!(!Application::Common.is_concrete());
    }

    #[test]
    fn rejects_unknown_application_names() {
        assert_eq!("unknown".parse::<Application>(), Err(ParseError));
    }
}
