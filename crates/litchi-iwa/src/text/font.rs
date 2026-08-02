//! Strict font identities used by the shared iWork text layer.

use crate::{Error, Result};

/// A validated PostScript font name stored by Pages, Numbers, and Keynote.
#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct TextFontName(Box<str>);

impl TextFontName {
    /// Maximum UTF-8 length accepted for a native font identifier.
    pub const MAX_UTF8_BYTES: usize = 255;

    /// Validate and own a PostScript font name.
    pub fn new(name: impl Into<String>) -> Result<Self> {
        let name = name.into();
        validate_font_name(&name)?;
        Ok(Self(name.into_boxed_str()))
    }

    /// Borrow the native PostScript font name.
    pub fn as_str(&self) -> &str {
        &self.0
    }

    pub(crate) fn into_string(self) -> String {
        self.0.into()
    }
}

impl AsRef<str> for TextFontName {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

impl TryFrom<String> for TextFontName {
    type Error = Error;

    fn try_from(name: String) -> Result<Self> {
        Self::new(name)
    }
}

impl TryFrom<&str> for TextFontName {
    type Error = Error;

    fn try_from(name: &str) -> Result<Self> {
        Self::new(name)
    }
}

/// Effective font identity applied to uniformly styled text.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub enum TextFont {
    /// Use the native application default font.
    #[default]
    Default,
    /// Use a specific PostScript font identity.
    Named(TextFontName),
}

impl TextFont {
    /// Construct a named font after validating its PostScript identifier.
    pub fn named(name: impl Into<String>) -> Result<Self> {
        TextFontName::new(name).map(Self::Named)
    }

    /// Borrow the PostScript identifier, or `None` for the native default.
    pub fn name(&self) -> Option<&str> {
        match self {
            Self::Default => None,
            Self::Named(name) => Some(name.as_str()),
        }
    }
}

fn validate_font_name(name: &str) -> Result<()> {
    if name.is_empty() {
        return Err(Error::InvalidFormat(
            "text font name must not be empty".to_owned(),
        ));
    }
    if name.len() > TextFontName::MAX_UTF8_BYTES {
        return Err(Error::InvalidFormat(format!(
            "text font name exceeds {} UTF-8 bytes",
            TextFontName::MAX_UTF8_BYTES
        )));
    }
    if name.trim() != name {
        return Err(Error::InvalidFormat(
            "text font name must not have leading or trailing whitespace".to_owned(),
        ));
    }
    if name.chars().any(char::is_control) {
        return Err(Error::InvalidFormat(
            "text font name must not contain control characters".to_owned(),
        ));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn font_names_are_bounded_and_strict() {
        let name = TextFontName::new("AvenirNext-BoldItalic").unwrap();
        assert_eq!(name.as_str(), "AvenirNext-BoldItalic");
        assert!(TextFontName::new("").is_err());
        assert!(TextFontName::new(" Helvetica").is_err());
        assert!(TextFontName::new("Helvetica ").is_err());
        assert!(TextFontName::new("Helvet\0ica").is_err());
        assert!(TextFontName::new("x".repeat(TextFontName::MAX_UTF8_BYTES + 1)).is_err());
    }
}
