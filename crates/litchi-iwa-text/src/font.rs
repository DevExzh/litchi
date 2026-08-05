//! Strict font identities used by shared iWork text values.

/// Why a font name was rejected before it was stored.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum NameError {
    /// The name contains no UTF-8 bytes.
    Empty,
    /// The name exceeds the native identifier limit.
    TooLong,
    /// The name has whitespace at either boundary.
    SurroundingWhitespace,
    /// The name contains a Unicode control character.
    ControlCharacter,
}

impl std::fmt::Display for NameError {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let message = match self {
            Self::Empty => "text font name must not be empty",
            Self::TooLong => "text font name exceeds 255 UTF-8 bytes",
            Self::SurroundingWhitespace => {
                "text font name must not have leading or trailing whitespace"
            },
            Self::ControlCharacter => "text font name must not contain control characters",
        };
        formatter.write_str(message)
    }
}

impl std::error::Error for NameError {}

/// A validated PostScript font name stored by Pages, Numbers, and Keynote.
#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Name(Box<str>);

impl Name {
    /// Maximum UTF-8 length accepted for a native font identifier.
    pub const MAX_UTF8_BYTES: usize = 255;

    /// Validate and own a PostScript font name.
    ///
    /// Validation precedes allocation for borrowed inputs, so rejected names
    /// do not create a second owned copy in the semantic layer.
    ///
    /// # Errors
    ///
    /// Returns [`NameError`] when the native identifier is empty, too long,
    /// surrounded by whitespace, or contains a control character.
    pub fn new(name: &str) -> Result<Self, NameError> {
        validate(name)?;
        Ok(Self(name.into()))
    }

    /// Borrow the native PostScript font name.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }

    /// Consume the validated name as an owned string.
    #[must_use]
    pub fn into_string(self) -> String {
        self.0.into()
    }
}

impl AsRef<str> for Name {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

impl TryFrom<String> for Name {
    type Error = NameError;

    fn try_from(name: String) -> Result<Self, Self::Error> {
        validate(&name)?;
        Ok(Self(name.into_boxed_str()))
    }
}

impl TryFrom<&str> for Name {
    type Error = NameError;

    fn try_from(name: &str) -> Result<Self, Self::Error> {
        Self::new(name)
    }
}

/// Effective font identity applied to uniformly styled text.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub enum Font {
    /// Use the native application default font.
    #[default]
    Default,
    /// Use a specific PostScript font identity.
    Named(Name),
}

impl Font {
    /// Construct a named font after validating its PostScript identifier.
    ///
    /// # Errors
    ///
    /// Returns [`NameError`] when the native identifier fails validation.
    pub fn named(name: &str) -> Result<Self, NameError> {
        Name::new(name).map(Self::Named)
    }

    /// Borrow the PostScript identifier, or `None` for the native default.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        match self {
            Self::Default => None,
            Self::Named(name) => Some(name.as_str()),
        }
    }
}

impl From<Name> for Font {
    fn from(name: Name) -> Self {
        Self::Named(name)
    }
}

fn validate(name: &str) -> Result<(), NameError> {
    if name.is_empty() {
        return Err(NameError::Empty);
    }
    if name.len() > Name::MAX_UTF8_BYTES {
        return Err(NameError::TooLong);
    }
    if name.trim() != name {
        return Err(NameError::SurroundingWhitespace);
    }
    if name.chars().any(char::is_control) {
        return Err(NameError::ControlCharacter);
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn names_are_bounded_and_strict() {
        assert_eq!(
            Name::new("AvenirNext-BoldItalic").map(|name| name.as_str().to_owned()),
            Ok("AvenirNext-BoldItalic".to_owned())
        );
        assert!(matches!(Name::new(""), Err(NameError::Empty)));
        assert!(matches!(
            Name::new(" Helvetica"),
            Err(NameError::SurroundingWhitespace)
        ));
        assert!(matches!(
            Name::new("Helvetica "),
            Err(NameError::SurroundingWhitespace)
        ));
        assert!(matches!(
            Name::new("Helvet\0ica"),
            Err(NameError::ControlCharacter)
        ));
        assert!(matches!(
            Name::try_from("x".repeat(Name::MAX_UTF8_BYTES + 1)),
            Err(NameError::TooLong)
        ));
    }

    #[test]
    fn owned_input_is_validated_without_reallocating_the_box() {
        assert_eq!(
            Name::try_from("AvenirNext-Regular".to_owned()).map(Name::into_string),
            Ok("AvenirNext-Regular".to_owned())
        );
    }

    #[test]
    fn font_has_a_compact_default_and_named_form() {
        assert_eq!(Font::default().name(), None);
        assert_eq!(
            Font::named("Menlo-Regular").map(|font| font.name().map(str::to_owned)),
            Ok(Some("Menlo-Regular".to_owned()))
        );
        assert_eq!(
            Name::new("Helvetica")
                .map(Font::from)
                .map(|font| font.name().map(str::to_owned)),
            Ok(Some("Helvetica".to_owned()))
        );
    }
}
