//! Archive-free bookmark values shared by iWork text adapters.
//!
//! This module owns only the validated semantic bookmark identity, display
//! settings, and UTF-16 range. Native object lookup, protobuf fields, and wire
//! mutation remain in the concrete IWA adapter.

use std::num::NonZeroU64;

use crate::position::TextRange;

/// Maximum UTF-8 byte length accepted for a bookmark display name.
pub const MAX_NAME_BYTES: usize = 1_024;

/// Validation failures produced while constructing bookmark values.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// The bookmark identifier is zero.
    ZeroId,
    /// The bookmark name is empty.
    EmptyName,
    /// The bookmark name exceeds [`MAX_NAME_BYTES`].
    NameTooLong,
    /// The bookmark name has whitespace at either boundary.
    NameSurroundingWhitespace,
    /// The bookmark name contains a Unicode control character.
    NameControlCharacter,
}

impl std::fmt::Display for Error {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let message = match self {
            Self::ZeroId => "text bookmark identifier must be non-zero",
            Self::EmptyName => "text bookmark name must not be empty",
            Self::NameTooLong => "text bookmark name exceeds 1024 UTF-8 bytes",
            Self::NameSurroundingWhitespace => {
                "text bookmark name must not have leading or trailing whitespace"
            },
            Self::NameControlCharacter => "text bookmark name must not contain control characters",
        };
        formatter.write_str(message)
    }
}

impl std::error::Error for Error {}

/// Result type for bookmark value construction.
pub type Result<T> = std::result::Result<T, Error>;

/// A non-zero semantic bookmark identity.
//
// The value is intentionally not described as an archive or object ID. A
// concrete adapter decides how its native identity is represented and uses
// this compact value only after crossing the archive boundary.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
#[repr(transparent)]
pub struct Id(NonZeroU64);

impl Id {
    /// Validate a bookmark identity.
    ///
    /// # Errors
    ///
    /// Returns [`Error::ZeroId`] when `value` is zero.
    pub fn new(value: u64) -> Result<Self> {
        NonZeroU64::new(value).map(Self).ok_or(Error::ZeroId)
    }

    /// Return the compact numeric identity for adapter-side lookup.
    #[must_use]
    pub const fn get(self) -> u64 {
        self.0.get()
    }
}

impl TryFrom<u64> for Id {
    type Error = Error;

    fn try_from(value: u64) -> Result<Self> {
        Self::new(value)
    }
}

/// A validated optional bookmark display name.
#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Name(Box<str>);

impl Name {
    /// Validate and own a bookmark display name.
    ///
    /// Validation happens before allocation for borrowed input. Owned
    /// `String` input can be converted without a second allocation through
    /// [`TryFrom<String>`].
    ///
    /// # Errors
    ///
    /// Returns a typed validation error when `value` is empty, too long,
    /// surrounded by whitespace, or contains a control character.
    pub fn new(value: &str) -> Result<Self> {
        validate_name(value)?;
        Ok(Self(value.into()))
    }

    /// Borrow the validated display name.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

impl AsRef<str> for Name {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

impl TryFrom<String> for Name {
    type Error = Error;

    fn try_from(value: String) -> Result<Self> {
        validate_name(&value)?;
        Ok(Self(value.into_boxed_str()))
    }
}

impl TryFrom<&str> for Name {
    type Error = Error;

    fn try_from(value: &str) -> Result<Self> {
        Self::new(value)
    }
}

/// Native bookmark visibility, preserving future discriminants.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum Visibility {
    /// The bookmark is shown in the ordinary bookmark presentation.
    #[default]
    Visible,
    /// The bookmark is hidden in the ordinary bookmark presentation.
    Hidden,
    /// A value introduced by a newer iWork version.
    Unknown(u32),
}

/// Validated settings for one bookmark field.
#[derive(Debug, Clone, PartialEq, Eq, Hash, Default)]
pub struct Settings {
    name: Option<Name>,
    visibility: Visibility,
}

impl Settings {
    /// Create visible, unnamed settings.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            name: None,
            visibility: Visibility::Visible,
        }
    }

    /// Set a validated display name.
    #[must_use]
    pub fn with_name(mut self, name: Name) -> Self {
        self.name = Some(name);
        self
    }

    /// Set the semantic visibility value.
    #[must_use]
    pub const fn with_visibility(mut self, visibility: Visibility) -> Self {
        self.visibility = visibility;
        self
    }

    /// Borrow the optional display name.
    #[must_use]
    pub fn name(&self) -> Option<&Name> {
        self.name.as_ref()
    }

    /// Return the semantic visibility value.
    #[must_use]
    pub const fn visibility(&self) -> Visibility {
        self.visibility
    }
}

/// One semantic bookmark attached to a nonempty UTF-16 text range.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Bookmark {
    /// Semantic identity assigned by the owning adapter.
    pub id: Id,
    /// Half-open UTF-16 text range covered by the bookmark.
    pub range: TextRange,
    /// Display settings for the bookmark.
    pub settings: Settings,
}

impl Bookmark {
    /// Construct a bookmark from validated semantic components.
    #[must_use]
    pub const fn new(id: Id, range: TextRange, settings: Settings) -> Self {
        Self {
            id,
            range,
            settings,
        }
    }
}

fn validate_name(value: &str) -> Result<()> {
    if value.is_empty() {
        return Err(Error::EmptyName);
    }
    if value.len() > MAX_NAME_BYTES {
        return Err(Error::NameTooLong);
    }
    if value.trim() != value {
        return Err(Error::NameSurroundingWhitespace);
    }
    if value.chars().any(char::is_control) {
        return Err(Error::NameControlCharacter);
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn ids_are_nonzero_and_compact() {
        assert_eq!(Id::new(0), Err(Error::ZeroId));
        assert_eq!(Id::new(7).unwrap().get(), 7);
    }

    #[test]
    fn names_are_strict_and_bounded() {
        assert_eq!(Name::new("Methods").unwrap().as_str(), "Methods");
        assert!(matches!(Name::new(""), Err(Error::EmptyName)));
        assert!(matches!(
            Name::new(" padded"),
            Err(Error::NameSurroundingWhitespace)
        ));
        assert!(matches!(
            Name::new("padded "),
            Err(Error::NameSurroundingWhitespace)
        ));
        assert!(matches!(
            Name::new("line\nbreak"),
            Err(Error::NameControlCharacter)
        ));
        assert!(matches!(
            Name::new(&"x".repeat(MAX_NAME_BYTES + 1)),
            Err(Error::NameTooLong)
        ));
    }

    #[test]
    fn owned_names_do_not_need_a_second_string_allocation() {
        assert_eq!(
            Name::try_from("Methods".to_owned()).map(|name| name.as_str().to_owned()),
            Ok("Methods".to_owned())
        );
    }

    #[test]
    fn settings_and_bookmarks_are_semantic_values() {
        let settings = Settings::new()
            .with_name(Name::new("Methods").unwrap())
            .with_visibility(Visibility::Hidden);
        let bookmark = Bookmark::new(
            Id::new(9).unwrap(),
            TextRange::from_utf16_indexes(2, 7).unwrap(),
            settings.clone(),
        );
        assert_eq!(bookmark.id.get(), 9);
        assert_eq!(bookmark.settings, settings);
        assert_eq!(bookmark.settings.name().unwrap().as_str(), "Methods");
        assert_eq!(bookmark.settings.visibility(), Visibility::Hidden);
    }
}
