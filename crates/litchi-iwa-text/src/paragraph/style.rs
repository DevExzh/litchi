//! Archive-free paragraph-style identity and following-style values.
//!
//! Native object lookup, paragraph-style catalogs, protobuf fields, and package
//! mutation remain in the concrete IWA adapter. This module owns only the
//! checked semantic values that cross that boundary.

#![allow(
    clippy::module_name_repetitions,
    reason = "The explicit names distinguish paragraph-style identities at format boundaries."
)]

use std::{fmt, num::NonZeroU64};

/// Maximum UTF-8 byte length retained by a paragraph-style name.
pub const MAX_NAME_BYTES: usize = 255;

/// Validation failures produced by paragraph-style semantic constructors.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// A paragraph-style identifier was zero.
    ZeroId,
    /// A paragraph-style name was empty.
    EmptyName,
    /// A paragraph-style name exceeded [`MAX_NAME_BYTES`].
    NameTooLong { bytes: usize, maximum: usize },
    /// A paragraph-style name had leading or trailing whitespace.
    NameSurroundingWhitespace,
    /// A paragraph-style name contained a Unicode control character.
    NameControlCharacter { byte_index: usize },
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::ZeroId => formatter.write_str("paragraph-style identifier must be non-zero"),
            Self::EmptyName => formatter.write_str("paragraph-style name must not be empty"),
            Self::NameTooLong { bytes, maximum } => write!(
                formatter,
                "paragraph-style name uses {bytes} UTF-8 bytes; maximum is {maximum}"
            ),
            Self::NameSurroundingWhitespace => {
                formatter.write_str("paragraph-style name must not have surrounding whitespace")
            },
            Self::NameControlCharacter { byte_index } => write!(
                formatter,
                "paragraph-style name contains a control character at byte index {byte_index}"
            ),
        }
    }
}

impl std::error::Error for Error {}

/// Result returned by checked paragraph-style constructors.
pub type Result<T> = std::result::Result<T, Error>;

/// A validated, bounded user-visible paragraph-style name.
#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct ParagraphStyleName(Box<str>);

impl ParagraphStyleName {
    /// Validate a borrowed name before allocating its owned representation.
    ///
    /// # Errors
    ///
    /// Returns a typed [`Error`] when the name is empty, exceeds
    /// [`MAX_NAME_BYTES`], has surrounding whitespace, or contains a control
    /// character.
    pub fn new(input: impl AsRef<str>) -> Result<Self> {
        let value = input.as_ref();
        validate_name(value)?;
        Ok(Self(value.into()))
    }

    /// Validate and adopt an owned name without copying its string buffer.
    ///
    /// # Errors
    ///
    /// Returns the same validation errors as [`Self::new`].
    pub fn from_owned(value: String) -> Result<Self> {
        validate_name(&value)?;
        Ok(Self(value.into_boxed_str()))
    }

    /// Borrow the validated paragraph-style name.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }

    /// Consume the name and return its owned string without another copy.
    #[must_use]
    pub fn into_string(self) -> String {
        self.0.into()
    }
}

impl AsRef<str> for ParagraphStyleName {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

impl TryFrom<&str> for ParagraphStyleName {
    type Error = Error;

    fn try_from(value: &str) -> Result<Self> {
        Self::new(value)
    }
}

impl TryFrom<String> for ParagraphStyleName {
    type Error = Error;

    fn try_from(value: String) -> Result<Self> {
        Self::from_owned(value)
    }
}

/// A compact non-zero identifier for a paragraph style in one semantic
/// document snapshot.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct ParagraphStyleId(NonZeroU64);

impl ParagraphStyleId {
    /// Validate a paragraph-style identifier.
    ///
    /// # Errors
    ///
    /// Returns [`Error::ZeroId`] when `value` is zero.
    pub fn new(value: u64) -> Result<Self> {
        NonZeroU64::new(value).map(Self).ok_or(Error::ZeroId)
    }

    /// Return the numeric identifier used by the archive adapter.
    #[must_use]
    pub const fn get(self) -> u64 {
        self.0.get()
    }
}

impl TryFrom<u64> for ParagraphStyleId {
    type Error = Error;

    fn try_from(value: u64) -> Result<Self> {
        Self::new(value)
    }
}

/// One named paragraph-style preset in a semantic document snapshot.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct NamedParagraphStyle {
    id: ParagraphStyleId,
    name: ParagraphStyleName,
}

impl NamedParagraphStyle {
    /// Combine an already validated identifier and name without revalidating
    /// or duplicating the owned name.
    #[must_use]
    pub fn new(id: ParagraphStyleId, name: ParagraphStyleName) -> Self {
        Self { id, name }
    }

    /// Validate and construct a named paragraph style from a borrowed or
    /// otherwise string-like name.
    ///
    /// # Errors
    ///
    /// Returns the name validation errors from [`ParagraphStyleName::new`].
    pub fn try_new(id: ParagraphStyleId, name: impl AsRef<str>) -> Result<Self> {
        ParagraphStyleName::new(name).map(|validated_name| Self::new(id, validated_name))
    }

    /// Validate and construct a named paragraph style while adopting an owned
    /// name buffer without copying it.
    ///
    /// # Errors
    ///
    /// Returns the name validation errors from [`ParagraphStyleName::from_owned`].
    pub fn from_owned(id: ParagraphStyleId, name: String) -> Result<Self> {
        ParagraphStyleName::from_owned(name).map(|validated_name| Self::new(id, validated_name))
    }

    /// Return the compact paragraph-style identifier.
    #[must_use]
    pub const fn id(&self) -> ParagraphStyleId {
        self.id
    }

    /// Borrow the checked paragraph-style name.
    #[must_use]
    pub const fn name(&self) -> &ParagraphStyleName {
        &self.name
    }

    /// Consume the style and return its checked components.
    #[must_use]
    pub fn into_parts(self) -> (ParagraphStyleId, ParagraphStyleName) {
        (self.id, self.name)
    }
}

/// Paragraph style applied after the current paragraph.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum ParagraphFollowingStyle {
    /// Preserve the native “same style” behavior.
    #[default]
    Same,
    /// Apply one named paragraph style from the current semantic snapshot.
    Named(ParagraphStyleId),
}

impl ParagraphFollowingStyle {
    /// Select a named paragraph style for the following paragraph.
    #[must_use]
    pub const fn named(id: ParagraphStyleId) -> Self {
        Self::Named(id)
    }
}

fn validate_name(value: &str) -> Result<()> {
    if value.is_empty() {
        return Err(Error::EmptyName);
    }
    if value.len() > MAX_NAME_BYTES {
        return Err(Error::NameTooLong {
            bytes: value.len(),
            maximum: MAX_NAME_BYTES,
        });
    }
    if value.trim() != value {
        return Err(Error::NameSurroundingWhitespace);
    }
    if let Some((byte_index, _)) = value
        .char_indices()
        .find(|(_, character)| character.is_control())
    {
        return Err(Error::NameControlCharacter { byte_index });
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::mem::size_of;

    #[test]
    fn identifiers_are_nonzero_and_compact() {
        assert_eq!(ParagraphStyleId::new(0), Err(Error::ZeroId));
        assert_eq!(ParagraphStyleId::new(42).unwrap().get(), 42);
        assert_eq!(size_of::<ParagraphStyleId>(), size_of::<u64>());
    }

    #[test]
    fn names_reject_invalid_and_unbounded_input() {
        assert_eq!(ParagraphStyleName::new(""), Err(Error::EmptyName));
        assert_eq!(
            ParagraphStyleName::new(" Heading"),
            Err(Error::NameSurroundingWhitespace)
        );
        assert_eq!(
            ParagraphStyleName::new("Heading "),
            Err(Error::NameSurroundingWhitespace)
        );
        assert!(matches!(
            ParagraphStyleName::new("Bad\nName"),
            Err(Error::NameControlCharacter { .. })
        ));
        assert!(matches!(
            ParagraphStyleName::new("x".repeat(MAX_NAME_BYTES + 1)),
            Err(Error::NameTooLong { .. })
        ));
    }

    #[test]
    fn valid_names_are_bounded_owned_values() {
        let name = ParagraphStyleName::from_owned("Heading".to_owned()).unwrap();
        assert_eq!(name.as_str(), "Heading");
        assert_eq!(name.clone().into_string(), "Heading");

        let id = ParagraphStyleId::new(7).unwrap();
        let style = NamedParagraphStyle::new(id, name.clone());
        assert_eq!(style.id(), id);
        assert_eq!(style.name(), &name);
        assert_eq!(style.into_parts(), (id, name));
    }

    #[test]
    fn following_style_defaults_to_same_and_named_is_explicit() {
        assert_eq!(
            ParagraphFollowingStyle::default(),
            ParagraphFollowingStyle::Same
        );

        let id = ParagraphStyleId::new(9).unwrap();
        assert_eq!(
            ParagraphFollowingStyle::named(id),
            ParagraphFollowingStyle::Named(id)
        );
    }
}
