//! Checked worksheet-name values and locale-independent identity semantics.

/// Worksheet-level `sheetPr` metadata.
pub mod properties;
/// Worksheet sparkline extension models and codecs.
pub mod sparklines;

use std::borrow::Cow;
use std::convert::Infallible;
use std::fmt;
use std::hash::{Hash, Hasher};
use std::ops::Deref;
use std::str::FromStr;

use caseless::Caseless;
use thiserror::Error;
use unicode_normalization::UnicodeNormalization;

pub use crate::workbook::{Selector, Visibility, Worksheet, WorksheetKind};

/// Maximum number of Unicode scalar values Office accepts in a sheet name.
pub const MAX_NAME_CHARS: usize = 31;

/// A validated, case-preserving Office worksheet name.
///
/// Equality and hashing use canonical, locale-independent Unicode caseless
/// matching because Office requires sheet names to be case-insensitively
/// unique in every locale. Display and serialization preserve the spelling
/// supplied by the developer.
#[derive(Debug, Clone)]
pub struct Name {
    value: Box<str>,
    key: Box<str>,
}

impl Name {
    /// Validate and copy a borrowed name.
    pub fn new(value: &str) -> Result<Self, NameError> {
        Self::try_from(value)
    }

    /// Borrow the case-preserving spelling.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.value
    }

    /// Move the case-preserving spelling out of this checked value.
    #[must_use]
    pub fn into_string(self) -> String {
        self.value.into_string()
    }

    /// Whether a borrowed spelling denotes the same Office sheet identity.
    #[must_use]
    pub fn matches(&self, other: &str) -> bool {
        folded(other).eq(self.key.chars())
    }

    pub(crate) fn identity_key(&self) -> &str {
        &self.key
    }

    fn from_string(value: String) -> Result<Self, NameError> {
        validate_str(&value)?;
        let key = key(&value);
        Ok(Self {
            value: value.into_boxed_str(),
            key,
        })
    }
}

impl PartialEq for Name {
    fn eq(&self, other: &Self) -> bool {
        self.key == other.key
    }
}

impl Eq for Name {}

impl Hash for Name {
    fn hash<H: Hasher>(&self, state: &mut H) {
        self.key.hash(state);
    }
}

impl AsRef<str> for Name {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

impl Deref for Name {
    type Target = str;

    fn deref(&self) -> &Self::Target {
        self.as_str()
    }
}

impl fmt::Display for Name {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}

impl FromStr for Name {
    type Err = NameError;

    fn from_str(value: &str) -> Result<Self, Self::Err> {
        Self::new(value)
    }
}

impl TryFrom<&str> for Name {
    type Error = NameError;

    fn try_from(value: &str) -> Result<Self, Self::Error> {
        Self::from_string(value.to_owned())
    }
}

impl TryFrom<&String> for Name {
    type Error = NameError;

    fn try_from(value: &String) -> Result<Self, Self::Error> {
        Self::try_from(value.as_str())
    }
}

impl TryFrom<String> for Name {
    type Error = NameError;

    fn try_from(value: String) -> Result<Self, Self::Error> {
        Self::from_string(value)
    }
}

impl TryFrom<Box<str>> for Name {
    type Error = NameError;

    fn try_from(value: Box<str>) -> Result<Self, Self::Error> {
        Self::from_string(value.into_string())
    }
}

impl<'a> From<&'a Name> for litchi_core::Selector<'a, Infallible> {
    fn from(value: &'a Name) -> Self {
        Self::Name(Cow::Borrowed(value.as_str()))
    }
}

impl From<Name> for litchi_core::Selector<'_, Infallible> {
    fn from(value: Name) -> Self {
        Self::Name(Cow::Owned(value.into_string()))
    }
}

/// Why a value cannot be an Office worksheet name.
#[derive(Debug, Clone, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum NameError {
    #[error("sheet name cannot be empty")]
    Empty,
    #[error("sheet name has {characters} characters; Office allows at most {MAX_NAME_CHARS}")]
    TooLong { characters: usize },
    #[error("sheet name cannot begin with an apostrophe")]
    LeadingApostrophe,
    #[error("sheet name cannot end with an apostrophe")]
    TrailingApostrophe,
    #[error("sheet name contains forbidden character {character:?}")]
    Forbidden { character: char },
    #[error("sheet name contains character {character:?}, which XML 1.0 cannot represent")]
    XmlCharacter { character: char },
}

pub(crate) fn key(value: &str) -> Box<str> {
    folded(value).collect::<String>().into_boxed_str()
}

pub(crate) fn equivalent(left: &str, right: &str) -> bool {
    folded(left).eq(folded(right))
}

fn folded(value: &str) -> impl Iterator<Item = char> + '_ {
    value.chars().nfd().default_case_fold().nfd()
}

pub(crate) fn validate_str(value: &str) -> Result<(), NameError> {
    if value.is_empty() {
        return Err(NameError::Empty);
    }
    let characters = value.chars().count();
    if characters > MAX_NAME_CHARS {
        return Err(NameError::TooLong { characters });
    }
    if value.starts_with('\'') {
        return Err(NameError::LeadingApostrophe);
    }
    if value.ends_with('\'') {
        return Err(NameError::TrailingApostrophe);
    }
    for character in value.chars() {
        if matches!(
            character,
            '\0' | '\u{3}' | '*' | '/' | ':' | '?' | '[' | '\\' | ']'
        ) {
            return Err(NameError::Forbidden { character });
        }
        if !is_xml_10_character(character) {
            return Err(NameError::XmlCharacter { character });
        }
    }
    Ok(())
}

const fn is_xml_10_character(character: char) -> bool {
    matches!(character, '\u{9}' | '\u{a}' | '\u{d}')
        || (character >= '\u{20}' && character <= '\u{d7ff}')
        || (character >= '\u{e000}' && character <= '\u{fffd}')
        || (character >= '\u{10000}' && character <= '\u{10ffff}')
}

#[cfg(test)]
mod tests {
    use std::collections::HashSet;

    use super::*;

    #[test]
    fn validates_office_name_domain() {
        assert_eq!(Name::new(""), Err(NameError::Empty));
        assert!(matches!(
            Name::new(&"x".repeat(32)),
            Err(NameError::TooLong { characters: 32 })
        ));
        assert_eq!(
            Name::new("bad/name"),
            Err(NameError::Forbidden { character: '/' })
        );
        assert_eq!(Name::new("'bad"), Err(NameError::LeadingApostrophe));
        assert_eq!(Name::new("bad'"), Err(NameError::TrailingApostrophe));
        assert!(Name::new("O'Brien 📈").is_ok());
    }

    #[test]
    fn equality_and_hashing_are_canonical_and_locale_independent() {
        let german = Name::new("Straße").expect("valid name");
        let upper = Name::new("STRASSE").expect("valid name");
        let composed = Name::new("Café").expect("valid name");
        let decomposed = Name::new("Cafe\u{301}").expect("valid name");
        assert_eq!(german, upper);
        assert_eq!(composed, decomposed);

        let mut names = HashSet::new();
        names.insert(german);
        assert!(!names.insert(upper));
        names.insert(composed);
        assert!(!names.insert(decomposed));
    }

    #[test]
    fn owned_conversion_reuses_the_supplied_spelling() {
        let name = Name::try_from(String::from("Data 2026")).expect("valid name");
        assert_eq!(name.as_str(), "Data 2026");
        assert!(name.matches("data 2026"));
    }
}
