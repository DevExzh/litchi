//! Strict public types for iWork text-language runs.

use crate::{Error, Result};

use super::position::TextPosition;

const MAX_LANGUAGE_TAG_BYTES: usize = 255;

/// A validated BCP 47-style language tag, preserving native spelling and case.
#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct TextLanguageTag(Box<str>);

impl TextLanguageTag {
    /// Construct a validated language tag such as `en`, `fr-CA`, or `zh-Hant`.
    pub fn new(tag: impl Into<Box<str>>) -> Result<Self> {
        let tag = tag.into();
        validate_language_tag(&tag)?;
        Ok(Self(tag))
    }

    /// Borrow the language tag exactly as stored by iWork.
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

impl AsRef<str> for TextLanguageTag {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

/// The language effective from one text boundary onward.
#[derive(Debug, Clone, PartialEq, Eq, Hash, Default)]
pub enum TextLanguage {
    /// Use iWork's automatic or document-level language selection.
    #[default]
    Automatic,
    /// Use an explicit BCP 47-style language tag.
    Tag(TextLanguageTag),
}

impl TextLanguage {
    /// Construct an explicit language value.
    pub fn tag(tag: impl Into<Box<str>>) -> Result<Self> {
        TextLanguageTag::new(tag).map(Self::Tag)
    }

    /// Return the explicit tag, or `None` for automatic language selection.
    pub fn as_tag(&self) -> Option<&TextLanguageTag> {
        match self {
            Self::Automatic => None,
            Self::Tag(tag) => Some(tag),
        }
    }

    pub(crate) fn from_native(value: Option<&str>) -> Result<Self> {
        value
            .map(|tag| TextLanguageTag::new(tag.to_owned().into_boxed_str()).map(Self::Tag))
            .transpose()
            .map(Option::unwrap_or_default)
    }

    pub(crate) fn native_value(&self) -> Option<&str> {
        self.as_tag().map(TextLanguageTag::as_str)
    }
}

/// One explicit language-table boundary in a text storage.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct TextLanguageRun {
    /// UTF-16 boundary where this language begins.
    pub position: TextPosition,
    /// Language effective from this boundary.
    pub language: TextLanguage,
}

impl TextLanguageRun {
    /// Construct a typed language boundary.
    pub fn new(position: TextPosition, language: TextLanguage) -> Self {
        Self { position, language }
    }
}

fn validate_language_tag(tag: &str) -> Result<()> {
    if tag.is_empty() {
        return Err(Error::ParseError(
            "iWork text language tag cannot be empty".to_owned(),
        ));
    }
    if tag.len() > MAX_LANGUAGE_TAG_BYTES {
        return Err(Error::ParseError(format!(
            "iWork text language tag exceeds {MAX_LANGUAGE_TAG_BYTES} bytes"
        )));
    }
    if tag.starts_with('-') || tag.ends_with('-') || tag.contains("--") {
        return Err(Error::ParseError(
            "iWork text language tag contains an empty subtag".to_owned(),
        ));
    }
    if !tag
        .bytes()
        .all(|byte| byte.is_ascii_alphanumeric() || byte == b'-')
    {
        return Err(Error::ParseError(
            "iWork text language tag must contain only ASCII letters, digits, and hyphens"
                .to_owned(),
        ));
    }
    let mut subtags = tag.split('-');
    let primary = subtags.next().expect("nonempty tag has a primary subtag");
    if !primary.bytes().all(|byte| byte.is_ascii_alphabetic()) {
        return Err(Error::ParseError(
            "iWork text language tag must begin with an ASCII language subtag".to_owned(),
        ));
    }
    if primary.len() > 8 || subtags.any(|subtag| subtag.len() > 8) {
        return Err(Error::ParseError(
            "iWork text language subtags cannot exceed eight bytes".to_owned(),
        ));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn language_tags_and_positions_are_strictly_typed() {
        assert_eq!(
            TextLanguage::tag("fr-CA")
                .unwrap()
                .as_tag()
                .unwrap()
                .as_str(),
            "fr-CA"
        );
        for invalid in ["", " fr", "fr ", "fr_CA", "fr--CA", "toolongtag", "123"] {
            assert!(TextLanguage::tag(invalid).is_err(), "accepted {invalid:?}");
        }
        assert_eq!(TextPosition::from_utf16_index(7).unwrap().utf16_index(), 7);
    }
}
