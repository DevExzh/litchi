//! Strict language values shared by native text attributes.

const MAX_LANGUAGE_TAG_BYTES: usize = 255;
const MAX_LANGUAGE_SUBTAG_BYTES: usize = 8;

/// Validation failures produced while constructing a text-language value.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// The language tag contains no bytes.
    Empty,
    /// The language tag exceeds the native byte limit.
    TooLong,
    /// The language tag contains an empty subtag.
    EmptySubtag,
    /// The language tag contains a character outside the ASCII tag alphabet.
    InvalidCharacter,
    /// The primary language subtag is not alphabetic.
    PrimaryNotAlphabetic,
    /// A language subtag exceeds the native subtag limit.
    SubtagTooLong,
}

impl std::fmt::Display for Error {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let message = match self {
            Self::Empty => "iWork text language tag cannot be empty",
            Self::TooLong => "iWork text language tag exceeds 255 bytes",
            Self::EmptySubtag => "iWork text language tag contains an empty subtag",
            Self::InvalidCharacter => {
                "iWork text language tag must contain only ASCII letters, digits, and hyphens"
            },
            Self::PrimaryNotAlphabetic => {
                "iWork text language tag must begin with an ASCII language subtag"
            },
            Self::SubtagTooLong => "iWork text language subtags cannot exceed eight bytes",
        };
        formatter.write_str(message)
    }
}

impl std::error::Error for Error {}

/// Result type for text-language value construction.
pub type Result<T> = std::result::Result<T, Error>;

/// A validated BCP 47-style language tag, preserving native spelling and case.
#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct TextLanguageTag(Box<str>);

impl TextLanguageTag {
    /// Maximum UTF-8 byte length accepted for a native language tag.
    pub const MAX_BYTES: usize = MAX_LANGUAGE_TAG_BYTES;

    /// Maximum UTF-8 byte length accepted for one language subtag.
    pub const MAX_SUBTAG_BYTES: usize = MAX_LANGUAGE_SUBTAG_BYTES;

    /// Construct a validated language tag such as `en`, `fr-CA`, or `zh-Hant`.
    ///
    /// # Errors
    ///
    /// Returns [`Error`] when the tag is empty, exceeds the native size
    /// limits, or is not composed of valid ASCII language subtags.
    pub fn new(tag: impl Into<Box<str>>) -> Result<Self> {
        let boxed_tag = tag.into();
        validate_language_tag(&boxed_tag)?;
        Ok(Self(boxed_tag))
    }

    /// Borrow the language tag exactly as stored by iWork.
    #[must_use]
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
#[allow(
    clippy::module_name_repetitions,
    reason = "TextLanguage names the semantic value represented by this language module."
)]
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
    ///
    /// # Errors
    ///
    /// Returns [`Error`] when `tag` is not a valid language tag.
    pub fn tag(tag: impl Into<Box<str>>) -> Result<Self> {
        TextLanguageTag::new(tag).map(Self::Tag)
    }

    /// Return the explicit tag, or `None` for automatic language selection.
    #[must_use]
    pub fn as_tag(&self) -> Option<&TextLanguageTag> {
        match self {
            Self::Automatic => None,
            Self::Tag(tag) => Some(tag),
        }
    }
}

/// One explicit language-table boundary in a text storage.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct TextLanguageRun {
    /// UTF-16 boundary where this language begins.
    pub position: crate::position::TextPosition,
    /// Language effective from this boundary.
    pub language: TextLanguage,
}

impl TextLanguageRun {
    /// Construct a typed language boundary.
    #[must_use]
    pub const fn new(position: crate::position::TextPosition, language: TextLanguage) -> Self {
        Self { position, language }
    }
}

fn validate_language_tag(tag: &str) -> Result<()> {
    if tag.is_empty() {
        return Err(Error::Empty);
    }
    if tag.len() > MAX_LANGUAGE_TAG_BYTES {
        return Err(Error::TooLong);
    }
    if tag.starts_with('-') || tag.ends_with('-') || tag.contains("--") {
        return Err(Error::EmptySubtag);
    }
    if !tag
        .bytes()
        .all(|byte| byte.is_ascii_alphanumeric() || byte == b'-')
    {
        return Err(Error::InvalidCharacter);
    }
    let mut subtags = tag.split('-');
    let Some(primary) = subtags.next() else {
        return Err(Error::Empty);
    };
    if !primary.bytes().all(|byte| byte.is_ascii_alphabetic()) {
        return Err(Error::PrimaryNotAlphabetic);
    }
    if primary.len() > MAX_LANGUAGE_SUBTAG_BYTES
        || subtags.any(|subtag| subtag.len() > MAX_LANGUAGE_SUBTAG_BYTES)
    {
        return Err(Error::SubtagTooLong);
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::position::TextPosition;

    #[test]
    fn language_tags_preserve_spelling_and_support_automatic_selection() {
        let tag = TextLanguageTag::new("fr-CA");
        assert_eq!(tag.as_ref().map(TextLanguageTag::as_str), Ok("fr-CA"));
        assert_eq!(TextLanguage::default(), TextLanguage::Automatic);
        assert_eq!(
            TextLanguage::tag("zh-Hant")
                .map(|language| language.as_tag().map(|tag| tag.as_str().to_owned())),
            Ok(Some("zh-Hant".to_owned()))
        );
        assert_eq!(
            TextLanguageRun::new(TextPosition::ZERO, TextLanguage::Automatic).position,
            TextPosition::ZERO
        );
    }

    #[test]
    fn invalid_language_tags_report_typed_failures() {
        assert_eq!(TextLanguageTag::new(""), Err(Error::Empty));
        assert_eq!(
            TextLanguageTag::new("x".repeat(MAX_LANGUAGE_TAG_BYTES + 1)),
            Err(Error::TooLong)
        );
        assert_eq!(TextLanguageTag::new("-fr"), Err(Error::EmptySubtag));
        assert_eq!(TextLanguageTag::new("fr--CA"), Err(Error::EmptySubtag));
        assert_eq!(TextLanguageTag::new("fr_CA"), Err(Error::InvalidCharacter));
        assert_eq!(
            TextLanguageTag::new("123"),
            Err(Error::PrimaryNotAlphabetic)
        );
        assert_eq!(
            TextLanguageTag::new("toolongtag"),
            Err(Error::SubtagTooLong)
        );
    }
}
