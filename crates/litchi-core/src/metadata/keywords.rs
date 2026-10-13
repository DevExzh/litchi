//! Lossless keyword mixed content.
//!
//! [`Keywords`] keeps the exact item sequence an OOXML `cp:keywords`
//! element holds: plain text runs and structured `cp:value` children, each
//! with an optional validated `xml:lang`. Formats that store keywords as one
//! string use [`Keywords::joined`]; the wire spelling stays in each codec.

use std::fmt;
use std::str::FromStr;

/// Largest keyword text or language value accepted, in UTF-8 bytes.
pub const MAX_TEXT_BYTES: usize = 1_048_576;

/// Why a keyword value was refused.
#[derive(Debug, Clone, PartialEq, Eq, thiserror::Error)]
#[non_exhaustive]
pub enum Error {
    /// The value is not an `xsd:language` lexical value.
    #[error("invalid xml:lang value '{0}'")]
    InvalidLang(String),
    /// The value exceeds [`MAX_TEXT_BYTES`].
    #[error("keyword text of {observed} bytes exceeds the {maximum}-byte limit")]
    Limit {
        /// Observed UTF-8 length.
        observed: usize,
        /// Largest accepted length.
        maximum: usize,
    },
}

/// A validated `xml:lang` lexical value.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub struct Lang(String);

impl Lang {
    /// Validates and retains an `xml:lang` value without canonicalizing it.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Limit`] for an oversized value and
    /// [`Error::InvalidLang`] for a value outside `xsd:language`.
    pub fn new(value: impl Into<String>) -> Result<Self, Error> {
        let value = value.into();
        if value.len() > MAX_TEXT_BYTES {
            return Err(Error::Limit {
                observed: value.len(),
                maximum: MAX_TEXT_BYTES,
            });
        }
        let collapsed = collapse_xml_whitespace(&value);
        if collapsed.is_empty() {
            return Ok(Self(value));
        }
        let mut segments = collapsed.split('-');
        let Some(primary) = segments.next() else {
            return Err(Error::InvalidLang(value));
        };
        if primary.is_empty()
            || primary.len() > 8
            || !primary.bytes().all(|byte| byte.is_ascii_alphabetic())
            || segments.any(|segment| {
                segment.is_empty()
                    || segment.len() > 8
                    || !segment.bytes().all(|byte| byte.is_ascii_alphanumeric())
            })
        {
            return Err(Error::InvalidLang(value));
        }
        Ok(Self(value))
    }

    /// Returns the retained lexical value.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }

    /// Moves out the retained lexical value.
    #[must_use]
    pub fn into_string(self) -> String {
        self.0
    }
}

impl AsRef<str> for Lang {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

impl fmt::Display for Lang {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}

impl FromStr for Lang {
    type Err = Error;

    fn from_str(value: &str) -> Result<Self, Error> {
        Self::new(value)
    }
}

impl TryFrom<String> for Lang {
    type Error = Error;

    fn try_from(value: String) -> Result<Self, Error> {
        Self::new(value)
    }
}

/// A structured keyword value within [`Keywords`].
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Value {
    /// Keyword text.
    pub text: String,
    /// Optional language carried by this value.
    pub lang: Option<Lang>,
}

impl Value {
    /// Creates a value with no language annotation.
    #[must_use]
    pub fn new(text: impl Into<String>) -> Self {
        Self {
            text: text.into(),
            lang: None,
        }
    }

    /// Adds a validated language annotation.
    #[must_use]
    pub fn lang(mut self, lang: Lang) -> Self {
        self.lang = Some(lang);
        self
    }
}

impl From<String> for Value {
    fn from(text: String) -> Self {
        Self::new(text)
    }
}

impl From<&str> for Value {
    fn from(text: &str) -> Self {
        Self::new(text)
    }
}

/// One ordered item in the mixed content of [`Keywords`].
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Item {
    /// Text directly inside the keywords element.
    Text(String),
    /// A structured value child.
    Value(Value),
}

/// Lossless keyword mixed content.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Keywords {
    /// Optional language for the whole keyword content.
    pub lang: Option<Lang>,
    /// Text and value items in document order.
    pub items: Vec<Item>,
}

impl Keywords {
    /// Creates a present but empty keyword value.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Adds text directly inside the keywords element.
    #[must_use]
    pub fn text(mut self, text: impl Into<String>) -> Self {
        self.append_text(text.into());
        self
    }

    /// Adds a structured value child.
    #[must_use]
    pub fn value(mut self, value: impl Into<Value>) -> Self {
        self.items.push(Item::Value(value.into()));
        self
    }

    /// Adds a validated language to the whole keyword content.
    #[must_use]
    pub fn lang(mut self, lang: Lang) -> Self {
        self.lang = Some(lang);
        self
    }

    /// Concatenates all text and value content in document order.
    #[must_use]
    pub fn joined(&self) -> String {
        let capacity = self
            .items
            .iter()
            .map(|item| match item {
                Item::Text(text) => text.len(),
                Item::Value(value) => value.text.len(),
            })
            .sum();
        let mut joined = String::with_capacity(capacity);
        for item in &self.items {
            match item {
                Item::Text(text) => joined.push_str(text),
                Item::Value(value) => joined.push_str(&value.text),
            }
        }
        joined
    }

    /// Returns a direct borrow for the common plain-text form.
    #[must_use]
    pub fn plain(&self) -> Option<&str> {
        if self.lang.is_some() || self.items.len() != 1 {
            return None;
        }
        match &self.items[0] {
            Item::Text(text) => Some(text),
            Item::Value(_) => None,
        }
    }

    /// Appends text, merging it into a directly preceding text item.
    pub fn append_text(&mut self, text: String) {
        if text.is_empty() {
            return;
        }
        if let Some(Item::Text(previous)) = self.items.last_mut() {
            previous.push_str(&text);
        } else {
            self.items.push(Item::Text(text));
        }
    }
}

impl From<String> for Keywords {
    fn from(text: String) -> Self {
        Self::new().text(text)
    }
}

impl From<&str> for Keywords {
    fn from(text: &str) -> Self {
        Self::new().text(text)
    }
}

impl fmt::Display for Keywords {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(&self.joined())
    }
}

fn collapse_xml_whitespace(value: &str) -> String {
    value
        .split([' ', '\t', '\r', '\n'])
        .filter(|segment| !segment.is_empty())
        .collect::<Vec<_>>()
        .join(" ")
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn validates_language_without_rewriting_its_lexical_form() {
        let language = Lang::new(" en-US ").expect("xsd:language whitespace collapses");
        assert_eq!(language.as_str(), " en-US ");
        assert!(Lang::new("en_US").is_err());
        assert!(Lang::new("en-123456789").is_err());
        assert_eq!(Lang::new("").unwrap().as_str(), "");
    }

    #[test]
    fn rejects_oversized_language_before_whitespace_collapse() {
        let oversized = "a".repeat(MAX_TEXT_BYTES + 1);
        assert_eq!(
            Lang::new(oversized),
            Err(Error::Limit {
                observed: MAX_TEXT_BYTES + 1,
                maximum: MAX_TEXT_BYTES,
            })
        );
    }

    #[test]
    fn mixed_content_joins_in_order_and_plain_borrows_the_common_form() {
        let keywords = Keywords::new().text("a").value("b").text("c");
        assert_eq!(keywords.joined(), "abc");
        assert_eq!(keywords.plain(), None);
        assert_eq!(Keywords::from("plain").plain(), Some("plain"));
    }
}
