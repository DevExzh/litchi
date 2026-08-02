//! Lossless `CT_Keywords` values.

use super::MAX_PROPERTY_TEXT;
use crate::{Error, Result};
use std::fmt;
use std::str::FromStr;

/// A validated `xml:lang` lexical value.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub struct Lang(String);

impl Lang {
    /// Validates and retains an `xml:lang` value without canonicalizing it.
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        if value.len() > MAX_PROPERTY_TEXT {
            return Err(Error::Limit {
                resource: "core property text bytes",
                max: MAX_PROPERTY_TEXT,
                actual: value.len(),
            });
        }
        let collapsed = collapse_xml_whitespace(&value);
        if collapsed.is_empty() {
            return Ok(Self(value));
        }
        let mut segments = collapsed.split('-');
        let Some(primary) = segments.next() else {
            return Err(invalid_lang(&value));
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
            return Err(invalid_lang(&value));
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

    fn from_str(value: &str) -> Result<Self> {
        Self::new(value)
    }
}

impl TryFrom<String> for Lang {
    type Error = Error;

    fn try_from(value: String) -> Result<Self> {
        Self::new(value)
    }
}

/// A `cp:value` child within [`super::Keywords`].
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

/// One ordered item in the mixed content of [`super::Keywords`].
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Item {
    /// Text directly inside `cp:keywords`.
    Text(String),
    /// A structured `cp:value` child.
    Value(Value),
}

pub(super) fn collapse_xml_whitespace(value: &str) -> String {
    value
        .split([' ', '\t', '\r', '\n'])
        .filter(|segment| !segment.is_empty())
        .collect::<Vec<_>>()
        .join(" ")
}

fn invalid_lang(value: &str) -> Error {
    Error::Invalid(format!("invalid xml:lang value '{value}'"))
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
        let oversized = "a".repeat(MAX_PROPERTY_TEXT + 1);
        assert!(matches!(
            Lang::new(oversized),
            Err(Error::Limit {
                resource: "core property text bytes",
                max: MAX_PROPERTY_TEXT,
                actual,
            }) if actual == MAX_PROPERTY_TEXT + 1
        ));
    }
}
