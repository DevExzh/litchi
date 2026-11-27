//! Lossless keyword mixed content.
//!
//! [`Keywords`] keeps the exact item sequence an OOXML `cp:keywords`
//! element holds: plain text runs and structured `cp:value` children, each
//! with an optional validated `xml:lang`. The wire spelling stays in each
//! codec.
//!
//! # The keyword list
//!
//! [`Metadata::keywords`](crate::Metadata::keywords) is a list. Formats
//! that store keywords as one string (OOXML `cp:keywords` text, the OLE
//! summary information, RTF `\keywords`, iWork) follow one rule, the one
//! Office producers use:
//!
//! - **Reading** ([`split`], [`Keywords::list`]): the text is split at
//!   every `;` and `,`; each piece is trimmed of whitespace and empty pieces
//!   are dropped. An OOXML `cp:value` child is one keyword whatever it
//!   contains.
//! - **Writing** ([`join`], [`Keywords::from_list`]): the keywords are
//!   joined with `, `. A keyword that the split would not give back (one
//!   that contains `;` or `,`, has surrounding whitespace, or is empty)
//!   cannot be stored in a plain string: [`join`] refuses it with
//!   [`Error::Separator`], while OOXML writes it as a `cp:value` child.
//!
//! The OOXML codec writes an unchanged list from the source's own
//! [`Keywords`], so its bytes survive a round trip whatever separators the
//! producer used.

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
    /// A keyword contains a separator (`;` or `,`), has surrounding
    /// whitespace, or is empty, so a single keyword string cannot keep it.
    #[error("a keyword cannot be stored in a separated keyword string")]
    Separator,
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

    /// The keyword list: text items split per [`split`], each `cp:value`
    /// one keyword (trimmed; an empty value is dropped).
    #[must_use]
    pub fn list(&self) -> Vec<String> {
        let mut list = Vec::new();
        for item in &self.items {
            match item {
                Item::Text(text) => push_split(&mut list, text),
                Item::Value(value) => {
                    let keyword = value.text.trim();
                    if !keyword.is_empty() {
                        list.push(keyword.to_owned());
                    }
                },
            }
        }
        list
    }

    /// Mixed content for a keyword list: plain keywords as text joined with
    /// `, `, and each keyword [`split`] would not give back as a `cp:value`
    /// child, so [`Self::list`] gives every keyword back (trimmed). Empty
    /// keywords are dropped.
    #[must_use]
    pub fn from_list(list: &[String]) -> Self {
        let mut keywords = Self::new();
        for (index, keyword) in list
            .iter()
            .filter(|keyword| !keyword.trim().is_empty())
            .enumerate()
        {
            if index > 0 {
                keywords.append_text(JOINER.to_owned());
            }
            if is_plain(keyword) {
                keywords.append_text(keyword.clone());
            } else {
                keywords
                    .items
                    .push(Item::Value(Value::new(keyword.clone())));
            }
        }
        keywords
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

/// The separators [`split`] recognizes.
pub const SEPARATORS: [char; 2] = [';', ','];

/// The separator [`join`] writes.
pub const JOINER: &str = ", ";

/// Split a keyword string into its keywords: at every `;` and `,`, each
/// piece trimmed, empty pieces dropped.
///
/// ```
/// use litchi_core::metadata::keywords::split;
///
/// assert_eq!(split("q4, review;strategy ;"), ["q4", "review", "strategy"]);
/// assert!(split(" ; ").is_empty());
/// ```
#[must_use]
pub fn split(text: &str) -> Vec<String> {
    let mut list = Vec::new();
    push_split(&mut list, text);
    list
}

fn push_split(list: &mut Vec<String>, text: &str) {
    list.extend(
        text.split(SEPARATORS)
            .map(str::trim)
            .filter(|keyword| !keyword.is_empty())
            .map(str::to_owned),
    );
}

/// Whether [`split`] gives `keyword` back as one keyword.
#[must_use]
pub fn is_plain(keyword: &str) -> bool {
    !keyword.is_empty() && keyword.trim() == keyword && !keyword.contains(SEPARATORS)
}

/// Join keywords into one string with `, `, the inverse of [`split`].
///
/// ```
/// use litchi_core::metadata::keywords::{join, split};
///
/// let list = ["q4".to_owned(), "review".to_owned()];
/// assert_eq!(join(&list).unwrap(), "q4, review");
/// assert_eq!(split(&join(&list).unwrap()), list);
/// assert!(join(&["a,b".to_owned()]).is_err());
/// ```
///
/// # Errors
///
/// [`Error::Separator`] for a keyword [`split`] would not give back.
pub fn join(list: &[String]) -> Result<String, Error> {
    if !list.iter().all(|keyword| is_plain(keyword)) {
        return Err(Error::Separator);
    }
    Ok(list.join(JOINER))
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
    fn list_splits_text_and_keeps_values_whole() {
        let keywords = Keywords::new()
            .text("lead")
            .value("colour, hue")
            .text("; a,b ;")
            .value("  ");
        assert_eq!(keywords.list(), ["lead", "colour, hue", "a", "b"]);
        let list = [
            "a".to_owned(),
            "b; c".to_owned(),
            " d".to_owned(),
            String::new(),
        ];
        let built = Keywords::from_list(&list);
        assert_eq!(built.list(), ["a", "b; c", "d"]);
        assert_eq!(
            built.items,
            [
                Item::Text("a, ".to_owned()),
                Item::Value(Value::new("b; c")),
                Item::Text(", ".to_owned()),
                Item::Value(Value::new(" d")),
            ]
        );
        assert_eq!(
            Keywords::from_list(&["x".to_owned(), "y".to_owned()]).plain(),
            Some("x, y")
        );
        assert_eq!(join(&[" d".to_owned()]), Err(Error::Separator));
        assert_eq!(join(&[]), Ok(String::new()));
    }

    #[test]
    fn mixed_content_joins_in_order_and_plain_borrows_the_common_form() {
        let keywords = Keywords::new().text("a").value("b").text("c");
        assert_eq!(keywords.joined(), "abc");
        assert_eq!(keywords.plain(), None);
        assert_eq!(Keywords::from("plain").plain(), Some("plain"));
    }
}
