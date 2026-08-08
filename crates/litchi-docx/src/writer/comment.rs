/// Comment writer support for DOCX documents.
use crate::error::Result;
use chrono::DateTime;
use litchi_core::xml::escape_xml;
use std::collections::TryReserveError;
use std::fmt::Write as FmtWrite;

/// Maximum UTF-8 size accepted for a comment timestamp lexical value.
pub const MAX_DATE_BYTES: usize = 128;

/// Failure to construct a checked comment timestamp.
#[derive(Debug, thiserror::Error)]
#[non_exhaustive]
pub enum DateError {
    /// The lexical value exceeds the bounded semantic model.
    #[error("comment date length {actual} exceeds the {limit}-byte limit")]
    TooLong {
        /// Supplied UTF-8 byte length.
        actual: usize,
        /// Maximum accepted UTF-8 byte length.
        limit: usize,
    },
    /// The value contains a scalar forbidden by XML 1.0.
    #[error("comment date contains a character forbidden by XML 1.0")]
    InvalidXml,
    /// The value is not an RFC 3339 timestamp.
    #[error("comment date must be valid W3CDTF/RFC 3339")]
    InvalidLexical,
    /// Storage for an otherwise valid timestamp could not be reserved.
    #[error("allocation failed for comment date: {0}")]
    Allocation(#[source] TryReserveError),
}

/// An explicitly supplied WordprocessingML comment timestamp.
///
/// The original RFC 3339 spelling is retained so deterministic saves preserve
/// the caller's chosen offset and fractional-second precision.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Date(String);

impl Date {
    /// Parse and validate a comment timestamp without consulting an ambient
    /// clock.
    pub fn parse(value: impl AsRef<str>) -> std::result::Result<Self, DateError> {
        let value = value.as_ref();
        if value.len() > MAX_DATE_BYTES {
            return Err(DateError::TooLong {
                actual: value.len(),
                limit: MAX_DATE_BYTES,
            });
        }
        if value.chars().any(|ch| {
            !matches!(ch, '\u{9}' | '\u{A}' | '\u{D}' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}')
        }) {
            return Err(DateError::InvalidXml);
        }
        DateTime::parse_from_rfc3339(value).map_err(|_| DateError::InvalidLexical)?;

        let mut owned = String::new();
        owned
            .try_reserve_exact(value.len())
            .map_err(DateError::Allocation)?;
        owned.push_str(value);
        Ok(Self(owned))
    }

    /// Return the preserved RFC 3339 spelling.
    #[inline]
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

impl AsRef<str> for Date {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

/// A mutable comment in a Word document.
///
/// Comments are annotations attached to specific locations in the document.
#[derive(Debug, Clone)]
pub struct MutableComment {
    /// Comment ID
    id: u32,
    /// Author name
    author: String,
    /// Comment date (ISO 8601 format)
    date: Option<Date>,
    /// Comment text/content
    text: String,
    /// Initials (optional)
    initials: Option<String>,
}

impl MutableComment {
    /// Create a new comment without a timestamp.
    ///
    /// WordprocessingML makes `w:date` optional. Omitting it is the
    /// deterministic default; use [`Self::new_with_date`] when timestamp
    /// metadata is required.
    ///
    /// # Arguments
    ///
    /// * `id` - Unique comment ID
    /// * `author` - Author name
    /// * `text` - Comment text
    pub fn new(id: u32, author: String, text: String) -> Self {
        Self {
            id,
            author,
            date: None,
            text,
            initials: None,
        }
    }

    /// Create a new comment with a caller-supplied, validated timestamp.
    pub fn new_with_date(id: u32, author: String, text: String, date: Date) -> Self {
        Self {
            id,
            author,
            date: Some(date),
            text,
            initials: None,
        }
    }

    /// Get the comment ID.
    #[inline]
    pub fn id(&self) -> u32 {
        self.id
    }

    /// Get the author name.
    #[inline]
    pub fn author(&self) -> &str {
        &self.author
    }

    /// Set the author name.
    pub fn set_author(&mut self, author: String) {
        self.author = author;
    }

    /// Get the comment text.
    #[inline]
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Set the comment text.
    pub fn set_text(&mut self, text: String) {
        self.text = text;
    }

    /// Get the comment date.
    #[inline]
    pub fn date(&self) -> Option<&Date> {
        self.date.as_ref()
    }

    /// Set an explicitly constructed comment date.
    ///
    /// [`Date::parse`] validates lexical input before it can reach this
    /// mutation boundary. Passing `None` removes `w:date`.
    pub fn set_date(&mut self, date: Option<Date>) -> &mut Self {
        self.date = date;
        self
    }

    /// Get the author initials.
    #[inline]
    pub fn initials(&self) -> Option<&str> {
        self.initials.as_deref()
    }

    /// Set the author initials.
    pub fn set_initials(&mut self, initials: Option<String>) {
        self.initials = initials;
    }

    /// Generate XML for this comment.
    #[allow(dead_code)]
    pub(crate) fn to_xml(&self) -> Result<String> {
        let mut xml = String::with_capacity(256);

        write!(
            &mut xml,
            r#"<w:comment w:id="{}" w:author="{}""#,
            self.id,
            escape_xml(&self.author)
        )?;

        if let Some(date) = &self.date {
            write!(&mut xml, r#" w:date="{}""#, escape_xml(date.as_str()))?;
        }

        if let Some(initials) = &self.initials {
            write!(&mut xml, r#" w:initials="{}""#, escape_xml(initials))?;
        }

        xml.push('>');

        // Add comment content as a paragraph
        if requires_space_preservation(&self.text) {
            write!(
                &mut xml,
                r#"<w:p><w:r><w:t xml:space="preserve">{}</w:t></w:r></w:p>"#,
                escape_xml(&self.text)
            )?;
        } else {
            write!(
                &mut xml,
                "<w:p><w:r><w:t>{}</w:t></w:r></w:p>",
                escape_xml(&self.text)
            )?;
        }

        xml.push_str("</w:comment>");

        Ok(xml)
    }
}

fn requires_space_preservation(text: &str) -> bool {
    text.as_bytes()
        .first()
        .into_iter()
        .chain(text.as_bytes().last())
        .any(u8::is_ascii_whitespace)
}

#[cfg(test)]
mod tests {
    use super::{Date, DateError, MAX_DATE_BYTES, MutableComment};

    #[test]
    fn creation_is_deterministic_without_ambient_time() {
        let first = MutableComment::new(1, "John Doe".to_string(), "Test comment".to_string());
        let second = MutableComment::new(1, "John Doe".to_string(), "Test comment".to_string());
        assert_eq!(first.id(), 1);
        assert_eq!(first.author(), "John Doe");
        assert_eq!(first.text(), "Test comment");
        assert_eq!(first.date(), None);
        assert_eq!(first.to_xml().ok(), second.to_xml().ok());
    }

    #[test]
    fn explicit_date_and_xml_are_validated_and_compact() {
        let Ok(date) = Date::parse("2026-08-08T12:34:56+08:00") else {
            panic!("valid test timestamp must parse");
        };
        let mut comment = MutableComment::new_with_date(
            1,
            "Jane & Smith".to_string(),
            "Review this".to_string(),
            date,
        );
        comment.set_initials(Some("JS".to_string()));

        assert_eq!(
            comment.date().map(Date::as_str),
            Some("2026-08-08T12:34:56+08:00")
        );
        assert_eq!(
            comment.to_xml().ok().as_deref(),
            Some(
                r#"<w:comment w:id="1" w:author="Jane &amp; Smith" w:date="2026-08-08T12:34:56+08:00" w:initials="JS"><w:p><w:r><w:t>Review this</w:t></w:r></w:p></w:comment>"#
            )
        );
    }

    #[test]
    fn invalid_date_cannot_mutate_and_semantic_text_space_is_preserved() {
        let mut comment = MutableComment::new(7, "A".to_string(), " keep ".to_string());
        let Ok(original) = Date::parse("2026-08-08T04:34:56Z") else {
            panic!("valid test timestamp must parse");
        };
        comment.set_date(Some(original.clone()));
        assert!(Date::parse("not-a-date").is_err());
        assert_eq!(comment.date(), Some(&original));
        comment.set_date(None);
        assert_eq!(comment.date(), None);
        assert_eq!(
            comment.to_xml().ok().as_deref(),
            Some(
                r#"<w:comment w:id="7" w:author="A"><w:p><w:r><w:t xml:space="preserve"> keep </w:t></w:r></w:p></w:comment>"#
            )
        );
    }

    #[test]
    fn oversized_date_is_rejected_before_owned_storage() {
        let oversized = "2".repeat(MAX_DATE_BYTES + 1);
        assert!(matches!(
            Date::parse(oversized.as_str()),
            Err(DateError::TooLong {
                actual,
                limit: MAX_DATE_BYTES,
            }) if actual == MAX_DATE_BYTES + 1
        ));
    }
}
