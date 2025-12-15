use crate::common::xml::escape_xml;
/// Comment writer support for DOCX documents.
use crate::ooxml::error::Result;
use std::fmt::Write as FmtWrite;

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
    date: Option<String>,
    /// Comment text/content
    text: String,
    /// Initials (optional)
    initials: Option<String>,
}

impl MutableComment {
    /// Create a new comment.
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
            date: Some(chrono::Utc::now().to_rfc3339()),
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
    pub fn date(&self) -> Option<&str> {
        self.date.as_deref()
    }

    /// Set the comment date (ISO 8601 format).
    pub fn set_date(&mut self, date: Option<String>) {
        self.date = date;
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
            r#"<w:comment w:id="{}" w:author="{}"#,
            self.id,
            escape_xml(&self.author)
        )?;

        if let Some(date) = &self.date {
            write!(&mut xml, r#" w:date="{}""#, escape_xml(date))?;
        }

        if let Some(initials) = &self.initials {
            write!(&mut xml, r#" w:initials="{}""#, escape_xml(initials))?;
        }

        xml.push('>');

        // Add comment content as a paragraph
        write!(
            &mut xml,
            "<w:p><w:r><w:t>{}</w:t></w:r></w:p>",
            escape_xml(&self.text)
        )?;

        xml.push_str("</w:comment>");

        Ok(xml)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_comment_creation() {
        let comment = MutableComment::new(1, "John Doe".to_string(), "Test comment".to_string());
        assert_eq!(comment.id(), 1);
        assert_eq!(comment.author(), "John Doe");
        assert_eq!(comment.text(), "Test comment");
        assert!(comment.date().is_some());
    }

    #[test]
    fn test_comment_xml() {
        let mut comment =
            MutableComment::new(1, "Jane Smith".to_string(), "Review this".to_string());
        comment.set_initials(Some("JS".to_string()));

        let xml = comment.to_xml().unwrap();
        assert!(xml.contains(r#"w:id="1""#));
        // Author name is XML-escaped, so we check for the presence without quotes
        assert!(xml.contains("w:author="));
        assert!(xml.contains("Jane Smith"));
        assert!(xml.contains(r#"w:initials="JS""#));
        assert!(xml.contains("Review this"));
    }
}
