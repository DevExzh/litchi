/// Bookmark writer support for DOCX documents.
use crate::ooxml::error::Result;
use std::fmt::Write as FmtWrite;

/// A mutable bookmark in a Word document.
///
/// Bookmarks mark named locations in a document for quick navigation and cross-referencing.
#[derive(Debug, Clone)]
pub struct MutableBookmark {
    /// Bookmark ID
    id: u32,
    /// Bookmark name
    name: String,
}

impl MutableBookmark {
    /// Create a new bookmark.
    ///
    /// # Arguments
    ///
    /// * `id` - Unique bookmark ID
    /// * `name` - Bookmark name (must not start with underscore for user bookmarks)
    pub fn new(id: u32, name: String) -> Self {
        Self { id, name }
    }

    /// Get the bookmark ID.
    #[inline]
    pub fn id(&self) -> u32 {
        self.id
    }

    /// Get the bookmark name.
    #[inline]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Set the bookmark name.
    pub fn set_name(&mut self, name: String) {
        self.name = name;
    }

    /// Generate XML for bookmark start tag.
    #[allow(dead_code)]
    pub(crate) fn to_xml_start(&self) -> Result<String> {
        let mut xml = String::with_capacity(128);
        write!(
            &mut xml,
            r#"<w:bookmarkStart w:id="{}" w:name="{}"/>"#,
            self.id,
            escape_xml(&self.name)
        )?;
        Ok(xml)
    }

    /// Generate XML for bookmark end tag.
    #[allow(dead_code)]
    pub(crate) fn to_xml_end(&self) -> Result<String> {
        let mut xml = String::with_capacity(64);
        write!(&mut xml, r#"<w:bookmarkEnd w:id="{}"/>"#, self.id)?;
        Ok(xml)
    }
}

/// Escape XML special characters.
#[allow(dead_code)]
fn escape_xml(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_bookmark_creation() {
        let bookmark = MutableBookmark::new(1, "Section1".to_string());
        assert_eq!(bookmark.id(), 1);
        assert_eq!(bookmark.name(), "Section1");
    }

    #[test]
    fn test_bookmark_xml() {
        let bookmark = MutableBookmark::new(42, "MyBookmark".to_string());

        let start_xml = bookmark.to_xml_start().unwrap();
        assert!(start_xml.contains(r#"w:id="42""#));
        assert!(start_xml.contains(r#"w:name="MyBookmark""#));
        assert!(start_xml.contains("<w:bookmarkStart"));

        let end_xml = bookmark.to_xml_end().unwrap();
        assert!(end_xml.contains(r#"w:id="42""#));
        assert!(end_xml.contains("<w:bookmarkEnd"));
    }
}
