//! Hyperlinks writer for DOC files
//!
//! Generates HYPERLINK field codes for embedding links in documents.

/// Hyperlink type
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HyperlinkType {
    /// URL hyperlink
    Url,
    /// Email hyperlink
    Email,
    /// File hyperlink
    File,
    /// Bookmark hyperlink
    Bookmark,
}

/// A hyperlink entry
#[derive(Debug, Clone)]
pub struct HyperlinkEntry {
    /// Character position where hyperlink starts
    pub start_position: u32,
    /// Destination URL/path
    pub destination: String,
    /// Display text
    pub display_text: String,
    /// Link type
    pub link_type: HyperlinkType,
}

impl HyperlinkEntry {
    /// Create a new URL hyperlink
    pub fn url(
        start_position: u32,
        url: impl Into<String>,
        display_text: impl Into<String>,
    ) -> Self {
        Self {
            start_position,
            destination: url.into(),
            display_text: display_text.into(),
            link_type: HyperlinkType::Url,
        }
    }

    /// Create a new email hyperlink
    pub fn email(
        start_position: u32,
        email: impl Into<String>,
        display_text: impl Into<String>,
    ) -> Self {
        Self {
            start_position,
            destination: email.into(),
            display_text: display_text.into(),
            link_type: HyperlinkType::Email,
        }
    }

    /// Create a new bookmark hyperlink
    pub fn bookmark(
        start_position: u32,
        bookmark: impl Into<String>,
        display_text: impl Into<String>,
    ) -> Self {
        Self {
            start_position,
            destination: bookmark.into(),
            display_text: display_text.into(),
            link_type: HyperlinkType::Bookmark,
        }
    }

    /// Generate HYPERLINK field code
    pub fn to_field_code(&self) -> String {
        match self.link_type {
            HyperlinkType::Url => format!("HYPERLINK \"{}\"", self.destination),
            HyperlinkType::Email => format!("HYPERLINK \"mailto:{}\"", self.destination),
            HyperlinkType::File => format!("HYPERLINK \"{}\"", self.destination),
            HyperlinkType::Bookmark => format!("HYPERLINK \\l \"{}\"", self.destination),
        }
    }
}

/// Hyperlinks writer
#[derive(Debug)]
pub struct HyperlinksWriter {
    hyperlinks: Vec<HyperlinkEntry>,
}

impl HyperlinksWriter {
    /// Create a new hyperlinks writer
    pub fn new() -> Self {
        Self {
            hyperlinks: Vec::new(),
        }
    }

    /// Add a hyperlink
    pub fn add_hyperlink(&mut self, hyperlink: HyperlinkEntry) {
        self.hyperlinks.push(hyperlink);
    }

    /// Get all hyperlinks
    pub fn hyperlinks(&self) -> &[HyperlinkEntry] {
        &self.hyperlinks
    }

    /// Check if empty
    pub fn is_empty(&self) -> bool {
        self.hyperlinks.is_empty()
    }

    /// Get hyperlinks sorted by position
    pub fn hyperlinks_sorted(&self) -> Vec<&HyperlinkEntry> {
        let mut sorted: Vec<_> = self.hyperlinks.iter().collect();
        sorted.sort_by_key(|h| h.start_position);
        sorted
    }
}

impl Default for HyperlinksWriter {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_hyperlink_type_variants() {
        // Test all hyperlink types are distinct
        let types = vec![
            HyperlinkType::Url,
            HyperlinkType::Email,
            HyperlinkType::File,
            HyperlinkType::Bookmark,
        ];
        assert_eq!(types.len(), 4);
    }

    #[test]
    fn test_hyperlink_entry_url() {
        let entry = HyperlinkEntry::url(100, "https://example.com", "Click here");
        assert_eq!(entry.start_position, 100);
        assert_eq!(entry.destination, "https://example.com");
        assert_eq!(entry.display_text, "Click here");
        assert_eq!(entry.link_type, HyperlinkType::Url);
    }

    #[test]
    fn test_hyperlink_entry_email() {
        let entry = HyperlinkEntry::email(200, "test@example.com", "Email me");
        assert_eq!(entry.start_position, 200);
        assert_eq!(entry.destination, "test@example.com");
        assert_eq!(entry.display_text, "Email me");
        assert_eq!(entry.link_type, HyperlinkType::Email);
    }

    #[test]
    fn test_hyperlink_entry_bookmark() {
        let entry = HyperlinkEntry::bookmark(50, "Section1", "Go to Section 1");
        assert_eq!(entry.start_position, 50);
        assert_eq!(entry.destination, "Section1");
        assert_eq!(entry.display_text, "Go to Section 1");
        assert_eq!(entry.link_type, HyperlinkType::Bookmark);
    }

    #[test]
    fn test_to_field_code_url() {
        let entry = HyperlinkEntry::url(0, "https://example.com", "Example");
        assert_eq!(entry.to_field_code(), "HYPERLINK \"https://example.com\"");
    }

    #[test]
    fn test_to_field_code_email() {
        let entry = HyperlinkEntry::email(0, "user@test.com", "Email");
        assert_eq!(entry.to_field_code(), "HYPERLINK \"mailto:user@test.com\"");
    }

    #[test]
    fn test_to_field_code_bookmark() {
        let entry = HyperlinkEntry::bookmark(0, "Introduction", "Intro");
        assert_eq!(entry.to_field_code(), "HYPERLINK \\l \"Introduction\"");
    }

    #[test]
    fn test_hyperlinks_writer_new() {
        let writer = HyperlinksWriter::new();
        assert!(writer.is_empty());
        assert!(writer.hyperlinks().is_empty());
    }

    #[test]
    fn test_hyperlinks_writer_default() {
        let writer: HyperlinksWriter = Default::default();
        assert!(writer.is_empty());
    }

    #[test]
    fn test_add_hyperlink() {
        let mut writer = HyperlinksWriter::new();
        let link = HyperlinkEntry::url(100, "https://example.com", "Example");
        writer.add_hyperlink(link);
        assert!(!writer.is_empty());
        assert_eq!(writer.hyperlinks().len(), 1);
    }

    #[test]
    fn test_add_multiple_hyperlinks() {
        let mut writer = HyperlinksWriter::new();
        writer.add_hyperlink(HyperlinkEntry::url(300, "https://a.com", "A"));
        writer.add_hyperlink(HyperlinkEntry::url(100, "https://b.com", "B"));
        writer.add_hyperlink(HyperlinkEntry::url(200, "https://c.com", "C"));
        assert_eq!(writer.hyperlinks().len(), 3);
    }

    #[test]
    fn test_hyperlinks_sorted() {
        let mut writer = HyperlinksWriter::new();
        writer.add_hyperlink(HyperlinkEntry::url(300, "https://a.com", "A"));
        writer.add_hyperlink(HyperlinkEntry::url(100, "https://b.com", "B"));
        writer.add_hyperlink(HyperlinkEntry::url(200, "https://c.com", "C"));

        let sorted = writer.hyperlinks_sorted();
        assert_eq!(sorted.len(), 3);
        assert_eq!(sorted[0].start_position, 100);
        assert_eq!(sorted[1].start_position, 200);
        assert_eq!(sorted[2].start_position, 300);
    }

    #[test]
    fn test_hyperlinks_sorted_empty() {
        let writer = HyperlinksWriter::new();
        let sorted = writer.hyperlinks_sorted();
        assert!(sorted.is_empty());
    }

    #[test]
    fn test_hyperlink_entry_clone() {
        let entry = HyperlinkEntry::url(100, "https://example.com", "Example");
        let cloned = entry.clone();
        assert_eq!(entry.start_position, cloned.start_position);
        assert_eq!(entry.destination, cloned.destination);
        assert_eq!(entry.display_text, cloned.display_text);
        assert_eq!(entry.link_type, cloned.link_type);
    }

    #[test]
    fn test_hyperlink_entry_debug() {
        let entry = HyperlinkEntry::url(100, "https://example.com", "Example");
        let debug_str = format!("{:?}", entry);
        assert!(debug_str.contains("HyperlinkEntry"));
    }

    #[test]
    fn test_hyperlinks_writer_debug() {
        let writer = HyperlinksWriter::new();
        let debug_str = format!("{:?}", writer);
        assert!(debug_str.contains("HyperlinksWriter"));
    }

    #[test]
    fn test_field_code_with_special_chars() {
        let entry = HyperlinkEntry::url(0, "https://example.com?foo=bar&baz=qux", "Link");
        let code = entry.to_field_code();
        assert!(code.contains("https://example.com?foo=bar&baz=qux"));
    }

    #[test]
    fn test_all_hyperlink_types_in_writer() {
        let mut writer = HyperlinksWriter::new();
        writer.add_hyperlink(HyperlinkEntry::url(0, "https://example.com", "URL"));
        writer.add_hyperlink(HyperlinkEntry::email(0, "test@example.com", "Email"));
        writer.add_hyperlink(HyperlinkEntry::bookmark(0, "Section1", "Bookmark"));
        assert_eq!(writer.hyperlinks().len(), 3);
    }
}
