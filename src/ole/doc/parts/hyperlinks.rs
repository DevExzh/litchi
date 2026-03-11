/// Hyperlinks parser for Word binary format.
///
/// Based on Apache POI's implementation and the MS-DOC specification.
/// Hyperlinks in DOC files are implemented as HYPERLINK fields with destinations
/// stored in various formats (URLs, file paths, bookmarks, etc.).
use super::super::package::Result;
use super::fields::{FieldType, FieldsTable};

/// Hyperlink destination type
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum HyperlinkType {
    /// URL (http://, https://, ftp://, etc.)
    Url,
    /// File path
    File,
    /// Bookmark/anchor within document
    Bookmark,
    /// Email address (mailto:)
    Email,
    /// Unknown/other
    Unknown,
}

impl HyperlinkType {
    /// Determine hyperlink type from destination string
    pub fn from_destination(dest: &str) -> Self {
        let dest_lower = dest.to_lowercase();
        if dest_lower.starts_with("http://")
            || dest_lower.starts_with("https://")
            || dest_lower.starts_with("ftp://")
        {
            HyperlinkType::Url
        } else if dest_lower.starts_with("mailto:") {
            HyperlinkType::Email
        } else if dest_lower.starts_with("file://")
            || dest_lower.contains(":\\")
            || dest_lower.starts_with("\\\\")
        {
            HyperlinkType::File
        } else if dest_lower.starts_with("#") {
            HyperlinkType::Bookmark
        } else {
            HyperlinkType::Unknown
        }
    }
}

/// A hyperlink in the document
#[derive(Debug, Clone)]
pub struct Hyperlink {
    /// Character position range in the main document
    pub start_cp: u32,
    pub end_cp: u32,
    /// The raw field code (e.g., "HYPERLINK \"http://example.com\"")
    pub field_code: String,
    /// The destination (extracted from field code)
    pub destination: String,
    /// Display text (the visible text in the document)
    pub display_text: String,
    /// Hyperlink type
    pub link_type: HyperlinkType,
}

impl Hyperlink {
    /// Create a new hyperlink
    pub fn new(start_cp: u32, end_cp: u32, field_code: String, display_text: String) -> Self {
        let destination = Self::extract_destination(&field_code);
        let link_type = HyperlinkType::from_destination(&destination);

        Self {
            start_cp,
            end_cp,
            field_code,
            destination,
            display_text,
            link_type,
        }
    }

    /// Extract the destination URL/path from a HYPERLINK field code
    ///
    /// Examples:
    /// - `HYPERLINK "http://example.com"` -> `http://example.com`
    /// - `HYPERLINK "http://example.com" \o "Tooltip"` -> `http://example.com`
    /// - `HYPERLINK \l "bookmark"` -> `#bookmark`
    fn extract_destination(field_code: &str) -> String {
        let code = field_code.trim();

        // Remove "HYPERLINK" prefix
        let code = if let Some(stripped) = code.strip_prefix("HYPERLINK") {
            stripped.trim()
        } else {
            code
        };

        // Check for bookmark flag (\l)
        if code.starts_with("\\l") {
            // Extract bookmark name
            let bookmark_part = code.strip_prefix("\\l").unwrap_or("").trim();
            if let Some(dest) = Self::extract_quoted_string(bookmark_part) {
                return format!("#{}", dest);
            }
        }

        // Extract the first quoted string as the destination
        if let Some(dest) = Self::extract_quoted_string(code) {
            dest
        } else {
            // Fallback: take the first token
            code.split_whitespace().next().unwrap_or("").to_string()
        }
    }

    /// Extract a quoted string from text
    fn extract_quoted_string(text: &str) -> Option<String> {
        let text = text.trim();

        // Handle double quotes
        if let Some(start) = text.find('"') {
            let after_start = &text[start + 1..];
            if let Some(end) = after_start.find('"') {
                return Some(after_start[..end].to_string());
            }
        }

        // Handle single quotes (less common but valid)
        if let Some(start) = text.find('\'') {
            let after_start = &text[start + 1..];
            if let Some(end) = after_start.find('\'') {
                return Some(after_start[..end].to_string());
            }
        }

        None
    }

    /// Get the length of the hyperlink text
    pub fn length(&self) -> u32 {
        self.end_cp.saturating_sub(self.start_cp)
    }
}

/// Hyperlinks table parser
pub struct HyperlinksTable {
    /// All hyperlinks in the document
    hyperlinks: Vec<Hyperlink>,
}

impl HyperlinksTable {
    /// Parse hyperlinks from the fields table and document text
    ///
    /// # Arguments
    ///
    /// * `fields_table` - The parsed fields table
    /// * `text_extractor` - Function to extract text from character positions
    ///
    /// # Returns
    ///
    /// A parsed HyperlinksTable
    pub fn from_fields<F>(fields_table: &FieldsTable, text_extractor: F) -> Result<Self>
    where
        F: Fn(u32, u32) -> Result<String>,
    {
        let mut hyperlinks = Vec::new();

        for field in fields_table.main_document_fields() {
            // Only process HYPERLINK fields
            if field.field_type != FieldType::Hyperlink {
                continue;
            }

            // Extract field code (between begin and separator/end)
            let (code_start, code_end) = field.code_range();
            let field_code = text_extractor(code_start, code_end)
                .unwrap_or_default()
                .trim()
                .to_string();

            // Extract display text (between separator and end, if separator exists)
            let display_text = if let Some((result_start, result_end)) = field.result_range() {
                text_extractor(result_start, result_end).unwrap_or_default()
            } else {
                // No separator means field code is visible
                field_code.clone()
            };

            hyperlinks.push(Hyperlink::new(
                field.start_cp,
                field.end_cp,
                field_code,
                display_text,
            ));
        }

        Ok(Self { hyperlinks })
    }

    /// Get all hyperlinks
    pub fn hyperlinks(&self) -> &[Hyperlink] {
        &self.hyperlinks
    }

    /// Find hyperlinks that overlap with a character position
    pub fn find_at_position(&self, cp: u32) -> Vec<&Hyperlink> {
        self.hyperlinks
            .iter()
            .filter(|h| h.start_cp <= cp && cp < h.end_cp)
            .collect()
    }

    /// Get hyperlinks by type
    pub fn by_type(&self, link_type: HyperlinkType) -> Vec<&Hyperlink> {
        self.hyperlinks
            .iter()
            .filter(|h| h.link_type == link_type)
            .collect()
    }

    /// Get the count of hyperlinks
    pub fn count(&self) -> usize {
        self.hyperlinks.len()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_extract_destination() {
        assert_eq!(
            Hyperlink::extract_destination("HYPERLINK \"http://example.com\""),
            "http://example.com"
        );

        assert_eq!(
            Hyperlink::extract_destination("HYPERLINK \"http://example.com\" \\o \"Tooltip\""),
            "http://example.com"
        );

        assert_eq!(
            Hyperlink::extract_destination("HYPERLINK \\l \"bookmark1\""),
            "#bookmark1"
        );
    }

    #[test]
    fn test_hyperlink_type() {
        assert_eq!(
            HyperlinkType::from_destination("http://example.com"),
            HyperlinkType::Url
        );
        assert_eq!(
            HyperlinkType::from_destination("https://example.com"),
            HyperlinkType::Url
        );
        assert_eq!(
            HyperlinkType::from_destination("mailto:test@example.com"),
            HyperlinkType::Email
        );
        assert_eq!(
            HyperlinkType::from_destination("file://C:\\path\\to\\file"),
            HyperlinkType::File
        );
        assert_eq!(
            HyperlinkType::from_destination("#bookmark"),
            HyperlinkType::Bookmark
        );
    }

    #[test]
    fn test_hyperlink_creation() {
        let link = Hyperlink::new(
            100,
            200,
            "HYPERLINK \"http://example.com\"".to_string(),
            "Example Site".to_string(),
        );

        assert_eq!(link.destination, "http://example.com");
        assert_eq!(link.display_text, "Example Site");
        assert_eq!(link.link_type, HyperlinkType::Url);
        assert_eq!(link.length(), 100);
    }

    #[test]
    fn test_hyperlink_type_url_variations() {
        assert_eq!(
            HyperlinkType::from_destination("http://example.com/path"),
            HyperlinkType::Url
        );
        assert_eq!(
            HyperlinkType::from_destination("https://example.com/path?query=1"),
            HyperlinkType::Url
        );
        assert_eq!(
            HyperlinkType::from_destination("ftp://ftp.example.com/file.txt"),
            HyperlinkType::Url
        );
        assert_eq!(
            HyperlinkType::from_destination("HTTP://EXAMPLE.COM"),
            HyperlinkType::Url
        );
        assert_eq!(
            HyperlinkType::from_destination("HTTPS://EXAMPLE.COM"),
            HyperlinkType::Url
        );
    }

    #[test]
    fn test_hyperlink_type_email_variations() {
        assert_eq!(
            HyperlinkType::from_destination("mailto:test@example.com"),
            HyperlinkType::Email
        );
        assert_eq!(
            HyperlinkType::from_destination("MAILTO:Test@Example.COM"),
            HyperlinkType::Email
        );
        assert_eq!(
            HyperlinkType::from_destination("mailto:user+tag@example.com?subject=Hello"),
            HyperlinkType::Email
        );
    }

    #[test]
    fn test_hyperlink_type_file_variations() {
        assert_eq!(
            HyperlinkType::from_destination("file://C:\\path\\to\\file.txt"),
            HyperlinkType::File
        );
        assert_eq!(
            HyperlinkType::from_destination("C:\\path\\to\\file.txt"),
            HyperlinkType::File
        );
        assert_eq!(
            HyperlinkType::from_destination("\\\\server\\share\\file.txt"),
            HyperlinkType::File
        );
    }

    #[test]
    fn test_hyperlink_type_bookmark_variations() {
        assert_eq!(
            HyperlinkType::from_destination("#bookmark"),
            HyperlinkType::Bookmark
        );
        assert_eq!(
            HyperlinkType::from_destination("#Section1"),
            HyperlinkType::Bookmark
        );
        assert_eq!(
            HyperlinkType::from_destination("#_Toc123"),
            HyperlinkType::Bookmark
        );
    }

    #[test]
    fn test_hyperlink_type_unknown() {
        assert_eq!(
            HyperlinkType::from_destination("unknown://example.com"),
            HyperlinkType::Unknown
        );
        assert_eq!(
            HyperlinkType::from_destination("just-some-text"),
            HyperlinkType::Unknown
        );
        assert_eq!(HyperlinkType::from_destination(""), HyperlinkType::Unknown);
    }

    #[test]
    fn test_hyperlink_extract_destination_single_quotes() {
        // Single quotes (less common but valid)
        assert_eq!(
            Hyperlink::extract_destination("HYPERLINK 'http://example.com'"),
            "http://example.com"
        );
    }

    #[test]
    fn test_hyperlink_extract_destination_no_quotes() {
        // Fallback: take first token when no quotes
        assert_eq!(
            Hyperlink::extract_destination("HYPERLINK http://example.com"),
            "http://example.com"
        );
    }

    #[test]
    fn test_hyperlink_extract_destination_no_hyperlink_prefix() {
        // Without HYPERLINK prefix
        assert_eq!(
            Hyperlink::extract_destination("\"http://example.com\""),
            "http://example.com"
        );
    }

    #[test]
    fn test_hyperlink_extract_destination_empty() {
        assert_eq!(Hyperlink::extract_destination("HYPERLINK"), "");
        assert_eq!(Hyperlink::extract_destination(""), "");
    }

    #[test]
    fn test_hyperlink_new_https() {
        let link = Hyperlink::new(
            0,
            100,
            "HYPERLINK \"https://secure.example.com/path\"".to_string(),
            "Secure Link".to_string(),
        );

        assert_eq!(link.start_cp, 0);
        assert_eq!(link.end_cp, 100);
        assert_eq!(link.destination, "https://secure.example.com/path");
        assert_eq!(link.display_text, "Secure Link");
        assert_eq!(link.link_type, HyperlinkType::Url);
        assert_eq!(
            link.field_code,
            "HYPERLINK \"https://secure.example.com/path\""
        );
    }

    #[test]
    fn test_hyperlink_new_mailto() {
        let link = Hyperlink::new(
            50,
            150,
            "HYPERLINK \"mailto:user@example.com\"".to_string(),
            "Send Email".to_string(),
        );

        // The destination includes the mailto: prefix (extracted from quoted string)
        assert_eq!(link.destination, "mailto:user@example.com");
        assert_eq!(link.display_text, "Send Email");
        assert_eq!(link.link_type, HyperlinkType::Email);
        assert_eq!(link.length(), 100);
    }

    #[test]
    fn test_hyperlink_new_bookmark() {
        let link = Hyperlink::new(
            200,
            250,
            "HYPERLINK \\l \"Introduction\"".to_string(),
            "Go to Intro".to_string(),
        );

        assert_eq!(link.destination, "#Introduction");
        assert_eq!(link.display_text, "Go to Intro");
        assert_eq!(link.link_type, HyperlinkType::Bookmark);
        assert_eq!(link.length(), 50);
    }

    #[test]
    fn test_hyperlink_new_file() {
        let link = Hyperlink::new(
            100,
            200,
            "HYPERLINK \"C:\\Users\\test\\document.doc\"".to_string(),
            "Open Document".to_string(),
        );

        assert_eq!(link.destination, "C:\\Users\\test\\document.doc");
        assert_eq!(link.link_type, HyperlinkType::File);
    }

    #[test]
    fn test_hyperlink_length_zero() {
        let link = Hyperlink::new(
            100,
            100,
            "HYPERLINK \"http://example.com\"".to_string(),
            "Zero Length".to_string(),
        );

        assert_eq!(link.length(), 0);
    }

    #[test]
    fn test_hyperlink_length_large() {
        let link = Hyperlink::new(
            0,
            u32::MAX,
            "HYPERLINK \"http://example.com\"".to_string(),
            "Large Length".to_string(),
        );

        assert_eq!(link.length(), u32::MAX);
    }

    #[test]
    fn test_hyperlink_clone() {
        let link = Hyperlink::new(
            100,
            200,
            "HYPERLINK \"http://example.com\"".to_string(),
            "Clone Test".to_string(),
        );
        let cloned = link.clone();

        assert_eq!(cloned.start_cp, link.start_cp);
        assert_eq!(cloned.end_cp, link.end_cp);
        assert_eq!(cloned.destination, link.destination);
        assert_eq!(cloned.display_text, link.display_text);
        assert_eq!(cloned.link_type, link.link_type);
        assert_eq!(cloned.field_code, link.field_code);
    }

    #[test]
    fn test_hyperlink_type_clone() {
        let url_type = HyperlinkType::Url;
        let cloned = url_type.clone();
        assert_eq!(url_type, cloned);
    }

    #[test]
    fn test_hyperlink_type_equality() {
        assert_eq!(HyperlinkType::Url, HyperlinkType::Url);
        assert_ne!(HyperlinkType::Url, HyperlinkType::Email);
        assert_ne!(HyperlinkType::File, HyperlinkType::Bookmark);
    }

    #[test]
    fn test_hyperlink_debug() {
        let link = Hyperlink::new(
            100,
            200,
            "HYPERLINK \"http://example.com\"".to_string(),
            "Debug Test".to_string(),
        );
        let debug_str = format!("{:?}", link);

        assert!(debug_str.contains("Hyperlink"));
        assert!(debug_str.contains("http://example.com"));
        assert!(debug_str.contains("Debug Test"));
        assert!(debug_str.contains("100"));
        assert!(debug_str.contains("200"));
    }

    #[test]
    fn test_hyperlinks_table_empty() {
        let table = HyperlinksTable { hyperlinks: vec![] };

        assert_eq!(table.count(), 0);
        assert!(table.hyperlinks().is_empty());
    }

    #[test]
    fn test_hyperlinks_table_single() {
        let link = Hyperlink::new(
            100,
            200,
            "HYPERLINK \"http://example.com\"".to_string(),
            "Example".to_string(),
        );
        let table = HyperlinksTable {
            hyperlinks: vec![link],
        };

        assert_eq!(table.count(), 1);
        assert_eq!(table.hyperlinks().len(), 1);
        assert_eq!(table.hyperlinks()[0].destination, "http://example.com");
    }

    #[test]
    fn test_hyperlinks_table_multiple() {
        let links = vec![
            Hyperlink::new(
                0,
                50,
                "HYPERLINK \"http://first.com\"".to_string(),
                "First".to_string(),
            ),
            Hyperlink::new(
                100,
                150,
                "HYPERLINK \"http://second.com\"".to_string(),
                "Second".to_string(),
            ),
            Hyperlink::new(
                200,
                250,
                "HYPERLINK \"mailto:test@example.com\"".to_string(),
                "Email".to_string(),
            ),
        ];
        let table = HyperlinksTable { hyperlinks: links };

        assert_eq!(table.count(), 3);
        assert_eq!(table.hyperlinks().len(), 3);
    }

    #[test]
    fn test_hyperlinks_table_find_at_position() {
        let links = vec![
            Hyperlink::new(
                0,
                100,
                "HYPERLINK \"http://first.com\"".to_string(),
                "First".to_string(),
            ),
            Hyperlink::new(
                200,
                300,
                "HYPERLINK \"http://second.com\"".to_string(),
                "Second".to_string(),
            ),
        ];
        let table = HyperlinksTable { hyperlinks: links };

        // Position within first link
        let found = table.find_at_position(50);
        assert_eq!(found.len(), 1);
        assert_eq!(found[0].display_text, "First");

        // Position within second link
        let found = table.find_at_position(250);
        assert_eq!(found.len(), 1);
        assert_eq!(found[0].display_text, "Second");

        // Position before all links
        let found = table.find_at_position(0);
        assert_eq!(found.len(), 1);

        // Position between links
        let found = table.find_at_position(150);
        assert!(found.is_empty());

        // Position after all links
        let found = table.find_at_position(400);
        assert!(found.is_empty());
    }

    #[test]
    fn test_hyperlinks_table_find_at_position_overlapping() {
        // Links at exact boundaries
        let links = vec![
            Hyperlink::new(
                0,
                10,
                "HYPERLINK \"http://first.com\"".to_string(),
                "First".to_string(),
            ),
            Hyperlink::new(
                10,
                20,
                "HYPERLINK \"http://second.com\"".to_string(),
                "Second".to_string(),
            ),
        ];
        let table = HyperlinksTable { hyperlinks: links };

        // Position at boundary (should match first, not second, due to < end_cp)
        let found = table.find_at_position(10);
        assert_eq!(found.len(), 1);
        assert_eq!(found[0].display_text, "Second");

        // Position at start of first
        let found = table.find_at_position(0);
        assert_eq!(found.len(), 1);
        assert_eq!(found[0].display_text, "First");
    }

    #[test]
    fn test_hyperlinks_table_by_type() {
        let links = vec![
            Hyperlink::new(
                0,
                50,
                "HYPERLINK \"http://example.com\"".to_string(),
                "URL".to_string(),
            ),
            Hyperlink::new(
                100,
                150,
                "HYPERLINK \"mailto:test@example.com\"".to_string(),
                "Email".to_string(),
            ),
            Hyperlink::new(
                200,
                250,
                "HYPERLINK \"https://another.com\"".to_string(),
                "Another URL".to_string(),
            ),
            Hyperlink::new(
                300,
                350,
                "HYPERLINK \\l \"bookmark\"".to_string(),
                "Bookmark".to_string(),
            ),
        ];
        let table = HyperlinksTable { hyperlinks: links };

        let urls = table.by_type(HyperlinkType::Url);
        assert_eq!(urls.len(), 2);
        assert_eq!(urls[0].display_text, "URL");
        assert_eq!(urls[1].display_text, "Another URL");

        let emails = table.by_type(HyperlinkType::Email);
        assert_eq!(emails.len(), 1);
        assert_eq!(emails[0].display_text, "Email");

        let bookmarks = table.by_type(HyperlinkType::Bookmark);
        assert_eq!(bookmarks.len(), 1);
        assert_eq!(bookmarks[0].display_text, "Bookmark");

        let files = table.by_type(HyperlinkType::File);
        assert!(files.is_empty());
    }

    #[test]
    fn test_hyperlink_unicode() {
        let link = Hyperlink::new(
            100,
            200,
            "HYPERLINK \"https://例子.com/路径\"".to_string(),
            "Unicode: 你好世界 🎉".to_string(),
        );

        assert_eq!(link.destination, "https://例子.com/路径");
        assert_eq!(link.display_text, "Unicode: 你好世界 🎉");
    }

    #[test]
    fn test_hyperlink_extract_destination_with_special_chars() {
        assert_eq!(
            Hyperlink::extract_destination("HYPERLINK \"http://example.com/path?query=1&other=2\""),
            "http://example.com/path?query=1&other=2"
        );
        assert_eq!(
            Hyperlink::extract_destination("HYPERLINK \"http://example.com/path#fragment\""),
            "http://example.com/path#fragment"
        );
    }
}
