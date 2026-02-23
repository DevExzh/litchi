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
}
