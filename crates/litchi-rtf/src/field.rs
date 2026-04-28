//! RTF field support (hyperlinks, cross-references, etc.).
//!
//! RTF fields are structured as:
//! {\field{\*\fldinst FIELD_INSTRUCTION}{\fldrslt FIELD_RESULT}}
//!
//! Common field types:
//! - HYPERLINK: External and internal hyperlinks
//! - REF: Cross-references
//! - PAGE: Page numbers
//! - DATE: Date/time
//! - TOC: Table of contents

use std::borrow::Cow;

/// Field type in RTF documents.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FieldType {
    /// Hyperlink field
    Hyperlink,
    /// Cross-reference
    Reference,
    /// Page number
    Page,
    /// Date/time
    Date,
    /// Table of contents
    Toc,
    /// Bookmark
    Bookmark,
    /// Equation/formula
    Equation,
    /// Index entry
    Index,
    /// Unknown or custom field
    Unknown,
}

/// Parsed RTF field.
#[derive(Debug, Clone)]
pub struct Field<'a> {
    /// Field type
    pub field_type: FieldType,
    /// Field instruction (the command)
    pub instruction: Cow<'a, str>,
    /// Field result (the displayed text)
    pub result: Cow<'a, str>,
}

impl<'a> Field<'a> {
    /// Create a new field.
    #[inline]
    pub fn new(field_type: FieldType, instruction: Cow<'a, str>, result: Cow<'a, str>) -> Self {
        Self {
            field_type,
            instruction,
            result,
        }
    }

    /// Parse a field instruction to extract the type and parameters.
    ///
    /// # Arguments
    ///
    /// * `instruction` - Raw field instruction string
    ///
    /// # Returns
    ///
    /// Parsed field with type and parameters
    pub fn parse_instruction(instruction: &'a str) -> Self {
        let trimmed = instruction.trim();

        // Extract field type (first word)
        let field_type = if trimmed.starts_with("HYPERLINK") {
            FieldType::Hyperlink
        } else if trimmed.starts_with("REF") {
            FieldType::Reference
        } else if trimmed.starts_with("PAGE") {
            FieldType::Page
        } else if trimmed.starts_with("DATE") || trimmed.starts_with("TIME") {
            FieldType::Date
        } else if trimmed.starts_with("TOC") {
            FieldType::Toc
        } else if trimmed.starts_with("BOOKMARK") {
            FieldType::Bookmark
        } else if trimmed.starts_with("EQ") {
            FieldType::Equation
        } else if trimmed.starts_with("INDEX") || trimmed.starts_with("XE") {
            FieldType::Index
        } else {
            FieldType::Unknown
        };

        Self {
            field_type,
            instruction: Cow::Borrowed(instruction),
            result: Cow::Borrowed(""),
        }
    }

    /// Extract URL from HYPERLINK field instruction.
    ///
    /// HYPERLINK fields have format: HYPERLINK "url" \o "tooltip"
    ///
    /// # Returns
    ///
    /// URL if this is a hyperlink field, None otherwise
    pub fn extract_url(&self) -> Option<String> {
        if self.field_type != FieldType::Hyperlink {
            return None;
        }

        let inst = self.instruction.trim();
        if !inst.starts_with("HYPERLINK") {
            return None;
        }

        // Find URL in quotes
        let after_hyperlink = &inst[9..].trim_start();

        // Handle quoted URL
        if let Some(start_quote) = after_hyperlink.find('"')
            && let Some(end_quote) = after_hyperlink[start_quote + 1..].find('"')
        {
            let url = &after_hyperlink[start_quote + 1..start_quote + 1 + end_quote];
            return Some(url.to_string());
        }

        // Handle unquoted URL (space-delimited)
        let parts: Vec<&str> = after_hyperlink.split_whitespace().collect();
        if !parts.is_empty() {
            return Some(parts[0].to_string());
        }

        None
    }

    /// Extract bookmark name from REF field instruction.
    ///
    /// REF fields have format: REF bookmark_name \h
    ///
    /// # Returns
    ///
    /// Bookmark name if this is a reference field, None otherwise
    pub fn extract_bookmark(&self) -> Option<String> {
        if self.field_type != FieldType::Reference {
            return None;
        }

        let inst = self.instruction.trim();
        if !inst.starts_with("REF") {
            return None;
        }

        // Extract bookmark name (second word)
        let parts: Vec<&str> = inst.split_whitespace().collect();
        if parts.len() >= 2 {
            return Some(parts[1].to_string());
        }

        None
    }

    /// Get the display text for the field.
    ///
    /// Returns the result text if available, otherwise the instruction.
    #[inline]
    pub fn display_text(&self) -> &str {
        if !self.result.is_empty() {
            &self.result
        } else {
            &self.instruction
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_hyperlink() {
        let field = Field::parse_instruction(r#"HYPERLINK "https://example.com""#);
        assert_eq!(field.field_type, FieldType::Hyperlink);
        assert_eq!(field.extract_url(), Some("https://example.com".to_string()));
    }

    #[test]
    fn test_parse_hyperlink_with_tooltip() {
        let field = Field::parse_instruction(r#"HYPERLINK "https://example.com" \o "Click here""#);
        assert_eq!(field.field_type, FieldType::Hyperlink);
        assert_eq!(field.extract_url(), Some("https://example.com".to_string()));
    }

    #[test]
    fn test_parse_ref() {
        let field = Field::parse_instruction("REF MyBookmark \\h");
        assert_eq!(field.field_type, FieldType::Reference);
        assert_eq!(field.extract_bookmark(), Some("MyBookmark".to_string()));
    }

    #[test]
    fn test_parse_page() {
        let field = Field::parse_instruction("PAGE");
        assert_eq!(field.field_type, FieldType::Page);
    }

    #[test]
    fn test_display_text() {
        let mut field = Field::new(
            FieldType::Hyperlink,
            Cow::Borrowed("HYPERLINK \"url\""),
            Cow::Borrowed("Click here"),
        );
        assert_eq!(field.display_text(), "Click here");

        field.result = Cow::Borrowed("");
        assert_eq!(field.display_text(), "HYPERLINK \"url\"");
    }
}
