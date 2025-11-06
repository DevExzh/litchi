/// Table of Contents support for DOCX documents.
///
/// A TOC is implemented using a complex field with switches to control its behavior.
use crate::ooxml::error::Result;
use std::fmt::Write as FmtWrite;

/// A table of contents (TOC) field.
///
/// The TOC uses heading styles to build an outline of the document.
///
/// # Examples
///
/// ```rust,ignore
/// use litchi::ooxml::docx::writer::TableOfContents;
///
/// let toc = TableOfContents::new()
///     .heading_levels(1, 3)
///     .hyperlinks(true);
/// ```
#[derive(Debug, Clone)]
pub struct TableOfContents {
    /// Starting heading level (default: 1)
    start_level: u32,
    /// Ending heading level (default: 3)
    end_level: u32,
    /// Include hyperlinks (default: true)
    hyperlinks: bool,
    /// Include page numbers (default: true)
    page_numbers: bool,
    /// Right-align page numbers (default: true)
    right_align_page_numbers: bool,
    /// Include outline levels (default: false)
    use_outline_levels: bool,
    /// Custom title (default: None, uses Word's default)
    title: Option<String>,
}

impl TableOfContents {
    /// Create a new table of contents with default settings.
    ///
    /// Default settings:
    /// - Heading levels 1-3
    /// - Includes hyperlinks
    /// - Includes page numbers (right-aligned)
    ///
    /// # Examples
    ///
    /// ```rust,ignore
    /// let toc = TableOfContents::new();
    /// ```
    pub fn new() -> Self {
        Self {
            start_level: 1,
            end_level: 3,
            hyperlinks: true,
            page_numbers: true,
            right_align_page_numbers: true,
            use_outline_levels: false,
            title: None,
        }
    }

    /// Set the heading levels to include.
    ///
    /// # Arguments
    ///
    /// * `start` - Starting level (1-9)
    /// * `end` - Ending level (1-9, must be >= start)
    ///
    /// # Examples
    ///
    /// ```rust,ignore
    /// let toc = TableOfContents::new().heading_levels(1, 4);
    /// ```
    pub fn heading_levels(mut self, start: u32, end: u32) -> Self {
        self.start_level = start.clamp(1, 9);
        self.end_level = end.clamp(start, 9);
        self
    }

    /// Set whether to include hyperlinks (default: true).
    pub fn hyperlinks(mut self, enabled: bool) -> Self {
        self.hyperlinks = enabled;
        self
    }

    /// Set whether to include page numbers (default: true).
    pub fn page_numbers(mut self, enabled: bool) -> Self {
        self.page_numbers = enabled;
        self
    }

    /// Set whether to right-align page numbers (default: true).
    pub fn right_align_page_numbers(mut self, enabled: bool) -> Self {
        self.right_align_page_numbers = enabled;
        self
    }

    /// Use outline levels instead of heading styles (default: false).
    pub fn use_outline_levels(mut self, enabled: bool) -> Self {
        self.use_outline_levels = enabled;
        self
    }

    /// Set a custom title for the TOC.
    pub fn title(mut self, title: impl Into<String>) -> Self {
        self.title = Some(title.into());
        self
    }

    /// Build the TOC field instruction string.
    ///
    /// Format: TOC \o "1-3" \h \z \u
    /// - \o "1-3" = outline levels 1 through 3
    /// - \h = hyperlinks
    /// - \z = hide tab leader and page numbers in Web Layout view
    /// - \u = use outline levels
    pub fn build_field_instruction(&self) -> String {
        let mut instruction = String::from("TOC");

        // Outline levels
        if self.use_outline_levels {
            write!(
                &mut instruction,
                r#" \u "{}-{}""#,
                self.start_level, self.end_level
            )
            .unwrap();
        } else {
            write!(
                &mut instruction,
                r#" \o "{}-{}""#,
                self.start_level, self.end_level
            )
            .unwrap();
        }

        // Hyperlinks
        if self.hyperlinks {
            instruction.push_str(" \\h");
        }

        // Hide in web layout
        instruction.push_str(" \\z");

        instruction
    }

    /// Get the title (if set).
    pub fn get_title(&self) -> Option<&str> {
        self.title.as_deref()
    }

    /// Get the starting heading level.
    pub fn start_level(&self) -> u8 {
        self.start_level as u8
    }

    /// Get the ending heading level.
    pub fn end_level(&self) -> u8 {
        self.end_level as u8
    }

    /// Generate XML for the TOC field.
    ///
    /// Returns a paragraph containing the TOC field with optional title.
    ///
    /// Note: This method generates standalone XML. The field-based approach via
    /// `MutableField::toc()` is preferred for integration with the document structure.
    #[allow(dead_code)]
    pub(crate) fn to_xml(&self) -> Result<String> {
        let mut xml = String::with_capacity(1024);

        // Optional title paragraph
        if let Some(title) = &self.title {
            xml.push_str("<w:p>");
            xml.push_str(r#"<w:pPr><w:pStyle w:val="TOCHeading"/></w:pPr>"#);
            xml.push_str("<w:r><w:t>");
            xml.push_str(&escape_xml(title));
            xml.push_str("</w:t></w:r>");
            xml.push_str("</w:p>");
        }

        // TOC field paragraph
        xml.push_str("<w:p>");

        // Field begin
        xml.push_str(r#"<w:fldSimple w:instr=""#);
        xml.push_str(&escape_xml(&self.build_field_instruction()));
        xml.push_str(r#"">"#);

        // Placeholder text (updated when user opens document)
        xml.push_str(r#"<w:r><w:t>"#);
        xml.push_str(
            "Right-click and select &quot;Update Field&quot; to generate the table of contents.",
        );
        xml.push_str("</w:t></w:r>");

        xml.push_str("</w:fldSimple>");
        xml.push_str("</w:p>");

        Ok(xml)
    }

    /// Generate XML for the TOC field using complex field structure.
    ///
    /// This is more compatible with some Word versions than fldSimple.
    ///
    /// Note: This method generates standalone XML. The field-based approach via
    /// `MutableField::toc()` is preferred for integration with the document structure.
    #[allow(dead_code)]
    pub(crate) fn to_complex_field_xml(&self) -> Result<String> {
        let mut xml = String::with_capacity(1024);

        // Optional title paragraph
        if let Some(title) = &self.title {
            xml.push_str("<w:p>");
            xml.push_str(r#"<w:pPr><w:pStyle w:val="TOCHeading"/></w:pPr>"#);
            xml.push_str("<w:r><w:t>");
            xml.push_str(&escape_xml(title));
            xml.push_str("</w:t></w:r>");
            xml.push_str("</w:p>");
        }

        // TOC field paragraph with complex field structure
        xml.push_str("<w:p>");

        // Field begin
        xml.push_str(r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#);

        // Field instruction
        xml.push_str("<w:r><w:instrText>");
        xml.push_str(&escape_xml(&self.build_field_instruction()));
        xml.push_str("</w:instrText></w:r>");

        // Field separator
        xml.push_str(r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#);

        // Field result (placeholder text)
        xml.push_str(r#"<w:r><w:t>"#);
        xml.push_str(
            "Right-click and select &quot;Update Field&quot; to generate the table of contents.",
        );
        xml.push_str("</w:t></w:r>");

        // Field end
        xml.push_str(r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#);

        xml.push_str("</w:p>");

        Ok(xml)
    }
}

impl Default for TableOfContents {
    fn default() -> Self {
        Self::new()
    }
}

/// Escape XML special characters.
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
    fn test_create_toc() {
        let toc = TableOfContents::new();
        assert_eq!(toc.start_level, 1);
        assert_eq!(toc.end_level, 3);
        assert!(toc.hyperlinks);
        assert!(toc.page_numbers);
    }

    #[test]
    fn test_toc_heading_levels() {
        let toc = TableOfContents::new().heading_levels(2, 5);
        assert_eq!(toc.start_level, 2);
        assert_eq!(toc.end_level, 5);
    }

    #[test]
    fn test_toc_field_instruction() {
        let toc = TableOfContents::new().heading_levels(1, 4).hyperlinks(true);

        let instr = toc.build_field_instruction();
        assert!(instr.contains(r#"TOC"#));
        assert!(instr.contains(r#""1-4""#));
        assert!(instr.contains(r#"\h"#));
    }

    #[test]
    fn test_toc_xml() {
        let toc = TableOfContents::new().title("Contents");
        let xml = toc.to_xml().unwrap();

        assert!(xml.contains("TOCHeading"));
        assert!(xml.contains("Contents"));
        assert!(xml.contains("fldSimple"));
    }

    #[test]
    fn test_toc_complex_field_xml() {
        let toc = TableOfContents::new();
        let xml = toc.to_complex_field_xml().unwrap();

        assert!(xml.contains("fldChar"));
        assert!(xml.contains("instrText"));
        assert!(xml.contains("TOC"));
    }
}
