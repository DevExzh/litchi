/// Header and footer access for Word documents
use super::package::Result;
use super::paragraph::Paragraph;
use super::parts::headers::HeaderFooterType;

/// A header or footer in a Word document
#[derive(Debug, Clone)]
pub struct HeaderFooter {
    /// Type of header/footer
    pub header_footer_type: HeaderFooterType,
    /// Text content
    pub text: String,
    /// Paragraphs in this header/footer
    pub paragraphs: Vec<Paragraph>,
}

impl HeaderFooter {
    /// Create a new header or footer
    pub fn new(header_footer_type: HeaderFooterType, text: String) -> Self {
        Self {
            header_footer_type,
            text,
            paragraphs: Vec::new(),
        }
    }

    /// Get the text content
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Get the paragraphs
    pub fn paragraphs(&self) -> Result<&[Paragraph]> {
        Ok(&self.paragraphs)
    }

    /// Check if this is a header
    pub fn is_header(&self) -> bool {
        self.header_footer_type.is_header()
    }

    /// Check if this is a footer
    pub fn is_footer(&self) -> bool {
        self.header_footer_type.is_footer()
    }
}
