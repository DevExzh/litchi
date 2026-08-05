//! Typed semantic values for legacy Word header and footer stories.

use crate::package::Result;
use crate::paragraph::Paragraph;
use crate::parts::headers::HeaderFooterType;

/// A header or footer in a legacy Word document.
#[derive(Debug, Clone)]
pub struct HeaderFooter {
    /// Type of header/footer.
    pub header_footer_type: HeaderFooterType,
    /// Text content extracted from the story.
    pub text: String,
    /// Paragraphs in this header/footer story.
    pub paragraphs: Vec<Paragraph>,
}

impl HeaderFooter {
    /// Create a semantic story from its decoded DOC payload.
    pub fn new(header_footer_type: HeaderFooterType, text: String) -> Self {
        super::codec::decode(super::package::Story::new(
            header_footer_type,
            text,
            Vec::new(),
        ))
    }

    /// Get the text content.
    #[inline]
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Get the paragraphs.
    pub fn paragraphs(&self) -> Result<&[Paragraph]> {
        Ok(&self.paragraphs)
    }

    /// Check if this is a header.
    #[inline]
    pub fn is_header(&self) -> bool {
        self.header_footer_type.is_header()
    }

    /// Check if this is a footer.
    #[inline]
    pub fn is_footer(&self) -> bool {
        self.header_footer_type.is_footer()
    }
}
