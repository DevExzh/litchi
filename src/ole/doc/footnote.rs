/// Footnote and endnote structures for Word documents
use super::package::Result;
use super::paragraph::Paragraph;

/// A footnote in a Word document
#[derive(Debug, Clone)]
pub struct Footnote {
    /// Reference position in main document
    pub reference_position: u32,
    /// Reference number/mark
    pub number: u16,
    /// Text content
    pub text: String,
    /// Paragraphs in this footnote
    pub paragraphs: Vec<Paragraph>,
}

impl Footnote {
    /// Create a new footnote
    pub fn new(reference_position: u32, number: u16, text: String) -> Self {
        Self {
            reference_position,
            number,
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
}

/// An endnote in a Word document (same structure as footnote)
pub type Endnote = Footnote;
