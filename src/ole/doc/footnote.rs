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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_footnote_new() {
        let footnote = Footnote::new(100, 1, "Footnote text".to_string());

        assert_eq!(footnote.reference_position, 100);
        assert_eq!(footnote.number, 1);
        assert_eq!(footnote.text, "Footnote text");
        assert!(footnote.paragraphs.is_empty());
    }

    #[test]
    fn test_footnote_text_method() {
        let footnote = Footnote::new(0, 1, "Test content".to_string());

        assert_eq!(footnote.text(), "Test content");
    }

    #[test]
    fn test_footnote_paragraphs_method() {
        let mut footnote = Footnote::new(50, 2, "Footnote with paragraphs".to_string());
        let para = Paragraph::new("Para text".to_string());
        footnote.paragraphs.push(para);

        let paragraphs = footnote.paragraphs().unwrap();
        assert_eq!(paragraphs.len(), 1);
    }

    #[test]
    fn test_footnote_with_multiple_paragraphs() {
        let mut footnote = Footnote::new(100, 1, "Multi-para footnote".to_string());

        for i in 0..3 {
            let para_text = format!("Paragraph {}", i);
            footnote.paragraphs.push(Paragraph::new(para_text));
        }

        assert_eq!(footnote.paragraphs.len(), 3);
    }

    #[test]
    fn test_footnote_clone() {
        let footnote = Footnote::new(200, 5, "Clonable footnote".to_string());
        let cloned = footnote.clone();

        assert_eq!(cloned.reference_position, footnote.reference_position);
        assert_eq!(cloned.number, footnote.number);
        assert_eq!(cloned.text, footnote.text);
    }

    #[test]
    fn test_footnote_debug() {
        let footnote = Footnote::new(100, 1, "Debug footnote".to_string());
        let debug_str = format!("{:?}", footnote);

        assert!(debug_str.contains("Footnote"));
        assert!(debug_str.contains("Debug footnote"));
    }

    #[test]
    fn test_footnote_empty_text() {
        let footnote = Footnote::new(0, 1, "".to_string());

        assert_eq!(footnote.text(), "");
        assert!(footnote.paragraphs.is_empty());
    }

    #[test]
    fn test_footnote_with_unicode() {
        let footnote = Footnote::new(100, 1, "Unicode: 你好世界 🎉".to_string());

        assert_eq!(footnote.text(), "Unicode: 你好世界 🎉");
    }

    #[test]
    fn test_footnote_large_number() {
        let footnote = Footnote::new(0, 9999, "High number footnote".to_string());

        assert_eq!(footnote.number, 9999);
    }

    #[test]
    fn test_footnote_large_position() {
        let footnote = Footnote::new(u32::MAX, 1, "Max position".to_string());

        assert_eq!(footnote.reference_position, u32::MAX);
    }

    #[test]
    fn test_endnote_as_footnote_alias() {
        // Endnote is a type alias for Footnote
        let endnote: Endnote = Footnote::new(300, 10, "Endnote content".to_string());

        assert_eq!(endnote.reference_position, 300);
        assert_eq!(endnote.number, 10);
        assert_eq!(endnote.text(), "Endnote content");
    }

    #[test]
    fn test_footnote_equality() {
        let fn1 = Footnote::new(100, 1, "Same".to_string());
        let fn2 = Footnote::new(100, 1, "Same".to_string());
        let fn3 = Footnote::new(200, 2, "Different".to_string());

        assert_eq!(fn1.reference_position, fn2.reference_position);
        assert_eq!(fn1.number, fn2.number);
        assert_ne!(fn1.reference_position, fn3.reference_position);
    }

    #[test]
    fn test_footnote_with_nested_paragraphs() {
        let mut footnote = Footnote::new(100, 1, "Complex footnote".to_string());

        // Add paragraphs with different content
        footnote
            .paragraphs
            .push(Paragraph::new("First para".to_string()));
        footnote
            .paragraphs
            .push(Paragraph::new("Second para".to_string()));
        footnote
            .paragraphs
            .push(Paragraph::new("Third para".to_string()));

        assert_eq!(footnote.paragraphs().unwrap().len(), 3);
    }

    #[test]
    fn test_multiple_footnotes() {
        let footnotes = vec![
            Footnote::new(10, 1, "First".to_string()),
            Footnote::new(50, 2, "Second".to_string()),
            Footnote::new(100, 3, "Third".to_string()),
        ];

        assert_eq!(footnotes.len(), 3);
        assert_eq!(footnotes[0].number, 1);
        assert_eq!(footnotes[1].number, 2);
        assert_eq!(footnotes[2].number, 3);
    }
}
