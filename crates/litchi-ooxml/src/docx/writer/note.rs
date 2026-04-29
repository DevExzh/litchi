use super::paragraph::MutableParagraph;

/// Footnote or endnote entry.
#[derive(Debug)]
pub struct Note {
    /// Note ID (starting from 1)
    pub(crate) id: u32,
    /// Note content (paragraphs)
    pub(crate) paragraphs: Vec<MutableParagraph>,
}

impl Note {
    pub(crate) fn new(id: u32) -> Self {
        Self {
            id,
            paragraphs: Vec::new(),
        }
    }

    /// Add a paragraph to this note.
    pub fn add_paragraph(&mut self) -> &mut MutableParagraph {
        let para = MutableParagraph::new();
        self.paragraphs.push(para);
        self.paragraphs.last_mut().unwrap()
    }

    /// Add a paragraph with text to this note.
    pub fn add_paragraph_with_text(&mut self, text: &str) -> &mut MutableParagraph {
        let para = self.add_paragraph();
        para.add_run_with_text(text);
        para
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_note_new() {
        let note = Note::new(1);
        assert_eq!(note.id, 1);
        assert!(note.paragraphs.is_empty());
    }

    #[test]
    fn test_note_add_paragraph() {
        let mut note = Note::new(1);
        note.add_paragraph();
        assert_eq!(note.paragraphs.len(), 1);
        assert!(note.paragraphs[0].elements.is_empty());
    }

    #[test]
    fn test_note_add_paragraph_with_text() {
        let mut note = Note::new(1);
        note.add_paragraph_with_text("This is a note paragraph.");
        assert_eq!(note.paragraphs.len(), 1);
        // Verify the paragraph has a run with the text
        let para = &note.paragraphs[0];
        assert!(!para.elements.is_empty());
    }

    #[test]
    fn test_note_multiple_paragraphs() {
        let mut note = Note::new(2);
        note.add_paragraph_with_text("First paragraph.");
        note.add_paragraph_with_text("Second paragraph.");
        assert_eq!(note.paragraphs.len(), 2);
    }

    #[test]
    fn test_note_debug() {
        let note = Note::new(5);
        let debug_str = format!("{:?}", note);
        assert!(debug_str.contains("5"));
        assert!(debug_str.contains("paragraphs"));
    }
}
