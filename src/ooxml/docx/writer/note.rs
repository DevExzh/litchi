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
