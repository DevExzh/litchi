//! Structured section model.
/// Represents a section in a Pages document
#[derive(Debug, Clone)]
pub struct Section {
    /// Section index (0-based)
    pub index: usize,
    /// Section heading
    pub heading: Option<String>,
    /// Paragraphs in this section
    pub paragraphs: Vec<String>,
}

impl Section {
    /// Create a new section
    pub fn new(index: usize) -> Self {
        Self {
            index,
            heading: None,
            paragraphs: Vec::new(),
        }
    }

    /// Get all text from the section
    pub fn all_text(&self) -> Vec<String> {
        let mut all = Vec::new();
        if let Some(ref heading) = self.heading {
            all.push(heading.clone());
        }
        all.extend(self.paragraphs.clone());
        all
    }
}
