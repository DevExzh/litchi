//! Structured slide model.
/// Represents a slide in a Keynote presentation
#[derive(Debug, Clone)]
pub struct Slide {
    /// Slide index (0-based)
    pub index: usize,
    /// Slide title
    pub title: Option<String>,
    /// Text content on the slide
    pub text_content: Vec<String>,
    /// Notes associated with the slide
    pub notes: Option<String>,
}

impl Slide {
    /// Create a new slide
    pub fn new(index: usize) -> Self {
        Self {
            index,
            title: None,
            text_content: Vec::new(),
            notes: None,
        }
    }

    /// Get all text from the slide (title + content + notes)
    pub fn all_text(&self) -> Vec<String> {
        let mut all = Vec::new();
        if let Some(ref title) = self.title {
            all.push(title.clone());
        }
        all.extend(self.text_content.clone());
        if let Some(ref notes) = self.notes {
            all.push(notes.clone());
        }
        all
    }
}
