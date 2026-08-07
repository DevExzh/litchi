//! Web-template paragraph semantics.

/// A semantic web-template paragraph.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Paragraph {
    text: String,
}

impl Paragraph {
    pub fn new(text: impl Into<String>) -> Self {
        Self { text: text.into() }
    }

    /// Returns the paragraph text.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }
}
