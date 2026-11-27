//! `PowerPoint` text-run record-tree orchestration and read views.

use super::model::{ParagraphRun, TextRun};

/// Text run extractor for `PowerPoint` slides.
///
/// Based on Apache POI's `TextRun`, `StyleTextPropAtom`, and related classes.
pub struct TextRunExtractor {
    /// Full text content
    pub(super) text: String,
    /// Text runs with formatting
    pub(super) runs: Vec<TextRun>,
    /// Paragraph-level formatting runs
    pub(super) paragraph_runs: Vec<ParagraphRun>,
    /// Most recently encountered text atom awaiting its style atom
    pub(super) pending_text: Option<(String, usize)>,
}

impl TextRunExtractor {
    /// Create a new text run extractor.
    #[must_use]
    pub fn new() -> Self {
        Self {
            text: String::new(),
            runs: Vec::new(),
            paragraph_runs: Vec::new(),
            pending_text: None,
        }
    }

    /// Get the full extracted text.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Get all text runs.
    #[must_use]
    pub fn runs(&self) -> &[TextRun] {
        &self.runs
    }

    /// Get paragraph-level formatting runs.
    #[must_use]
    pub fn paragraph_runs(&self) -> &[ParagraphRun] {
        &self.paragraph_runs
    }

    /// Get the number of runs.
    #[must_use]
    pub fn run_count(&self) -> usize {
        self.runs.len()
    }
}

impl Default for TextRunExtractor {
    fn default() -> Self {
        Self::new()
    }
}
