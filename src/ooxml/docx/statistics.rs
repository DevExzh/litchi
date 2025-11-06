/// Document statistics and analysis for DOCX documents.
///
/// This module provides structures and functions for analyzing document statistics
/// including word count, character count, page count, and other metrics.
///
/// # Example
///
/// ```rust,no_run
/// use litchi::ooxml::docx::Package;
///
/// let pkg = Package::open("document.docx")?;
/// let doc = pkg.document()?;
///
/// // Get document statistics
/// let stats = doc.statistics()?;
/// println!("Words: {}", stats.word_count());
/// println!("Characters: {}", stats.character_count());
/// println!("Paragraphs: {}", stats.paragraph_count());
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
/// Document statistics.
///
/// Provides comprehensive statistics about a Word document including
/// counts of words, characters, paragraphs, and other elements.
///
/// # Performance
///
/// Statistics are calculated on-demand. For large documents, consider
/// caching the results if you need to access them multiple times.
#[derive(Debug, Clone, Default)]
pub struct DocumentStatistics {
    /// Total word count
    word_count: usize,

    /// Total character count (with spaces)
    character_count: usize,

    /// Character count (without spaces)
    character_count_no_spaces: usize,

    /// Total paragraph count
    paragraph_count: usize,

    /// Total line count (approximate)
    line_count: usize,

    /// Total page count (approximate)
    page_count: usize,

    /// Total table count
    table_count: usize,

    /// Total image count
    image_count: usize,

    /// Total drawing object count (shapes, text boxes)
    drawing_count: usize,
}

impl DocumentStatistics {
    /// Create new document statistics.
    pub fn new() -> Self {
        Self::default()
    }

    /// Get the word count.
    #[inline]
    pub fn word_count(&self) -> usize {
        self.word_count
    }

    /// Get the character count (including spaces).
    #[inline]
    pub fn character_count(&self) -> usize {
        self.character_count
    }

    /// Get the character count (excluding spaces).
    #[inline]
    pub fn character_count_no_spaces(&self) -> usize {
        self.character_count_no_spaces
    }

    /// Get the paragraph count.
    #[inline]
    pub fn paragraph_count(&self) -> usize {
        self.paragraph_count
    }

    /// Get the line count (approximate).
    ///
    /// This is an approximation based on text length and formatting.
    #[inline]
    pub fn line_count(&self) -> usize {
        self.line_count
    }

    /// Get the page count (approximate).
    ///
    /// This is an approximation based on text length and formatting.
    /// Actual page count may vary based on fonts, images, and layout.
    #[inline]
    pub fn page_count(&self) -> usize {
        self.page_count
    }

    /// Get the table count.
    #[inline]
    pub fn table_count(&self) -> usize {
        self.table_count
    }

    /// Get the image count.
    #[inline]
    pub fn image_count(&self) -> usize {
        self.image_count
    }

    /// Get the drawing object count (shapes, text boxes).
    #[inline]
    pub fn drawing_count(&self) -> usize {
        self.drawing_count
    }

    /// Set the word count.
    #[inline]
    pub(crate) fn set_word_count(&mut self, count: usize) {
        self.word_count = count;
    }

    /// Set the character count.
    #[inline]
    pub(crate) fn set_character_count(&mut self, count: usize) {
        self.character_count = count;
    }

    /// Set the character count (no spaces).
    #[inline]
    pub(crate) fn set_character_count_no_spaces(&mut self, count: usize) {
        self.character_count_no_spaces = count;
    }

    /// Set the paragraph count.
    #[inline]
    pub(crate) fn set_paragraph_count(&mut self, count: usize) {
        self.paragraph_count = count;
    }

    /// Set the line count.
    #[inline]
    pub(crate) fn set_line_count(&mut self, count: usize) {
        self.line_count = count;
    }

    /// Set the page count.
    #[inline]
    pub(crate) fn set_page_count(&mut self, count: usize) {
        self.page_count = count;
    }

    /// Set the table count.
    #[inline]
    pub(crate) fn set_table_count(&mut self, count: usize) {
        self.table_count = count;
    }

    /// Set the image count.
    #[inline]
    pub(crate) fn set_image_count(&mut self, count: usize) {
        self.image_count = count;
    }

    /// Set the drawing count.
    #[inline]
    pub(crate) fn set_drawing_count(&mut self, count: usize) {
        self.drawing_count = count;
    }
}

/// Calculate word count from text.
///
/// Counts words separated by whitespace. This is a simple implementation
/// that matches typical word processor behavior.
///
/// # Arguments
///
/// * `text` - The text to count words in
///
/// # Performance
///
/// Uses iterator-based counting for optimal performance.
#[inline]
pub fn count_words(text: &str) -> usize {
    text.split_whitespace().count()
}

/// Calculate character count (with spaces) from text.
#[inline]
pub fn count_characters(text: &str) -> usize {
    text.chars().count()
}

/// Calculate character count (without spaces) from text.
///
/// # Performance
///
/// Uses iterator filtering for optimal performance.
#[inline]
pub fn count_characters_no_spaces(text: &str) -> usize {
    text.chars().filter(|c| !c.is_whitespace()).count()
}

/// Estimate line count from text and average characters per line.
///
/// This is an approximation. Actual line count depends on font, size,
/// page width, and formatting.
///
/// # Arguments
///
/// * `text` - The text to estimate lines for
/// * `avg_chars_per_line` - Average characters per line (default: 80)
#[inline]
pub fn estimate_line_count(text: &str, avg_chars_per_line: usize) -> usize {
    let char_count = count_characters(text);
    if avg_chars_per_line == 0 {
        return 0;
    }
    char_count.div_ceil(avg_chars_per_line)
}

/// Estimate page count from line count and average lines per page.
///
/// This is an approximation. Actual page count depends on font, size,
/// margins, and formatting.
///
/// # Arguments
///
/// * `line_count` - Total number of lines
/// * `avg_lines_per_page` - Average lines per page (default: 45)
#[inline]
pub fn estimate_page_count(line_count: usize, avg_lines_per_page: usize) -> usize {
    if avg_lines_per_page == 0 {
        return 0;
    }
    line_count.div_ceil(avg_lines_per_page)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_count_words() {
        assert_eq!(count_words(""), 0);
        assert_eq!(count_words("hello"), 1);
        assert_eq!(count_words("hello world"), 2);
        assert_eq!(count_words("  hello   world  "), 2);
        assert_eq!(count_words("one two three four five"), 5);
    }

    #[test]
    fn test_count_characters() {
        assert_eq!(count_characters(""), 0);
        assert_eq!(count_characters("hello"), 5);
        assert_eq!(count_characters("hello world"), 11);
        assert_eq!(count_characters("  spaces  "), 10);
    }

    #[test]
    fn test_count_characters_no_spaces() {
        assert_eq!(count_characters_no_spaces(""), 0);
        assert_eq!(count_characters_no_spaces("hello"), 5);
        assert_eq!(count_characters_no_spaces("hello world"), 10);
        assert_eq!(count_characters_no_spaces("  spaces  "), 6);
    }

    #[test]
    fn test_estimate_line_count() {
        assert_eq!(estimate_line_count("", 80), 0);
        assert_eq!(estimate_line_count("x".repeat(80).as_str(), 80), 1);
        assert_eq!(estimate_line_count("x".repeat(160).as_str(), 80), 2);
        assert_eq!(estimate_line_count("x".repeat(81).as_str(), 80), 2);
    }

    #[test]
    fn test_estimate_page_count() {
        assert_eq!(estimate_page_count(0, 45), 0);
        assert_eq!(estimate_page_count(45, 45), 1);
        assert_eq!(estimate_page_count(90, 45), 2);
        assert_eq!(estimate_page_count(46, 45), 2);
    }

    #[test]
    fn test_document_statistics() {
        let mut stats = DocumentStatistics::new();
        assert_eq!(stats.word_count(), 0);
        assert_eq!(stats.character_count(), 0);

        stats.set_word_count(100);
        stats.set_character_count(500);
        stats.set_paragraph_count(10);

        assert_eq!(stats.word_count(), 100);
        assert_eq!(stats.character_count(), 500);
        assert_eq!(stats.paragraph_count(), 10);
    }
}
