//! Bounded DOCX document statistics and text-derived metrics.
//!
//! The concrete crate owns the immutable statistics value and its allocation-free
//! text counters. Package adapters provide the traversal-specific counts.
//!
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
pub struct Statistics {
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

impl Statistics {
    /// Create new document statistics.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Get the word count.
    #[inline]
    #[must_use]
    pub fn word_count(&self) -> usize {
        self.word_count
    }

    /// Get the character count (including spaces).
    #[inline]
    #[must_use]
    pub fn character_count(&self) -> usize {
        self.character_count
    }

    /// Get the character count (excluding spaces).
    #[inline]
    #[must_use]
    pub fn character_count_no_spaces(&self) -> usize {
        self.character_count_no_spaces
    }

    /// Get the paragraph count.
    #[inline]
    #[must_use]
    pub fn paragraph_count(&self) -> usize {
        self.paragraph_count
    }

    /// Get the line count (approximate).
    ///
    /// This is an approximation based on text length and formatting.
    #[inline]
    #[must_use]
    pub fn line_count(&self) -> usize {
        self.line_count
    }

    /// Get the page count (approximate).
    ///
    /// This is an approximation based on text length and formatting.
    /// Actual page count may vary based on fonts, images, and layout.
    #[inline]
    #[must_use]
    pub fn page_count(&self) -> usize {
        self.page_count
    }

    /// Get the table count.
    #[inline]
    #[must_use]
    pub fn table_count(&self) -> usize {
        self.table_count
    }

    /// Get the image count.
    #[inline]
    #[must_use]
    pub fn image_count(&self) -> usize {
        self.image_count
    }

    /// Get the drawing object count (shapes, text boxes).
    #[inline]
    #[must_use]
    pub fn drawing_count(&self) -> usize {
        self.drawing_count
    }

    /// Build a statistics snapshot from precomputed document metrics.
    ///
    /// The DOCX package adapter owns traversal and counting; this constructor
    /// keeps the resulting value independent from package identifiers.
    #[must_use]
    pub const fn from_counts(
        word_count: usize,
        character_count: usize,
        character_count_no_spaces: usize,
        paragraph_count: usize,
        line_count: usize,
        page_count: usize,
        table_count: usize,
        image_count: usize,
        drawing_count: usize,
    ) -> Self {
        Self {
            word_count,
            character_count,
            character_count_no_spaces,
            paragraph_count,
            line_count,
            page_count,
            table_count,
            image_count,
            drawing_count,
        }
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
#[must_use]
pub fn count_words(text: &str) -> usize {
    text.split_whitespace().count()
}

/// Calculate character count (with spaces) from text.
#[inline]
#[must_use]
pub fn count_characters(text: &str) -> usize {
    text.chars().count()
}

/// Calculate character count (without spaces) from text.
///
/// # Performance
///
/// Uses iterator filtering for optimal performance.
#[inline]
#[must_use]
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
#[must_use]
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
#[must_use]
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
        let stats = Statistics::from_counts(100, 500, 450, 10, 12, 1, 2, 3, 4);

        assert_eq!(stats.word_count(), 100);
        assert_eq!(stats.character_count(), 500);
        assert_eq!(stats.paragraph_count(), 10);
    }
}
