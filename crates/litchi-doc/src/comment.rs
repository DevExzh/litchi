/// A comment attached to content in a legacy Word document.
use super::package::Result;
use super::paragraph::Paragraph;

/// A date and time stored in Word's packed DTTM representation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DateTime {
    /// Calendar year.
    pub year: u16,
    /// Month in the range 1 through 12.
    pub month: u8,
    /// Day of the month in the range 1 through 31.
    pub day: u8,
    /// Hour in the range 0 through 23.
    pub hour: u8,
    /// Minute in the range 0 through 59.
    pub minute: u8,
    /// Day of the week, where zero is Sunday and six is Saturday.
    pub weekday: u8,
}

/// Word 2002+ metadata associated with a comment.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ExtendedMetadata {
    /// Date and time when the comment was created or last modified.
    ///
    /// This is `None` when the packed DTTM has a zero month or day and Word
    /// requires the timestamp to be ignored.
    pub modified_at: Option<DateTime>,
    /// Depth in the reply tree. Top-level comments have depth zero.
    pub depth: u32,
    /// Zero-based index of the parent comment in main-document reference order.
    pub parent_index: Option<usize>,
    /// Whether this comment is an ink annotation.
    pub is_ink: bool,
}

/// User-facing DOC comment data.
#[derive(Debug, Clone)]
pub struct Comment {
    /// Position of the comment-reference character in the main document.
    pub reference_position: u32,
    /// Full author name.
    pub author: String,
    /// Author initials stored with the comment.
    pub initials: String,
    /// Annotation-bookmark tag, or `None` for a zero-length commented range.
    pub bookmark_tag: Option<u32>,
    /// Start CP of the commented main-document range.
    pub range_start: Option<u32>,
    /// Exclusive end CP of the commented main-document range.
    pub range_end: Option<u32>,
    /// Word 2002+ timestamp, reply-tree, and ink metadata, when present.
    pub extended_metadata: Option<ExtendedMetadata>,
    /// Comment body text, excluding its structural U+0005 marker.
    pub text: String,
    /// Paragraphs in the comment body.
    pub paragraphs: Vec<Paragraph>,
}

impl Comment {
    /// Construct a comment.
    pub fn new(
        reference_position: u32,
        author: String,
        initials: String,
        bookmark_tag: Option<u32>,
        text: String,
    ) -> Self {
        Self {
            reference_position,
            author,
            initials,
            bookmark_tag,
            range_start: None,
            range_end: None,
            extended_metadata: None,
            text,
            paragraphs: Vec::new(),
        }
    }

    /// Comment body text.
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Paragraphs in the comment body.
    pub fn paragraphs(&self) -> Result<&[Paragraph]> {
        Ok(&self.paragraphs)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn exposes_comment_metadata_and_text() {
        let comment = Comment::new(
            42,
            "Alice Example".to_string(),
            "AE".to_string(),
            Some(7),
            "Review this".to_string(),
        );
        assert_eq!(comment.reference_position, 42);
        assert_eq!(comment.author, "Alice Example");
        assert_eq!(comment.initials, "AE");
        assert_eq!(comment.bookmark_tag, Some(7));
        assert_eq!(comment.range_start, None);
        assert_eq!(comment.range_end, None);
        assert_eq!(comment.extended_metadata, None);
        assert_eq!(comment.text(), "Review this");
        assert!(comment.paragraphs().unwrap().is_empty());
    }
}
