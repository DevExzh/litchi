/// A comment attached to content in a legacy Word document.
use super::package::Result;
use super::paragraph::Paragraph;

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
        assert_eq!(comment.text(), "Review this");
        assert!(comment.paragraphs().unwrap().is_empty());
    }
}
