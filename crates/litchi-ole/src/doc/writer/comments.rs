//! Comment input types for the legacy Word writer.

/// A point comment to add to a DOC main story.
///
/// The writer emits a U+0005 reference at the requested position and stores the
/// body in the comment subdocument. Range comments require annotation bookmark
/// tables and are not represented by this type.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CommentEntry {
    /// Requested UTF-16 CP in the main document.
    pub ref_position: u32,
    /// Comment body text.
    pub text: String,
    /// Full comment author name.
    pub author: String,
    /// Author initials (at most nine UTF-16 code units).
    pub initials: String,
}

impl CommentEntry {
    /// Create a point comment.
    pub fn new(
        ref_position: u32,
        text: impl Into<String>,
        author: impl Into<String>,
        initials: impl Into<String>,
    ) -> Self {
        Self {
            ref_position,
            text: text.into(),
            author: author.into(),
            initials: initials.into(),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn creates_unicode_comment_entry() {
        let entry = CommentEntry::new(3, "检查 😀", "张三", "张");
        assert_eq!(entry.ref_position, 3);
        assert_eq!(entry.text, "检查 😀");
        assert_eq!(entry.author, "张三");
        assert_eq!(entry.initials, "张");
    }
}
