//! Comment input types for the legacy Word writer.

use crate::ExtendedMetadata;

/// A comment to add to a DOC main story.
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
    /// Optional `(start, exclusive_end)` range in the emitted main-story CP space.
    ///
    /// When absent, the writer creates a point comment. Both positions must be
    /// within the final main-document character count.
    pub range: Option<(u32, u32)>,
    /// Optional Word 2002+ timestamp, reply-tree, and ink metadata.
    ///
    /// Parent indexes refer to comments in emitted main-document reference
    /// order. The writer emits a default top-level metadata record when this is
    /// absent because `AtrdExtra` is parallel to all comment descriptors.
    pub extended_metadata: Option<ExtendedMetadata>,
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
            range: None,
            extended_metadata: None,
        }
    }

    /// Attach this comment to a main-document range.
    pub fn with_range(mut self, start: u32, exclusive_end: u32) -> Self {
        self.range = Some((start, exclusive_end));
        self
    }

    /// Set Word 2002+ timestamp, reply-tree, and ink metadata.
    pub fn with_extended_metadata(mut self, metadata: ExtendedMetadata) -> Self {
        self.extended_metadata = Some(metadata);
        self
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
        assert_eq!(entry.range, None);
        assert_eq!(entry.extended_metadata, None);
    }
}
