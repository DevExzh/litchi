//! Comment support for XLSB

/// Comment information
///
/// Represents a cell comment with author and text.
#[derive(Debug, Clone)]
pub struct Comment {
    /// Row (0-based)
    pub row: u32,
    /// Column (0-based)
    pub col: u32,
    /// Author of the comment
    pub author: String,
    /// Comment text
    pub text: String,
    /// Whether comment is visible
    pub visible: bool,
}

impl Comment {
    /// Create a new comment
    ///
    /// # Example
    ///
    /// ```rust
    /// use litchi::ooxml::xlsb::comments::Comment;
    ///
    /// let comment = Comment::new(0, 0, "John".to_string(), "This is a note".to_string());
    /// ```
    pub fn new(row: u32, col: u32, author: String, text: String) -> Self {
        Comment {
            row,
            col,
            author,
            text,
            visible: false,
        }
    }

    /// Set visibility
    pub fn set_visible(&mut self, visible: bool) {
        self.visible = visible;
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_comment_creation() {
        let comment = Comment::new(0, 0, "John".to_string(), "Note".to_string());
        assert_eq!(comment.row, 0);
        assert_eq!(comment.col, 0);
        assert_eq!(comment.author, "John");
        assert!(!comment.visible);
    }
}
