//! Contextual presentation-comment author models.

use crate::package::{PptError, Result};
use crate::slide::ParsedComment;

/// Document-level metadata for one presentation-comment author.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Author {
    /// Optional author display name.
    pub name: Option<String>,
    /// Optional zero-based application-defined display color index.
    pub color_index: Option<i32>,
    /// Optional seed for the next comment index created by this author.
    pub comment_index_seed: Option<i32>,
}

/// PowerPoint 10 presentation-comment authors.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Authors {
    pub authors: Vec<Author>,
}

impl Authors {
    /// Find the first author with the specified display name.
    pub fn find(&self, name: &str) -> Option<&Author> {
        self.authors
            .iter()
            .find(|author| author.name.as_deref() == Some(name))
    }

    /// Validate author index seeds against a collection of parsed slide comments.
    pub fn validate_comments(&self, comments: &[ParsedComment]) -> Result<()> {
        for author in &self.authors {
            let (Some(name), Some(seed)) = (&author.name, author.comment_index_seed) else {
                continue;
            };
            if comments
                .iter()
                .any(|comment| comment.author == *name && comment.index > seed)
            {
                return Err(PptError::Corrupted(format!(
                    "Comment index exceeds the seed for author {name:?}"
                )));
            }
        }
        Ok(())
    }
}
