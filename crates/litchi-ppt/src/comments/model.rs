//! Contextual presentation-comment author models.

use crate::package::{Error, Result};
use crate::presentation::ParsedSlideComments;
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
    /// Borrow authors in document order.
    pub fn iter(&self) -> impl Iterator<Item = &Author> {
        self.authors.iter()
    }

    /// Number of document-level comment authors.
    pub const fn len(&self) -> usize {
        self.authors.len()
    }

    /// Whether the document declares no comment authors.
    pub const fn is_empty(&self) -> bool {
        self.authors.is_empty()
    }

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
                return Err(Error::Corrupted(format!(
                    "Comment index exceeds the seed for author {name:?}"
                )));
            }
        }
        Ok(())
    }
}

/// Validated presentation-comment inventory.
///
/// PowerPoint stores author indexes in the document stream and comment atoms
/// in each slide's `___PPT10` extension. This owner joins those two scopes so
/// callers can validate the cross-record seed rule once and then traverse the
/// inert comments without rebuilding the binary record tree.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Catalog {
    authors: Authors,
    slides: Vec<ParsedSlideComments>,
}

impl Catalog {
    /// Build a catalog after checking every comment index against its author
    /// seed. No comment text, link, or external payload is activated.
    pub fn from_parts(authors: Authors, slides: Vec<ParsedSlideComments>) -> Result<Self> {
        for author in &authors.authors {
            let (Some(name), Some(seed)) = (&author.name, author.comment_index_seed) else {
                continue;
            };
            if slides.iter().any(|slide| {
                slide
                    .comments
                    .iter()
                    .any(|comment| comment.author == *name && comment.index > seed)
            }) {
                return Err(Error::Corrupted(format!(
                    "Comment index exceeds the seed for author {name:?}"
                )));
            }
        }
        Ok(Self { authors, slides })
    }

    /// Document-level comment authors in source order.
    pub const fn authors(&self) -> &Authors {
        &self.authors
    }

    /// Per-slide comment groups in presentation order.
    pub fn slides(&self) -> &[ParsedSlideComments] {
        &self.slides
    }

    /// Iterate over all comments in slide and source order.
    pub fn comments(&self) -> impl Iterator<Item = &ParsedComment> {
        self.slides.iter().flat_map(|slide| slide.comments.iter())
    }

    /// Find the comment group for a one-based slide number.
    pub fn slide(&self, slide_number: usize) -> Option<&ParsedSlideComments> {
        self.slides
            .iter()
            .find(|slide| slide.slide_number == slide_number)
    }
}
