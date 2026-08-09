//! Semantic model for classic (legacy-note) worksheet comments.

use std::collections::BTreeMap;

/// One classic note attached to a worksheet cell.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Comment {
    /// A checked one-based A1 cell reference.
    pub cell_ref: String,
    /// The resolved author text from the comments part's author table.
    pub author: String,
    /// The zero-based author index stored on `comment/@authorId`.
    pub author_id: u32,
    /// Plain text collected from simple or rich `SpreadsheetML` text runs.
    pub text: String,
    /// Optional producer extension GUID.
    pub guid: Option<String>,
    /// Optional legacy VML shape identifier.
    pub shape_id: Option<u32>,
}

/// All classic comments and their author table from one comments part.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Comments {
    /// Authors in the exact order addressed by `Comment::author_id`.
    pub authors: Vec<String>,
    /// Comments indexed by their checked cell reference.
    pub comments: BTreeMap<String, Comment>,
}

impl Comments {
    /// Return a comment by its checked A1 cell reference.
    #[must_use]
    pub fn get(&self, cell_ref: &str) -> Option<&Comment> {
        self.comments.get(cell_ref)
    }

    /// Number of notes in this part.
    #[must_use]
    pub fn len(&self) -> usize {
        self.comments.len()
    }

    /// Whether this part contains no notes.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.comments.is_empty()
    }

    /// Iterate over notes in deterministic cell-reference order.
    pub fn iter(&self) -> impl Iterator<Item = (&str, &Comment)> {
        self.comments
            .iter()
            .map(|(cell, comment)| (cell.as_str(), comment))
    }
}

/// A worksheet comments part together with its owning relationship identity.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Part {
    /// Worksheet OPC part name owning the relationship.
    pub worksheet_part_name: String,
    /// Relationship ID on the worksheet part.
    pub relationship_id: String,
    /// Absolute OPC part name of the comments resource.
    pub part_name: String,
    /// Parsed classic comments.
    pub comments: Comments,
}
