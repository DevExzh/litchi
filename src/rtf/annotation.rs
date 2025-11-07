//! RTF annotation and comment support.
//!
//! This module provides support for comments, revisions, and other annotations
//! in RTF documents.

use std::borrow::Cow;

/// Annotation type
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AnnotationType {
    /// Comment/note
    Comment,
    /// Revision mark (tracked change)
    Revision,
    /// Highlight
    Highlight,
}

/// Revision type (for tracked changes)
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RevisionType {
    /// Inserted text
    Insertion,
    /// Deleted text
    Deletion,
    /// Formatting change
    FormatChange,
    /// Moved from location
    MovedFrom,
    /// Moved to location
    MovedTo,
}

/// Comment or annotation
#[derive(Debug, Clone)]
pub struct Annotation<'a> {
    /// Annotation type
    pub annotation_type: AnnotationType,
    /// Annotation ID
    pub id: i32,
    /// Author name
    pub author: Cow<'a, str>,
    /// Creation date (RTF datetime format)
    pub date: Option<Cow<'a, str>>,
    /// Comment text
    pub text: Cow<'a, str>,
    /// Reference position in document
    pub position: usize,
}

impl<'a> Annotation<'a> {
    /// Create a new comment
    #[inline]
    pub fn comment(id: i32, author: Cow<'a, str>, text: Cow<'a, str>) -> Self {
        Self {
            annotation_type: AnnotationType::Comment,
            id,
            author,
            date: None,
            text,
            position: 0,
        }
    }

    /// Create a new revision mark
    #[inline]
    pub fn revision(id: i32, author: Cow<'a, str>) -> Self {
        Self {
            annotation_type: AnnotationType::Revision,
            id,
            author,
            date: None,
            text: Cow::Borrowed(""),
            position: 0,
        }
    }
}

/// Revision information (tracked change)
#[derive(Debug, Clone)]
pub struct Revision<'a> {
    /// Revision type
    pub revision_type: RevisionType,
    /// Author name
    pub author: Cow<'a, str>,
    /// Date of revision
    pub date: Option<Cow<'a, str>>,
    /// Revision ID
    pub id: i32,
    /// Text content affected by revision
    pub content: Cow<'a, str>,
}

impl<'a> Revision<'a> {
    /// Create a new revision
    #[inline]
    pub fn new(revision_type: RevisionType, author: Cow<'a, str>, content: Cow<'a, str>) -> Self {
        Self {
            revision_type,
            author,
            date: None,
            id: 0,
            content,
        }
    }

    /// Create an insertion revision
    #[inline]
    pub fn insertion(author: Cow<'a, str>, content: Cow<'a, str>) -> Self {
        Self::new(RevisionType::Insertion, author, content)
    }

    /// Create a deletion revision
    #[inline]
    pub fn deletion(author: Cow<'a, str>, content: Cow<'a, str>) -> Self {
        Self::new(RevisionType::Deletion, author, content)
    }
}
