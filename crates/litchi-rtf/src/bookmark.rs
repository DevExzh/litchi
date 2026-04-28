//! RTF bookmark support.
//!
//! Bookmarks are named locations in a document that can be referenced
//! by hyperlinks or cross-references.

use std::borrow::Cow;

/// A bookmark in an RTF document
#[derive(Debug, Clone)]
pub struct Bookmark<'a> {
    /// Bookmark name (unique identifier)
    pub name: Cow<'a, str>,
    /// Text position/offset where bookmark is located
    pub position: usize,
    /// Associated text content (if any)
    pub content: Cow<'a, str>,
}

impl<'a> Bookmark<'a> {
    /// Create a new bookmark
    #[inline]
    pub fn new(name: Cow<'a, str>) -> Self {
        Self {
            name,
            position: 0,
            content: Cow::Borrowed(""),
        }
    }

    /// Create a bookmark with content
    #[inline]
    pub fn with_content(name: Cow<'a, str>, content: Cow<'a, str>) -> Self {
        Self {
            name,
            position: 0,
            content,
        }
    }
}

/// Bookmark table for managing all bookmarks in a document
#[derive(Debug, Clone, Default)]
pub struct BookmarkTable<'a> {
    /// All bookmarks in the document
    bookmarks: Vec<Bookmark<'a>>,
}

impl<'a> BookmarkTable<'a> {
    /// Create a new bookmark table
    #[inline]
    pub fn new() -> Self {
        Self {
            bookmarks: Vec::new(),
        }
    }

    /// Add a bookmark
    #[inline]
    pub fn add(&mut self, bookmark: Bookmark<'a>) {
        self.bookmarks.push(bookmark);
    }

    /// Get a bookmark by name
    pub fn get(&self, name: &str) -> Option<&Bookmark<'a>> {
        self.bookmarks.iter().find(|b| b.name.as_ref() == name)
    }

    /// Get all bookmarks
    #[inline]
    pub fn bookmarks(&self) -> &[Bookmark<'a>] {
        &self.bookmarks
    }

    /// Check if a bookmark exists
    #[inline]
    pub fn contains(&self, name: &str) -> bool {
        self.get(name).is_some()
    }
}
