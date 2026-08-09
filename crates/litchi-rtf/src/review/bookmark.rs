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
    /// UTF-8 byte offset in the document body text where the bookmark starts.
    pub position: usize,
    /// Body text covered by the bookmark.
    pub content: Cow<'a, str>,
    /// First table column covered by the bookmark, when specified.
    pub first_column: Option<i32>,
    /// Last table column covered by the bookmark, when specified.
    pub last_column: Option<i32>,
    /// Whether this is a public bookmark.
    pub is_public: bool,
}

impl<'a> Bookmark<'a> {
    /// Create a new bookmark
    #[inline]
    #[must_use]
    pub fn new(name: Cow<'a, str>) -> Self {
        Self {
            name,
            position: 0,
            content: Cow::Borrowed(""),
            first_column: None,
            last_column: None,
            is_public: false,
        }
    }

    /// Create a bookmark with content
    #[inline]
    #[must_use]
    pub fn with_content(name: Cow<'a, str>, content: Cow<'a, str>) -> Self {
        Self {
            name,
            position: 0,
            content,
            first_column: None,
            last_column: None,
            is_public: false,
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
    #[must_use]
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
    #[must_use]
    pub fn get(&self, name: &str) -> Option<&Bookmark<'a>> {
        self.bookmarks.iter().find(|b| b.name.as_ref() == name)
    }

    /// Get all bookmarks
    #[inline]
    #[must_use]
    pub fn bookmarks(&self) -> &[Bookmark<'a>] {
        &self.bookmarks
    }

    /// Check if a bookmark exists
    #[inline]
    #[must_use]
    pub fn contains(&self, name: &str) -> bool {
        self.get(name).is_some()
    }
}
