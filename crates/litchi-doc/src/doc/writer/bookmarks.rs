//! Standard bookmark input types for the legacy Word writer.

/// A named bookmark to add to a DOC file.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BookmarkEntry {
    /// Bookmark name, containing 1 through 39 UTF-16 code units.
    pub name: String,
    /// Absolute start CP in the emitted set of document parts.
    pub start: u32,
    /// Absolute exclusive-end CP.
    pub end: u32,
    /// Whether exports to RTF, HTML, or XML should retain this bookmark.
    pub is_native: bool,
    /// Optional zero-based table-column range `(first, exclusive_limit)`.
    pub column_range: Option<(u8, u8)>,
}

impl BookmarkEntry {
    /// Create a standard bookmark.
    pub fn new(name: impl Into<String>, start: u32, end: u32) -> Self {
        Self {
            name: name.into(),
            start,
            end,
            is_native: true,
            column_range: None,
        }
    }

    /// Set whether non-DOC exports should retain this bookmark.
    pub fn with_native_export(mut self, is_native: bool) -> Self {
        self.is_native = is_native;
        self
    }

    /// Restrict this bookmark to a table-column range.
    pub fn with_column_range(mut self, first: u8, exclusive_limit: u8) -> Self {
        self.column_range = Some((first, exclusive_limit));
        self
    }
}
