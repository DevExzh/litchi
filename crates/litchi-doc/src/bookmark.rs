/// A named bookmark in a legacy Word document.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Bookmark {
    /// Bookmark name.
    pub name: String,
    /// Absolute CP at which the bookmark begins.
    pub start: u32,
    /// Absolute exclusive-end CP.
    pub end: u32,
    /// Whether Word should retain the bookmark when exporting to RTF, HTML, or XML.
    pub is_native: bool,
    /// Optional zero-based table-column range `(first, exclusive_limit)`.
    pub column_range: Option<(u8, u8)>,
}

impl Bookmark {
    /// Whether this is a hidden bookmark.
    #[must_use]
    pub fn is_hidden(&self) -> bool {
        self.name.starts_with('_')
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn identifies_hidden_bookmarks() {
        let bookmark = Bookmark {
            name: "_GoBack".to_string(),
            start: 1,
            end: 2,
            is_native: true,
            column_range: None,
        };
        assert!(bookmark.is_hidden());
    }
}
