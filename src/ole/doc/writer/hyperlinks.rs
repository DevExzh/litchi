//! Hyperlinks writer for DOC files
//!
//! Generates HYPERLINK field codes for embedding links in documents.

/// Hyperlink type
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HyperlinkType {
    /// URL hyperlink
    Url,
    /// Email hyperlink
    Email,
    /// File hyperlink
    File,
    /// Bookmark hyperlink
    Bookmark,
}

/// A hyperlink entry
#[derive(Debug, Clone)]
pub struct HyperlinkEntry {
    /// Character position where hyperlink starts
    pub start_position: u32,
    /// Destination URL/path
    pub destination: String,
    /// Display text
    pub display_text: String,
    /// Link type
    pub link_type: HyperlinkType,
}

impl HyperlinkEntry {
    /// Create a new URL hyperlink
    pub fn url(
        start_position: u32,
        url: impl Into<String>,
        display_text: impl Into<String>,
    ) -> Self {
        Self {
            start_position,
            destination: url.into(),
            display_text: display_text.into(),
            link_type: HyperlinkType::Url,
        }
    }

    /// Create a new email hyperlink
    pub fn email(
        start_position: u32,
        email: impl Into<String>,
        display_text: impl Into<String>,
    ) -> Self {
        Self {
            start_position,
            destination: email.into(),
            display_text: display_text.into(),
            link_type: HyperlinkType::Email,
        }
    }

    /// Create a new bookmark hyperlink
    pub fn bookmark(
        start_position: u32,
        bookmark: impl Into<String>,
        display_text: impl Into<String>,
    ) -> Self {
        Self {
            start_position,
            destination: bookmark.into(),
            display_text: display_text.into(),
            link_type: HyperlinkType::Bookmark,
        }
    }

    /// Generate HYPERLINK field code
    pub fn to_field_code(&self) -> String {
        match self.link_type {
            HyperlinkType::Url => format!("HYPERLINK \"{}\"", self.destination),
            HyperlinkType::Email => format!("HYPERLINK \"mailto:{}\"", self.destination),
            HyperlinkType::File => format!("HYPERLINK \"{}\"", self.destination),
            HyperlinkType::Bookmark => format!("HYPERLINK \\l \"{}\"", self.destination),
        }
    }
}

/// Hyperlinks writer
#[derive(Debug)]
pub struct HyperlinksWriter {
    hyperlinks: Vec<HyperlinkEntry>,
}

impl HyperlinksWriter {
    /// Create a new hyperlinks writer
    pub fn new() -> Self {
        Self {
            hyperlinks: Vec::new(),
        }
    }

    /// Add a hyperlink
    pub fn add_hyperlink(&mut self, hyperlink: HyperlinkEntry) {
        self.hyperlinks.push(hyperlink);
    }

    /// Get all hyperlinks
    pub fn hyperlinks(&self) -> &[HyperlinkEntry] {
        &self.hyperlinks
    }

    /// Check if empty
    pub fn is_empty(&self) -> bool {
        self.hyperlinks.is_empty()
    }

    /// Get hyperlinks sorted by position
    pub fn hyperlinks_sorted(&self) -> Vec<&HyperlinkEntry> {
        let mut sorted: Vec<_> = self.hyperlinks.iter().collect();
        sorted.sort_by_key(|h| h.start_position);
        sorted
    }
}

impl Default for HyperlinksWriter {
    fn default() -> Self {
        Self::new()
    }
}
