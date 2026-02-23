/// Hyperlink structure for Word documents
use super::parts::hyperlinks::{Hyperlink as InternalHyperlink, HyperlinkType};

/// A hyperlink in a Word document
#[derive(Debug, Clone)]
pub struct Hyperlink {
    /// Character position range
    pub start_position: u32,
    pub end_position: u32,
    /// Destination URL/path
    pub destination: String,
    /// Display text
    pub display_text: String,
    /// Link type
    pub link_type: HyperlinkType,
}

impl Hyperlink {
    /// Create from internal hyperlink
    pub(crate) fn from_internal(internal: &InternalHyperlink) -> Self {
        Self {
            start_position: internal.start_cp,
            end_position: internal.end_cp,
            destination: internal.destination.clone(),
            display_text: internal.display_text.clone(),
            link_type: internal.link_type.clone(),
        }
    }

    /// Get the destination
    pub fn destination(&self) -> &str {
        &self.destination
    }

    /// Get the display text
    pub fn display_text(&self) -> &str {
        &self.display_text
    }

    /// Check if this is a URL hyperlink
    pub fn is_url(&self) -> bool {
        self.link_type == HyperlinkType::Url
    }

    /// Check if this is an email hyperlink
    pub fn is_email(&self) -> bool {
        self.link_type == HyperlinkType::Email
    }

    /// Check if this is a file hyperlink
    pub fn is_file(&self) -> bool {
        self.link_type == HyperlinkType::File
    }

    /// Check if this is a bookmark hyperlink
    pub fn is_bookmark(&self) -> bool {
        self.link_type == HyperlinkType::Bookmark
    }
}
