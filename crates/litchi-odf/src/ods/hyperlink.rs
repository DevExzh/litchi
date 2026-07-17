//! Cell hyperlink (`text:a`) data structures for ODS spreadsheets.
//!
//! ODF stores spreadsheet hyperlinks as `text:a` elements inside the
//! paragraph content of a table cell (ODF 1.3 §6.1.8). The target IRI is
//! carried by the mandatory `xlink:href` attribute, while the visible link
//! text is the character content of the element.

/// A hyperlink parsed from a `text:a` element inside cell content.
///
/// The hyperlink is inert metadata: the target IRI is preserved verbatim and
/// is never dereferenced by this crate.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct CellHyperlink {
    /// Target IRI from the mandatory `xlink:href` attribute.
    pub href: String,
    /// Plain text of the link (character content of the `text:a` subtree).
    pub text: String,
    /// Optional `office:name` attribute naming the hyperlink.
    pub name: Option<String>,
    /// Optional `office:title` attribute with a short accessible title.
    pub title: Option<String>,
    /// Optional `office:target-frame-name` attribute.
    pub target_frame_name: Option<String>,
    /// Optional `text:style-name` applied to the unvisited link.
    pub style_name: Option<String>,
    /// Optional `text:visited-style-name` applied to the visited link.
    pub visited_style_name: Option<String>,
}

impl CellHyperlink {
    /// Create a hyperlink with the mandatory target IRI and empty metadata.
    pub fn new(href: impl Into<String>) -> Self {
        Self {
            href: href.into(),
            text: String::new(),
            name: None,
            title: None,
            target_frame_name: None,
            style_name: None,
            visited_style_name: None,
        }
    }

    /// The target IRI of the hyperlink.
    pub fn href(&self) -> &str {
        &self.href
    }

    /// The visible plain-text content of the hyperlink.
    pub fn text(&self) -> &str {
        &self.text
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn new_hyperlink_keeps_href_and_defaults_metadata() {
        let link = CellHyperlink::new("https://example.com/");
        assert_eq!(link.href(), "https://example.com/");
        assert_eq!(link.text(), "");
        assert!(link.name.is_none());
        assert!(link.title.is_none());
        assert!(link.target_frame_name.is_none());
        assert!(link.style_name.is_none());
        assert!(link.visited_style_name.is_none());
    }
}
