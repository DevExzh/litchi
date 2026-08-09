//! Web-template paragraph semantics.

use crate::link::Link;

/// A semantic web-template paragraph.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Paragraph {
    links: Vec<Link>,
    style_name: Option<String>,
    text: String,
}

impl Paragraph {
    /// Creates a detached plain paragraph.
    #[must_use]
    pub fn new(text: impl Into<String>) -> Self {
        Self {
            links: Vec::new(),
            style_name: None,
            text: text.into(),
        }
    }

    /// Creates a detached paragraph with a style reference.
    #[must_use]
    pub fn styled(text: impl Into<String>, style_name: impl Into<String>) -> Self {
        Self {
            links: Vec::new(),
            style_name: Some(style_name.into()),
            text: text.into(),
        }
    }

    pub(crate) const fn projected(
        text: String,
        style_name: Option<String>,
        links: Vec<Link>,
    ) -> Self {
        Self {
            links,
            style_name,
            text,
        }
    }

    /// Returns inert hyperlinks contained by the paragraph in document order.
    #[must_use]
    pub fn links(&self) -> &[Link] {
        &self.links
    }

    /// Returns the referenced paragraph style name, if present.
    #[must_use]
    pub fn style_name(&self) -> Option<&str> {
        self.style_name.as_deref()
    }

    /// Returns projected character data, including ODF whitespace elements.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }
}
