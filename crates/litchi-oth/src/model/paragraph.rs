//! Web-template paragraph semantics.

use crate::link::Link;

/// A semantic web-template paragraph.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Paragraph {
    fields: Vec<crate::field::Field>,
    links: Vec<Link>,
    runs: Vec<crate::formatting::Run>,
    style_name: Option<String>,
    text: String,
}

impl Paragraph {
    /// Creates a detached plain paragraph.
    #[must_use]
    pub fn new(text: impl Into<String>) -> Self {
        Self {
            fields: Vec::new(),
            links: Vec::new(),
            runs: Vec::new(),
            style_name: None,
            text: text.into(),
        }
    }

    /// Creates a detached paragraph with a style reference.
    #[must_use]
    pub fn styled(text: impl Into<String>, style_name: impl Into<String>) -> Self {
        Self {
            fields: Vec::new(),
            links: Vec::new(),
            runs: Vec::new(),
            style_name: Some(style_name.into()),
            text: text.into(),
        }
    }

    pub(crate) const fn projected(
        text: String,
        style_name: Option<String>,
        links: Vec<Link>,
        runs: Vec<crate::formatting::Run>,
        fields: Vec<crate::field::Field>,
    ) -> Self {
        Self {
            fields,
            links,
            runs,
            style_name,
            text,
        }
    }

    /// Returns inert hyperlinks contained by the paragraph in document order.
    #[must_use]
    pub fn links(&self) -> &[Link] {
        &self.links
    }

    /// Returns character formatting ranges in source-close order.
    #[must_use]
    pub fn formatting_runs(&self) -> &[crate::formatting::Run] {
        &self.runs
    }

    /// Returns inert fields in source-close order.
    #[must_use]
    pub fn fields(&self) -> &[crate::field::Field] {
        &self.fields
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
