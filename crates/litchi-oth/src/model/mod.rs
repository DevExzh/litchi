//! Immutable semantic values for this document family.

/// A semantic web-template paragraph.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Paragraph {
    text: String,
}
impl Paragraph {
    pub fn new(text: impl Into<String>) -> Self {
        Self { text: text.into() }
    }
    pub fn text(&self) -> &str {
        &self.text
    }
}
/// A hyperlink target and label.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Link {
    href: String,
    label: String,
}
impl Link {
    pub fn new(href: impl Into<String>, label: impl Into<String>) -> Self {
        Self {
            href: href.into(),
            label: label.into(),
        }
    }
    pub fn href(&self) -> &str {
        &self.href
    }
    pub fn label(&self) -> &str {
        &self.label
    }
}
