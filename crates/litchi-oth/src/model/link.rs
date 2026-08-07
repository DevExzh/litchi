//! Hyperlink semantics.

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

    /// Returns the hyperlink target.
    #[must_use]
    pub fn href(&self) -> &str {
        &self.href
    }

    /// Returns the hyperlink label.
    #[must_use]
    pub fn label(&self) -> &str {
        &self.label
    }
}
