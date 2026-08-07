//! Referenced subdocument semantics.

/// A referenced master-document subdocument.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Subdocument {
    href: String,
}

impl Subdocument {
    pub fn new(href: impl Into<String>) -> Self {
        Self { href: href.into() }
    }

    /// Returns the subdocument reference target.
    #[must_use]
    pub fn href(&self) -> &str {
        &self.href
    }
}
