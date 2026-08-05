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

    pub fn href(&self) -> &str {
        &self.href
    }
}
