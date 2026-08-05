//! Drawing-page semantics.

/// A semantic drawing page.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Page {
    name: String,
}

impl Page {
    pub fn new(name: impl Into<String>) -> Self {
        Self { name: name.into() }
    }

    pub fn name(&self) -> &str {
        &self.name
    }
}
