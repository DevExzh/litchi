//! Master-document section semantics.

use super::subdocument::Subdocument;

/// A master-document section.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Section {
    name: String,
    children: Vec<Subdocument>,
}

impl Section {
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            children: Vec::new(),
        }
    }

    /// Returns the section name.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Returns the subdocuments contained in the section.
    #[must_use]
    pub fn children(&self) -> &[Subdocument] {
        &self.children
    }

    pub fn push(&mut self, child: Subdocument) {
        self.children.push(child);
    }
}
