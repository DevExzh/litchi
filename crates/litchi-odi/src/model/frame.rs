//! Image-frame semantics.

use super::source::Source;

/// A semantic image frame.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Frame {
    name: Option<String>,
    source: Source,
}

impl Frame {
    pub fn new(source: Source) -> Self {
        Self { name: None, source }
    }

    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    pub fn source(&self) -> &Source {
        &self.source
    }

    pub fn with_name(mut self, name: impl Into<String>) -> Self {
        self.name = Some(name.into());
        self
    }
}
