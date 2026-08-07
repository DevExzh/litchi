//! Image-frame semantics.

use super::source::Source;

/// A semantic image frame.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Frame {
    name: Option<String>,
    source: Source,
}

impl Frame {
    /// Creates a frame for the given image source without a name.
    #[must_use]
    pub fn new(source: Source) -> Self {
        Self { name: None, source }
    }

    /// Returns the frame name, if set.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Returns the image payload source.
    #[must_use]
    pub fn source(&self) -> &Source {
        &self.source
    }

    /// Sets the frame name.
    #[must_use]
    pub fn with_name(mut self, name: impl Into<String>) -> Self {
        self.name = Some(name.into());
        self
    }
}
