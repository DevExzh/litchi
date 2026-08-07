//! Drawing-layer semantics.

/// A semantic drawing layer.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Layer {
    name: String,
}

impl Layer {
    pub fn new(name: impl Into<String>) -> Self {
        Self { name: name.into() }
    }

    /// Returns the layer name.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }
}
