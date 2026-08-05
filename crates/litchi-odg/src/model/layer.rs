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

    pub fn name(&self) -> &str {
        &self.name
    }
}
