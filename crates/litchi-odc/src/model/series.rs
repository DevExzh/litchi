//! Chart series selection semantics.

/// A semantic chart series selector.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Series {
    range: String,
}

impl Series {
    pub fn new(range: impl Into<String>) -> Self {
        Self {
            range: range.into(),
        }
    }

    pub fn range(&self) -> &str {
        &self.range
    }
}
