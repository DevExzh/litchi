//! Typed semantic reference-mark model.

/// A point or range target for `text:reference-ref` fields.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ReferenceMark {
    pub(super) name: String,
    pub(super) start: Option<(usize, usize)>,
    pub(super) end: Option<(usize, usize)>,
    pub(super) text: String,
    pub(super) range: bool,
}

impl ReferenceMark {
    /// Create a point reference target.
    pub fn point(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            start: None,
            end: None,
            text: String::new(),
            range: false,
        }
    }

    /// Create a range reference target.
    pub fn range(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            start: None,
            end: None,
            text: String::new(),
            range: true,
        }
    }

    pub fn name(&self) -> &str {
        &self.name
    }

    /// Zero-based paragraph/heading index and character offset.
    pub fn start(&self) -> Option<(usize, usize)> {
        self.start
    }

    /// Zero-based paragraph/heading index and character offset.
    pub fn end(&self) -> Option<(usize, usize)> {
        self.end
    }

    pub fn text(&self) -> &str {
        &self.text
    }

    pub fn is_range(&self) -> bool {
        self.range
    }
}
