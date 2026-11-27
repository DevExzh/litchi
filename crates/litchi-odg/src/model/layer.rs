//! Drawing-layer semantics.

/// A semantic drawing layer.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Layer {
    name: String,
    display: Option<String>,
    protected: Option<bool>,
}

impl Layer {
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            display: None,
            protected: None,
        }
    }

    /// Sets the optional display policy on a detached layer.
    #[must_use]
    pub fn with_display(mut self, display: impl Into<String>) -> Self {
        self.display = Some(display.into());
        self
    }

    /// Sets the optional protection flag on a detached layer.
    #[must_use]
    pub const fn with_protected(mut self, protected: bool) -> Self {
        self.protected = Some(protected);
        self
    }

    pub(crate) fn parsed(name: String, display: Option<String>, protected: Option<bool>) -> Self {
        Self {
            name,
            display,
            protected,
        }
    }

    /// Returns the layer name.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Returns the optional lexical `draw:display` policy.
    #[must_use]
    pub fn display(&self) -> Option<&str> {
        self.display.as_deref()
    }

    /// Returns whether the layer is protected, when explicitly declared.
    #[must_use]
    pub const fn protected(&self) -> Option<bool> {
        self.protected
    }
}
