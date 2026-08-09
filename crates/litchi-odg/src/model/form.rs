//! Inert drawing form-control declarations.

/// One bounded form element carrying a `form:id` inside `office:forms`.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Control {
    id: String,
    name: Option<String>,
    element: String,
}

impl Control {
    /// Creates a detached inert `form:control` declaration.
    #[must_use]
    pub fn new(id: impl Into<String>) -> Self {
        Self {
            id: id.into(),
            name: None,
            element: "control".to_string(),
        }
    }

    /// Sets the optional `form:name` on a detached declaration.
    #[must_use]
    pub fn with_name(mut self, name: impl Into<String>) -> Self {
        self.name = Some(name.into());
        self
    }

    pub(crate) fn parsed(id: String, name: Option<String>, element: String) -> Self {
        Self { id, name, element }
    }

    /// Exact inert `form:id` referenced by drawing control shapes.
    #[must_use]
    pub fn id(&self) -> &str {
        &self.id
    }

    /// Optional form name.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Source form element local name.
    #[must_use]
    pub fn element(&self) -> &str {
        &self.element
    }
}
