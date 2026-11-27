//! Inert drawing form-control declarations.

use std::collections::BTreeMap;

/// One bounded form element carrying a `form:id` inside `office:forms`.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Control {
    id: String,
    name: Option<String>,
    element: String,
    attributes: BTreeMap<String, String>,
}

impl Control {
    /// Creates a detached inert `form:control` declaration.
    #[must_use]
    pub fn new(id: impl Into<String>) -> Self {
        Self {
            id: id.into(),
            name: None,
            element: "control".to_string(),
            attributes: BTreeMap::new(),
        }
    }

    /// Sets the optional `form:name` on a detached declaration.
    #[must_use]
    pub fn with_name(mut self, name: impl Into<String>) -> Self {
        self.name = Some(name.into());
        self
    }

    /// Selects an inert form element kind such as `button`, `checkbox`, or `listbox`.
    #[must_use]
    pub fn with_element(mut self, element: impl Into<String>) -> Self {
        self.element = element.into();
        self
    }

    /// Adds or replaces one inert qualified form attribute.
    #[must_use]
    pub fn with_attribute(mut self, name: impl Into<String>, value: impl Into<String>) -> Self {
        self.attributes.insert(name.into(), value.into());
        self
    }

    pub(crate) fn parsed(
        id: String,
        name: Option<String>,
        element: String,
        attributes: BTreeMap<String, String>,
    ) -> Self {
        Self {
            id,
            name,
            element,
            attributes,
        }
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

    /// Stable arbitrary qualified attributes retained without activation.
    #[must_use]
    pub const fn attributes(&self) -> &BTreeMap<String, String> {
        &self.attributes
    }
}
