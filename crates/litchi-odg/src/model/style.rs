//! Inert drawing style definitions and arbitrary property attributes.

use std::collections::BTreeMap;

/// One drawing style with bounded, inert graphic-property attributes.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Style {
    name: String,
    family: String,
    parent: Option<String>,
    properties: BTreeMap<String, String>,
}

impl Style {
    /// Creates a detached style definition.
    #[must_use]
    pub fn new(name: impl Into<String>, family: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            family: family.into(),
            parent: None,
            properties: BTreeMap::new(),
        }
    }

    /// Adds or replaces one inert qualified property attribute.
    #[must_use]
    pub fn with_property(mut self, name: impl Into<String>, value: impl Into<String>) -> Self {
        let owner = self.default_property_owner();
        self.properties
            .insert(format!("{owner}/{}", name.into()), value.into());
        self
    }

    /// Sets an optional parent-style dependency.
    #[must_use]
    pub fn with_parent(mut self, parent: impl Into<String>) -> Self {
        self.parent = Some(parent.into());
        self
    }

    pub(crate) fn parsed(
        name: String,
        family: String,
        parent: Option<String>,
        properties: BTreeMap<String, String>,
    ) -> Self {
        Self {
            name,
            family,
            parent,
            properties,
        }
    }

    /// Style name used by drawing references.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// ODF style family, such as `graphic`, `presentation`, or `drawing-page`.
    #[must_use]
    pub fn family(&self) -> &str {
        &self.family
    }

    /// Optional parent-style dependency.
    #[must_use]
    pub fn parent(&self) -> Option<&str> {
        self.parent.as_deref()
    }

    /// Stable qualified property attributes retained without interpretation.
    #[must_use]
    pub const fn properties(&self) -> &BTreeMap<String, String> {
        &self.properties
    }

    /// Looks up a qualified property attribute across the style's property owners.
    #[must_use]
    pub fn property(&self, name: &str) -> Option<&str> {
        self.properties
            .iter()
            .find(|(path, _value)| {
                path.rsplit_once('/')
                    .is_some_and(|(_owner, key)| key == name)
            })
            .map(|(_path, value)| value.as_str())
    }

    fn default_property_owner(&self) -> &'static str {
        match self.family.as_str() {
            "drawing-page" => "style:drawing-page-properties",
            "paragraph" => "style:paragraph-properties",
            "text" => "style:text-properties",
            _ => "style:graphic-properties",
        }
    }
}
