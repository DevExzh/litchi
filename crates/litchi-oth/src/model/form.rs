//! Inert form and control inventory.

/// One form control declaration.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Control {
    id: Option<String>,
    kind: String,
    name: Option<String>,
}

impl Control {
    pub(crate) const fn projected(kind: String, id: Option<String>, name: Option<String>) -> Self {
        Self { id, kind, name }
    }

    /// Form-namespace local element name.
    #[must_use]
    pub fn kind(&self) -> &str {
        &self.kind
    }

    /// Producer control identifier.
    #[must_use]
    pub fn id(&self) -> Option<&str> {
        self.id.as_deref()
    }

    /// Producer control name.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }
}

/// One form and its inert controls.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Form {
    controls: Vec<Control>,
    name: Option<String>,
}

impl Form {
    pub(crate) const fn projected(name: Option<String>, controls: Vec<Control>) -> Self {
        Self { controls, name }
    }

    /// Producer form name.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Controls in source order. They are never activated.
    #[must_use]
    pub fn controls(&self) -> &[Control] {
        &self.controls
    }
}
