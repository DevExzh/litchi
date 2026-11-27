//! Inert form and control inventory.

/// One form control declaration.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Control {
    id: Option<String>,
    kind: String,
    name: Option<String>,
}

impl Control {
    /// Creates a detached inert control declaration.
    #[must_use]
    pub fn new(kind: impl Into<String>) -> Self {
        Self {
            id: None,
            kind: kind.into(),
            name: None,
        }
    }

    /// Sets the control identifier.
    #[must_use]
    pub fn with_id(mut self, id: impl Into<String>) -> Self {
        self.id = Some(id.into());
        self
    }

    /// Sets the producer-visible control name.
    #[must_use]
    pub fn with_name(mut self, name: impl Into<String>) -> Self {
        self.name = Some(name.into());
        self
    }

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
    /// Creates a detached named inert form.
    #[must_use]
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            controls: Vec::new(),
            name: Some(name.into()),
        }
    }

    /// Creates a detached anonymous inert form.
    #[must_use]
    pub const fn anonymous() -> Self {
        Self {
            controls: Vec::new(),
            name: None,
        }
    }

    /// Adds one inert control.
    #[must_use]
    pub fn with_control(mut self, control: Control) -> Self {
        self.controls.push(control);
        self
    }

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
