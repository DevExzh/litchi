//! Inert database form and report component declarations.

/// The user-facing database area that owns a component.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum ComponentKind {
    /// A form declared below `db:forms`.
    Form,
    /// A report declared below `db:reports`.
    Report,
}

/// A detached form or report component declaration.
///
/// Linked and embedded component payloads remain inert package resources.
/// Reading this value never follows its IRI or activates document content.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Component {
    kind: ComponentKind,
    name: Option<String>,
    title: Option<String>,
    description: Option<String>,
    href: Option<String>,
    as_template: Option<bool>,
}

impl Component {
    pub(crate) fn parsed(
        kind: ComponentKind,
        name: Option<String>,
        title: Option<String>,
        description: Option<String>,
        href: Option<String>,
        as_template: Option<bool>,
    ) -> Self {
        Self {
            kind,
            name,
            title,
            description,
            href,
            as_template,
        }
    }

    /// Returns whether this component is a form or report.
    #[must_use]
    pub const fn kind(&self) -> ComponentKind {
        self.kind
    }

    /// Returns its producer-visible name, if declared.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Returns its producer-visible title, if declared.
    #[must_use]
    pub fn title(&self) -> Option<&str> {
        self.title.as_deref()
    }

    /// Returns its inert description, if declared.
    #[must_use]
    pub fn description(&self) -> Option<&str> {
        self.description.as_deref()
    }

    /// Returns the inert linked package path or IRI, if declared.
    #[must_use]
    pub fn href(&self) -> Option<&str> {
        self.href.as_deref()
    }

    /// Returns whether the component is marked as a template, if declared.
    #[must_use]
    pub const fn as_template(&self) -> Option<bool> {
        self.as_template
    }
}
