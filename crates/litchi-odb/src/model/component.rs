//! Inert database form and report component declarations.

/// The user-facing database area that owns a component.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum ComponentKind {
    /// A form declared below `db:forms`.
    Form,
    /// A report declared below `db:reports`.
    Report,
}

/// The inert linkage class of a component declaration.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ComponentLinkKind {
    /// No `xlink:href` is declared.
    Absent,
    /// A safe relative package subtree is declared.
    LocalPackage,
    /// An external, fragment, absolute, or otherwise non-package IRI is declared.
    ExternalIri,
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
    /// Creates an inert form or report component declaration.
    #[must_use]
    pub fn new(kind: ComponentKind, name: impl Into<String>) -> Self {
        Self {
            kind,
            name: Some(name.into()),
            title: None,
            description: None,
            href: None,
            as_template: None,
        }
    }

    /// Sets the producer-visible title.
    #[must_use]
    pub fn with_title(mut self, value: impl Into<String>) -> Self {
        self.title = Some(value.into());
        self
    }

    /// Sets the inert description.
    #[must_use]
    pub fn with_description(mut self, value: impl Into<String>) -> Self {
        self.description = Some(value.into());
        self
    }

    /// Sets the inert linked package path or IRI.
    #[must_use]
    pub fn with_href(mut self, value: impl Into<String>) -> Self {
        self.href = Some(value.into());
        self
    }

    /// Sets whether the component is a template.
    #[must_use]
    pub const fn with_as_template(mut self, value: Option<bool>) -> Self {
        self.as_template = value;
        self
    }

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

    /// Classifies the link without following, opening, or decoding it.
    #[must_use]
    pub fn link_kind(&self) -> ComponentLinkKind {
        let Some(href) = self.href() else {
            return ComponentLinkKind::Absent;
        };
        let path = href.trim_end_matches('/');
        if !path.is_empty()
            && !href.starts_with('/')
            && !href
                .chars()
                .any(|character| matches!(character, ':' | '\\' | '?' | '#'))
            && path
                .split('/')
                .all(|segment| !segment.is_empty() && !matches!(segment, "." | ".."))
        {
            ComponentLinkKind::LocalPackage
        } else {
            ComponentLinkKind::ExternalIri
        }
    }

    /// Returns whether the component is marked as a template, if declared.
    #[must_use]
    pub const fn as_template(&self) -> Option<bool> {
        self.as_template
    }
}
