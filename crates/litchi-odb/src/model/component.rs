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

/// The ownership role of one package member in a linked component closure.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ComponentDependencyKind {
    /// A file below the component's declared package subtree.
    PayloadFile,
    /// A package-local file reached through an XML `xlink:href` chain.
    LinkedFile,
    /// A manifest directory below the component subtree.
    PayloadDirectory,
    /// A manifest directory that owns a linked file dependency.
    LinkedDirectory,
}

/// Why exact component payload publication is unavailable.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ComponentTransferRefusal {
    /// A relocated XML member is producer-formatted and the shared writer
    /// cannot yet attach its donor provenance to a different package path.
    FormattedXmlRequiresSourceProvenance,
}

/// Whether the inventoried payload can pass the current audited writer.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Hash)]
pub enum ComponentTransferSupport {
    #[default]
    Supported,
    Refused(ComponentTransferRefusal),
}

/// One inert member in a component's bounded package dependency closure.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ComponentDependency {
    kind: ComponentDependencyKind,
    path: String,
    media_type: String,
    byte_len: Option<usize>,
}

impl ComponentDependency {
    pub(crate) fn new(
        kind: ComponentDependencyKind,
        path: String,
        media_type: String,
        byte_len: Option<usize>,
    ) -> Self {
        Self {
            kind,
            path,
            media_type,
            byte_len,
        }
    }

    #[must_use]
    pub const fn kind(&self) -> ComponentDependencyKind {
        self.kind
    }

    #[must_use]
    pub fn path(&self) -> &str {
        &self.path
    }

    #[must_use]
    pub fn media_type(&self) -> &str {
        &self.media_type
    }

    /// Returns the decoded package-member size; directories return `None`.
    #[must_use]
    pub const fn byte_len(&self) -> Option<usize> {
        self.byte_len
    }
}

/// Bounded inert package dependencies for one linked form or report.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct ComponentDependencyInventory {
    entries: Vec<ComponentDependency>,
    active_content_count: usize,
    transfer_support: ComponentTransferSupport,
}

impl ComponentDependencyInventory {
    pub(crate) const fn new(
        entries: Vec<ComponentDependency>,
        active_content_count: usize,
        transfer_support: ComponentTransferSupport,
    ) -> Self {
        Self {
            entries,
            active_content_count,
            transfer_support,
        }
    }

    #[must_use]
    pub fn entries(&self) -> &[ComponentDependency] {
        &self.entries
    }

    #[must_use]
    pub const fn active_content_count(&self) -> usize {
        self.active_content_count
    }

    /// Returns whether exact bytes can pass the audited package writer.
    #[must_use]
    pub const fn transfer_support(&self) -> ComponentTransferSupport {
        self.transfer_support
    }

    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }
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
