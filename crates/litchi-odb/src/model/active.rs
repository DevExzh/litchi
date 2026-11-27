//! Inert inventory of package content that could become active in a database host.

/// A bounded class of potentially active ODB content.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ActiveContentKind {
    /// A member below the package's `Basic/` macro tree.
    BasicMacro,
    /// A member below `Scripts/` or an XML script declaration.
    Script,
    /// An XML event-listener declaration.
    EventListener,
    /// An XML form control declaration.
    FormControl,
    /// An XML action declaration.
    Action,
    /// An XML DDE declaration.
    DdeLink,
    /// An embedded object, plug-in, or applet declaration.
    EmbeddedObject,
}

/// One inert, source-located active-content finding.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ActiveContentEntry {
    kind: ActiveContentKind,
    package_path: String,
    declaration: Option<String>,
}

impl ActiveContentEntry {
    pub(crate) fn package_member(kind: ActiveContentKind, package_path: String) -> Self {
        Self {
            kind,
            package_path,
            declaration: None,
        }
    }

    pub(crate) fn declaration(
        kind: ActiveContentKind,
        package_path: String,
        declaration: String,
    ) -> Self {
        Self {
            kind,
            package_path,
            declaration: Some(declaration),
        }
    }

    /// Returns the finding class.
    #[must_use]
    pub const fn kind(&self) -> ActiveContentKind {
        self.kind
    }

    /// Returns the inert package member in which the finding occurs.
    #[must_use]
    pub fn package_path(&self) -> &str {
        &self.package_path
    }

    /// Returns the XML local name for a declaration finding.
    #[must_use]
    pub fn declaration_name(&self) -> Option<&str> {
        self.declaration.as_deref()
    }
}

/// A bounded active-content inventory. Producing it never activates a member.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct ActiveContentInventory {
    entries: Vec<ActiveContentEntry>,
}

impl ActiveContentInventory {
    pub(crate) fn new(entries: Vec<ActiveContentEntry>) -> Self {
        Self { entries }
    }

    /// Returns all findings in deterministic package and document order.
    #[must_use]
    pub fn entries(&self) -> &[ActiveContentEntry] {
        &self.entries
    }

    /// Returns whether any potentially active content was found.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }
}
