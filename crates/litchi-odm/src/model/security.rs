//! Inert security-relevant package inventory.

/// A kind of active content which is preserved but never executed.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum ActiveKind {
    /// An `office:dde-source` relationship.
    Dde,
    /// Script or macro XML.
    Script,
    /// A package entry under a conventional macro/script directory.
    ScriptResource,
    /// Form controls or forms declared in document XML.
    FormControl,
}

/// One security-relevant inert item.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ActiveContent {
    pub(crate) kind: ActiveKind,
    pub(crate) location: String,
}

impl ActiveContent {
    /// Returns the classified content kind.
    #[must_use]
    pub const fn kind(&self) -> ActiveKind {
        self.kind
    }

    /// Returns the package part containing the declaration.
    #[must_use]
    pub fn location(&self) -> &str {
        &self.location
    }
}

/// Immutable security state projected when the package is opened.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct State {
    pub(crate) signed: bool,
    pub(crate) encrypted: bool,
    pub(crate) active_content: Vec<ActiveContent>,
}

impl State {
    /// Reports whether the package contains a recognized ODF signature part.
    #[must_use]
    pub const fn is_signed(&self) -> bool {
        self.signed
    }

    /// Reports whether the manifest declares encrypted entries.
    #[must_use]
    pub const fn is_encrypted(&self) -> bool {
        self.encrypted
    }

    /// Returns inert active-content declarations and resources.
    #[must_use]
    pub fn active_content(&self) -> &[ActiveContent] {
        &self.active_content
    }
}
