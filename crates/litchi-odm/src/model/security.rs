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
    /// An inert script or presentation event-listener action.
    EventListener,
}

/// One security-relevant inert item.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ActiveContent {
    pub(crate) kind: ActiveKind,
    pub(crate) location: String,
    pub(crate) trigger: Option<String>,
    pub(crate) action: Option<String>,
    pub(crate) target: Option<String>,
    pub(crate) link: Option<String>,
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

    /// Returns the authored event name, when present.
    #[must_use]
    pub fn trigger(&self) -> Option<&str> {
        self.trigger.as_deref()
    }

    /// Returns the inert presentation action token, when present.
    #[must_use]
    pub fn action(&self) -> Option<&str> {
        self.action.as_deref()
    }

    /// Returns the inert macro/function target, when present.
    #[must_use]
    pub fn target(&self) -> Option<&str> {
        self.target.as_deref()
    }

    /// Returns the inert listener link target, when present.
    #[must_use]
    pub fn link(&self) -> Option<&str> {
        self.link.as_deref()
    }
}

/// Intrinsic disposition for publishing changed package bytes.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum ChangedWriteDisposition {
    /// A normal changed write is intrinsically supported.
    Allowed,
    /// Active content requires explicit inert-preservation opt-in.
    RequiresInertActiveContentOptIn,
    /// A recognized signature prevents changed publication.
    RefusedSigned,
    /// Manifest encryption prevents changed publication.
    RefusedEncrypted,
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

    /// Returns the strongest disposition for a changed write.
    #[must_use]
    pub const fn changed_write_disposition(&self) -> ChangedWriteDisposition {
        if self.encrypted {
            ChangedWriteDisposition::RefusedEncrypted
        } else if self.signed {
            ChangedWriteDisposition::RefusedSigned
        } else if self.active_content.is_empty() {
            ChangedWriteDisposition::Allowed
        } else {
            ChangedWriteDisposition::RequiresInertActiveContentOptIn
        }
    }
}
