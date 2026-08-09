//! Inert package-resource semantics.

/// Resource-bearing content kind.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Kind {
    /// `draw:image`.
    Image,
    /// `draw:object`.
    Object,
    /// `draw:object-ole`.
    OleObject,
    /// `draw:plugin`.
    Plugin,
    /// `draw:floating-frame`.
    FloatingFrame,
}

/// An inert linked or package-relative resource reference.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Resource {
    href: String,
    kind: Kind,
}

impl Resource {
    pub(crate) const fn projected(kind: Kind, href: String) -> Self {
        Self { href, kind }
    }

    /// Resource family.
    #[must_use]
    pub const fn kind(&self) -> Kind {
        self.kind
    }

    /// Unresolved `xlink:href` value.
    #[must_use]
    pub fn href(&self) -> &str {
        &self.href
    }

    /// Whether the reference denotes a package member rather than an external URI.
    #[must_use]
    pub fn is_embedded(&self) -> bool {
        !self.href.is_empty()
            && !self.href.starts_with('#')
            && !self.href.contains("://")
            && !self.href.starts_with("data:")
    }
}
