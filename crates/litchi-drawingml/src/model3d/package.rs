//! Host-neutral relationship graph seam for `model3d`.

use super::Id;

/// The target mode of a package relationship.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Target<'a> {
    /// A package-local target, represented by the host's canonical target text.
    Internal(&'a str),
    /// An external target retained inertly and never resolved by `DrawingML`.
    External(&'a str),
}

impl<'a> Target<'a> {
    /// Return whether this target is external.
    #[inline]
    #[must_use]
    pub const fn is_external(self) -> bool {
        matches!(self, Self::External(_))
    }

    /// Borrow the target text.
    #[inline]
    #[must_use]
    pub const fn as_str(self) -> &'a str {
        match self {
            Self::Internal(value) | Self::External(value) => value,
        }
    }
}

/// One resolved relationship occurrence supplied by a concrete package owner.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Relationship<'a> {
    /// The host relationship type URI, retained for owner-specific policy checks.
    pub relationship_type: &'a str,
    /// The internal or external target mode and target text.
    pub target: Target<'a>,
}

/// Read-only relationship lookup implemented by DOCX/PPTX/XLSX/XLSB adapters.
///
/// The shared `DrawingML` crate intentionally does not depend on OPC.  A host
/// adapter can therefore resolve a relationship ID against its own package
/// graph and still use the same semantic validation rules.
pub trait Resolver {
    /// Resolve one relationship ID from the owning model3d part.
    fn relationship<'a>(&'a self, id: &Id) -> Option<Relationship<'a>>;
}
