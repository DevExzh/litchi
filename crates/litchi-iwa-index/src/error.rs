use std::fmt;

use crate::{FragmentId, ObjectId, Reference};

/// The builder storage that could not reserve its next item.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum AllocationKind {
    /// The distinct fragment catalog.
    Fragments,
    /// The object-record collection.
    Objects,
    /// The reference collection.
    References,
    /// The temporary fragment duplicate catalog.
    FragmentCatalog,
    /// The temporary object duplicate catalog.
    ObjectCatalog,
    /// The temporary reference duplicate catalog.
    ReferenceCatalog,
    /// The temporary `(fragment, object)` ordering pairs assembled at build time.
    FragmentObjectPairs,
    /// The immutable object-identity storage grouped by fragment.
    FragmentObjectIds,
    /// The immutable fragment lookup entries assembled at build time.
    FragmentEntries,
}

impl fmt::Display for AllocationKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        let name = match self {
            Self::Fragments => "fragment storage",
            Self::Objects => "object storage",
            Self::References => "reference storage",
            Self::FragmentCatalog => "fragment duplicate catalog",
            Self::ObjectCatalog => "object duplicate catalog",
            Self::ReferenceCatalog => "reference duplicate catalog",
            Self::FragmentObjectPairs => "fragment/object ordering pairs",
            Self::FragmentObjectIds => "fragment object identity storage",
            Self::FragmentEntries => "fragment lookup entries",
        };
        formatter.write_str(name)
    }
}

/// Errors raised while assembling an immutable object index.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum IndexError {
    /// The builder could not reserve one of its record catalogs or derived
    /// snapshot tables.
    Allocation {
        /// The storage that failed to reserve.
        kind: AllocationKind,
        /// The item count the storage was required to accommodate.
        requested: usize,
    },
    /// A fragment identity was registered more than once.
    DuplicateFragment(FragmentId),
    /// An object identity was registered more than once.
    DuplicateObject(ObjectId),
    /// An object refers to a fragment that was not registered.
    UnknownFragment(FragmentId),
    /// A reference originates from an object that was not registered.
    UnknownSource(ObjectId),
    /// A reference targets an object that was not registered.
    UnknownTarget(ObjectId),
    /// A reference was registered more than once.
    DuplicateReference(Reference),
}

impl fmt::Display for IndexError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Allocation { kind, requested } => {
                write!(formatter, "could not reserve {kind} for {requested} items")
            },
            Self::DuplicateFragment(fragment) => {
                write!(formatter, "fragment {fragment:?} is already registered")
            },
            Self::DuplicateObject(object) => {
                write!(formatter, "object {object:?} is already registered")
            },
            Self::UnknownFragment(fragment) => {
                write!(formatter, "object refers to unknown fragment {fragment:?}")
            },
            Self::UnknownSource(object) => {
                write!(formatter, "reference source {object:?} is not registered")
            },
            Self::UnknownTarget(object) => {
                write!(formatter, "reference target {object:?} is not registered")
            },
            Self::DuplicateReference(reference) => {
                write!(formatter, "reference {reference:?} is already registered")
            },
        }
    }
}

impl std::error::Error for IndexError {}
