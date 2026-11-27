use std::fmt;
use std::num::NonZeroU32;

/// A compact, non-null identity assigned to one physical document fragment.
///
/// The value is an adapter-local ordinal. It is intentionally not a package
/// entry name, archive identifier, or native object identifier. A concrete
/// reader may assign ordinals while traversing its package and retain its
/// private mapping until it emits neutral index records.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct FragmentId(NonZeroU32);

impl FragmentId {
    /// Construct an identity from a checked non-zero ordinal.
    #[must_use]
    pub const fn new(ordinal: NonZeroU32) -> Self {
        Self(ordinal)
    }

    /// Return the checked ordinal without converting it to an unvalidated
    /// primitive.
    #[must_use]
    pub const fn ordinal(self) -> NonZeroU32 {
        self.0
    }
}

impl TryFrom<u32> for FragmentId {
    type Error = FragmentIdError;

    fn try_from(ordinal: u32) -> Result<Self, Self::Error> {
        NonZeroU32::new(ordinal)
            .map(Self)
            .ok_or(FragmentIdError::Null)
    }
}

impl From<NonZeroU32> for FragmentId {
    fn from(ordinal: NonZeroU32) -> Self {
        Self::new(ordinal)
    }
}

/// A strict upper bound for object records visited in one fragment query.
///
/// The limit counts records, not bytes or native/archive identifiers. A zero
/// limit is useful for callers that want to assert that a fragment is empty;
/// every non-empty fragment is refused rather than truncated.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct FragmentTraversalLimit(usize);

impl FragmentTraversalLimit {
    /// Construct a record-count limit.
    #[must_use]
    pub const fn new(max_objects: usize) -> Self {
        Self(max_objects)
    }

    /// Return the maximum number of records a query may visit.
    #[must_use]
    pub const fn max_objects(self) -> usize {
        self.0
    }
}

/// Archive-free metadata for one indexed fragment.
///
/// A summary deliberately contains only the adapter-local fragment identity
/// and the number of neutral object records assigned to it. Archive entry
/// names, native identifiers, source handles, and payload bytes remain owned
/// by the concrete format adapter.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct FragmentSummary {
    id: FragmentId,
    object_count: usize,
}

impl FragmentSummary {
    pub(crate) const fn new(id: FragmentId, object_count: usize) -> Self {
        Self { id, object_count }
    }

    /// Return the adapter-local fragment identity.
    #[must_use]
    pub const fn id(self) -> FragmentId {
        self.id
    }

    /// Return the number of neutral object records in this fragment.
    #[must_use]
    pub const fn object_count(self) -> usize {
        self.object_count
    }
}

/// Failure while constructing a fragment identity.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FragmentIdError {
    /// Zero is reserved as the absent/null sentinel at native boundaries.
    Null,
}

impl fmt::Display for FragmentIdError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Null => formatter.write_str("fragment identity must be non-zero"),
        }
    }
}

impl std::error::Error for FragmentIdError {}
