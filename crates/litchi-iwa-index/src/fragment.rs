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
