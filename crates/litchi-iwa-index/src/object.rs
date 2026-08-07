use litchi_iwa_graph::ObjectId;

use crate::{ByteSpan, FragmentId};

/// Neutral location metadata for one validated object.
///
/// The record contains no payload and no archive/package handle. Unknown
/// object bytes therefore remain owned by the concrete adapter and are not
/// discarded by building an index.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct ObjectRecord {
    id: ObjectId,
    fragment: FragmentId,
    span: ByteSpan,
}

impl ObjectRecord {
    /// Construct a location record from already validated typed values.
    #[must_use]
    pub const fn new(id: ObjectId, fragment: FragmentId, span: ByteSpan) -> Self {
        Self { id, fragment, span }
    }

    /// Return the object's validated identity.
    #[must_use]
    pub const fn id(&self) -> ObjectId {
        self.id
    }

    /// Return the fragment containing the object.
    #[must_use]
    pub const fn fragment(&self) -> FragmentId {
        self.fragment
    }

    /// Return the object's checked byte location within its fragment.
    #[must_use]
    pub const fn span(&self) -> ByteSpan {
        self.span
    }
}
