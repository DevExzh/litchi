//! Ordered object-list snapshot model.

use super::{ExternalObject, UnknownRecord};

/// Strict embedded and linked OLE definitions in document order.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Collection {
    pub id_seed: u32,
    pub objects: Vec<ExternalObject>,
    pub(crate) unknown_records: Vec<UnknownRecord>,
}

impl Collection {
    /// The empty list a document without an `ExObjListContainer` starts
    /// from: no objects and the minimum identifier seed, 1 (`[MS-PPT]`
    /// `ExObjListAtom.exObjIdSeed` is positive).
    pub(crate) const fn empty() -> Self {
        Self {
            id_seed: 1,
            objects: Vec::new(),
            unknown_records: Vec::new(),
        }
    }
}
