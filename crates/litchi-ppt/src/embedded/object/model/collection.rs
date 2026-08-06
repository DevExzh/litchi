//! Ordered object-list snapshot model.

use super::{ExternalObject, UnknownRecord};

/// Strict embedded and linked OLE definitions in document order.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Collection {
    pub id_seed: u32,
    pub objects: Vec<ExternalObject>,
    pub(crate) unknown_records: Vec<UnknownRecord>,
}
