//! Lossless storage for bounded records outside the typed OLE vocabulary.

use crate::records::Record;

/// A bounded, lossless child of `ExObjList` that this crate does not model.
///
/// The record header and payload are retained so a typed OLE edit does not
/// discard unrelated media, hyperlink, or future-version records. The record
/// is exposed through borrowed accessors; callers never need to clone its
/// payload merely to inspect it.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct UnknownRecord {
    pub(crate) record: Record,
    pub(crate) object_index: usize,
}
