//! Transactional persisted-record rewrite for PowerPoint OLE objects.
//!
//! [`Editor`] is intentionally kept as the small ergonomic facade.  Its
//! implementation is divided by responsibility so that opening an OLE
//! package, reading and mapping persisted records, staging mutations, and
//! emitting an incremental edit remain independently auditable.

mod lifecycle;
mod mapping;
mod mutation;
mod rewrite;
mod transaction;

#[cfg(test)]
mod tests;

use super::{Collection, ExternalObject};
use crate::embedded::storage::Storage;
use crate::package::Error;
use std::collections::{BTreeMap, HashMap, HashSet};
use std::sync::Arc;

type Result<T> = std::result::Result<T, Error>;

/// Appends a new PPT incremental edit; existing persisted bytes never move.
#[derive(Clone)]
pub struct Editor {
    pub(super) original: Arc<[u8]>,
    pub(super) max_output_bytes: usize,
    pub(super) streams: Vec<(Vec<String>, Vec<u8>)>,
    pub(super) document_path: Vec<String>,
    pub(super) current_user_path: Vec<String>,
    pub(super) document: Vec<u8>,
    pub(super) current_user: Vec<u8>,
    pub(super) mappings: BTreeMap<u32, u32>,
    pub(super) current_edit_offset: u32,
    pub(super) document_persist_id: u32,
    pub(super) collection: Collection,
    pub(super) staged_storage: HashMap<u32, Vec<u8>>,
    pub(super) removed_persist_ids: HashSet<u32>,
    pub(super) rewrite_object_list: bool,
    pub(super) changed: bool,
}

impl Editor {
    /// Resolve the exact live clear-text document through the current
    /// UserEdit/PersistDirectory chain without applying a mutation gate.
    #[allow(
        dead_code,
        reason = "used by capability-specific package snapshot paths"
    )]
    pub(crate) fn inspect_live_document(bytes: &[u8]) -> Result<(u32, Vec<u8>)> {
        lifecycle::inspect_live_document(bytes)
    }

    pub(crate) fn inspect_live_mapping(
        document: &[u8],
        current_user: &[u8],
    ) -> Result<crate::persist::PersistMapping> {
        lifecycle::inspect_live_mapping(document, current_user)
    }

    /// Opens a transactional editor over an OLE-backed PowerPoint package.
    pub fn open(bytes: Vec<u8>, collection: Collection) -> Result<Self> {
        lifecycle::open(Arc::from(bytes), collection)
    }

    /// Opens the shared incremental persisted-record editor without requiring
    /// an external-object collection. Used by non-OLE record editors.
    pub fn open_records(bytes: Vec<u8>) -> Result<Self> {
        lifecycle::open_records(Arc::from(bytes))
    }

    pub(crate) fn open_records_arc_with_limit(
        bytes: Arc<[u8]>,
        max_output_bytes: usize,
    ) -> Result<Self> {
        lifecycle::open::open_records_arc_with_limit(bytes, max_output_bytes)
    }

    /// Live persisted identifiers in ascending order.
    pub fn persist_ids(&self) -> Vec<u32> {
        lifecycle::persist_ids(self)
    }

    /// Returns one complete live persisted record.
    pub fn persisted_record(&self, persist_id: u32) -> Result<Vec<u8>> {
        lifecycle::persisted_record(self, persist_id)
    }

    /// Stages one complete replacement record in the next incremental edit.
    pub fn replace_persisted_record(&mut self, persist_id: u32, record: Vec<u8>) -> Result<()> {
        mutation::replace_persisted_record(self, persist_id, record)
    }

    /// Returns the typed embedded-object collection represented by this edit.
    pub fn objects(&self) -> &Collection {
        &self.collection
    }

    /// Stages a new external object and its persisted storage record.
    pub fn add(&mut self, object: ExternalObject, storage: Storage) -> Result<u32> {
        mutation::add(self, object, storage)
    }

    /// Stages a replacement for an existing external object's storage record.
    pub fn replace_storage(&mut self, persist_id: u32, storage: Storage) -> Result<()> {
        mutation::replace_storage(self, persist_id, storage)
    }

    /// Stages removal of one object reference and, when unreferenced, its
    /// persisted record.
    pub fn remove(&mut self, id: u32) -> Result<ExternalObject> {
        mutation::remove(self, id)
    }

    /// Stages a new object ordering for the external-object list.
    pub fn reorder(&mut self, ids: &[u32]) -> Result<()> {
        mutation::reorder(self, ids)
    }

    /// Emits the staged changes as one PPT incremental edit.
    pub fn finish(self) -> Result<Vec<u8>> {
        transaction::finish(self)
    }
}
