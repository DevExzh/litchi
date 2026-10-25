//! Append-only persisted-record rewrite for `PowerPoint` packages.
//!
//! [`Editor`] is the crate-private record publisher behind
//! [`super::Edit`] and the other scoped families.  Its
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
use crate::package::{Error, SourceLimits};
#[cfg(test)]
use litchi_cfb::OleFile;
use std::collections::{BTreeMap, HashSet};
#[cfg(test)]
use std::io::{Read, Seek};
use std::sync::Arc;

type Result<T> = std::result::Result<T, Error>;

/// Appends a new PPT incremental edit; existing persisted bytes never move.
#[derive(Clone)]
pub(crate) struct Editor {
    pub(super) original: Arc<[u8]>,
    pub(super) max_output_bytes: usize,
    /// Source-level record limits selected when the owning package was opened.
    /// Object-list serialization must use this same profile; it may not fall
    /// back to a fresh parser default during publication.
    pub(super) source_limits: SourceLimits,
    /// Physical CFB bounds retained from the source admission. Every
    /// source-backed rewrite and candidate reopen must use this same profile.
    pub(super) cfb_limits: litchi_cfb::OleFileLimits,
    /// Every stream in source order. Entries whose path is `document_path` or
    /// `current_user_path` carry no payload: `document` and `current_user`
    /// own those bytes, and `finish` emits them from there.
    pub(super) streams: Vec<(Vec<String>, Vec<u8>)>,
    pub(super) document_path: Vec<String>,
    pub(super) current_user_path: Vec<String>,
    pub(super) document: Vec<u8>,
    pub(super) current_user: Vec<u8>,
    pub(super) mappings: BTreeMap<u32, u32>,
    pub(super) current_edit_offset: u32,
    pub(super) document_persist_id: u32,
    pub(super) collection: Collection,
    pub(super) staged_storage: BTreeMap<u32, Vec<u8>>,
    pub(super) removed_persist_ids: HashSet<u32>,
    pub(super) rewrite_object_list: bool,
    pub(super) changed: bool,
    pub(super) layout: litchi_cfb::SectorLayoutPolicy,
    /// Whether the source carries a digital-signature component. Such an
    /// editor was opened for staging only: [`Self::finish`] refuses any
    /// changed publication, which would invalidate the signature.
    pub(super) signed: bool,
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

    #[cfg(test)]
    pub(crate) fn inspect_live_document_from_ole<R: Read + Seek>(
        ole: &mut OleFile<R>,
    ) -> Result<(u32, Vec<u8>)> {
        lifecycle::open::inspect_live_document_from_ole(ole)
    }

    #[cfg(test)]
    pub(crate) fn inspect_live_mapping(
        document: &[u8],
        current_user: &[u8],
    ) -> Result<crate::persist::PersistMapping> {
        lifecycle::inspect_live_mapping(document, current_user)
    }

    /// Read the live persist mapping from the `UserEdit` chain whose newest
    /// atom sits at `head_offset` (see
    /// `crate::slide::directory::locate_head_user_edit`), repairing
    /// persist-directory defects into `report` (finding F5.7) and, when
    /// `leniency` opts in, breaking a cyclic chain (finding F5.9).
    pub(crate) fn inspect_live_mapping_at(
        document: &[u8],
        head_offset: u32,
        leniency: crate::leniency::Leniency,
        report: &mut crate::leniency::ToleranceReport,
    ) -> Result<crate::persist::PersistMapping> {
        lifecycle::inspect_live_mapping_at(document, head_offset, leniency, report)
    }

    /// Opens the shared incremental persisted-record editor without requiring
    /// an external-object collection. Used by non-OLE record editors.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub(crate) fn open_records(bytes: Vec<u8>, source_limits: SourceLimits) -> Result<Self> {
        lifecycle::open_records(Arc::from(bytes), source_limits)
    }

    /// Opens a transactional editor over an OLE-backed `PowerPoint` package
    /// that may carry a digital-signature component, so a scoped family can
    /// stage and read over a signed deck.
    /// [`Self::finish`] still refuses a changed publication of a signed
    /// source; encryption and rights-management components stay refused.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub(crate) fn open_admitting_signature(
        bytes: Vec<u8>,
        collection: Collection,
        cfb_limits: litchi_cfb::OleFileLimits,
        source_limits: SourceLimits,
    ) -> Result<Self> {
        lifecycle::open::open_admitting_signature(
            Arc::from(bytes),
            collection,
            cfb_limits,
            source_limits,
        )
    }

    /// [`Self::open_records`] over a package that may carry a
    /// digital-signature component (see [`Self::open_admitting_signature`]).
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub(crate) fn open_records_admitting_signature(
        bytes: Vec<u8>,
        source_limits: SourceLimits,
    ) -> Result<Self> {
        lifecycle::open::open_admitting_signature(
            Arc::from(bytes),
            lifecycle::open::no_objects(),
            litchi_cfb::OleFileLimits::default(),
            source_limits,
        )
    }

    pub(crate) fn open_records_arc_with_limit(
        bytes: Arc<[u8]>,
        max_output_bytes: usize,
        source_limits: SourceLimits,
    ) -> Result<Self> {
        lifecycle::open::open_records_arc_with_limit(bytes, max_output_bytes, source_limits)
    }

    /// Selects where a finished package places its sectors.
    ///
    /// The default is [`litchi_cfb::SectorLayoutPolicy::Reuse`]: the appended
    /// edit keeps every untouched stream on the sectors the opened deck
    /// already gave it and appends only the sectors the grown `PowerPoint
    /// Document` stream needs. [`litchi_cfb::SectorLayoutPolicy::Rewrite`]
    /// re-lays out the whole container.
    ///
    /// Both policies publish the same logical package.
    #[cfg(test)]
    pub(crate) const fn set_sector_layout_policy(
        &mut self,
        policy: litchi_cfb::SectorLayoutPolicy,
    ) {
        self.layout = policy;
    }

    /// Live persisted identifiers in ascending order.
    #[must_use]
    pub(crate) fn persist_ids(&self) -> Vec<u32> {
        lifecycle::persist_ids(self)
    }

    /// Returns one complete live persisted record.
    ///
    /// # Errors
    ///
    /// Returns an error if serialization fails or the underlying writer reports an error.
    pub(crate) fn persisted_record(&self, persist_id: u32) -> Result<Vec<u8>> {
        lifecycle::persisted_record(self, persist_id)
    }

    /// Stages one complete replacement record in the next incremental edit.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub(crate) fn replace_persisted_record(
        &mut self,
        persist_id: u32,
        record: Vec<u8>,
    ) -> Result<()> {
        mutation::replace_persisted_record(self, persist_id, record)
    }

    /// Whether a replacement of `persist_id` is staged.
    pub(crate) fn is_record_staged(&self, persist_id: u32) -> bool {
        self.staged_storage.contains_key(&persist_id)
    }

    /// Whether any record replacement or insertion is staged.
    pub(crate) fn has_staged_records(&self) -> bool {
        !self.staged_storage.is_empty()
    }

    /// The source bytes of the live record `persist_id`, ignoring any staged
    /// replacement.
    ///
    /// # Errors
    ///
    /// Returns an error if `persist_id` has no live source record.
    pub(crate) fn source_record(&self, persist_id: u32) -> Result<&[u8]> {
        let offset = *self
            .mappings
            .get(&persist_id)
            .ok_or_else(|| Error::Corrupted(format!("unknown persist ID {persist_id}")))?;
        let offset = usize::try_from(offset)
            .map_err(|_err| Error::Corrupted("persist offset exceeds usize".into()))?;
        rewrite::slice(&self.document, offset)
    }

    /// Drops the staged replacement of `persist_id`, so the source record
    /// stays live.
    pub(crate) fn unstage_record(&mut self, persist_id: u32) {
        self.staged_storage.remove(&persist_id);
    }

    /// Stages a complete record under a caller-allocated unused persist ID.
    ///
    /// # Errors
    ///
    /// Returns an error when the identifier is invalid or already live, or
    /// when the record is not one complete bounded PPT record.
    pub(crate) fn insert_persisted_record(
        &mut self,
        persist_id: u32,
        record: Vec<u8>,
    ) -> Result<()> {
        mutation::insert_persisted_record(self, persist_id, record)
    }

    /// Replaces the unique stream with `name`, or creates it at the package
    /// root when it is absent.
    pub(crate) fn replace_stream_by_name(&mut self, name: &str, data: Vec<u8>) -> Result<()> {
        let mut matches = self
            .streams
            .iter_mut()
            .filter(|(path, _data)| path.last().is_some_and(|component| component == name));
        if let Some((_path, stream)) = matches.next() {
            if matches.next().is_some() {
                return Err(Error::Corrupted(format!(
                    "PPT package contains multiple {name} streams"
                )));
            }
            *stream = data;
        } else {
            self.streams.push((vec![name.to_string()], data));
        }
        self.changed = true;
        Ok(())
    }

    /// Returns the typed embedded-object collection represented by this edit.
    #[must_use]
    pub(crate) fn objects(&self) -> &Collection {
        &self.collection
    }

    /// Stages a new external object and its persisted storage record.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub(crate) fn add(&mut self, object: ExternalObject, storage: Storage) -> Result<u32> {
        mutation::add(self, object, storage)
    }

    /// Stages a replacement for an existing external object's storage record.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub(crate) fn replace_storage(&mut self, persist_id: u32, storage: Storage) -> Result<()> {
        mutation::replace_storage(self, persist_id, storage)
    }

    /// Stages removal of one object reference and, when unreferenced, its
    /// persisted record.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub(crate) fn remove(&mut self, id: u32) -> Result<ExternalObject> {
        mutation::remove(self, id)
    }

    /// Stages a new object ordering for the external-object list.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub(crate) fn reorder(&mut self, ids: &[u32]) -> Result<()> {
        mutation::reorder(self, ids)
    }

    /// Whether the staged state differs from `source` and the live records:
    /// a reorder back to the source order or a storage replacement with the
    /// exact live record is no net change.
    pub(crate) fn has_net_change(&self, source: &Collection) -> Result<bool> {
        if !self.changed {
            return Ok(false);
        }
        if self.collection != *source || !self.removed_persist_ids.is_empty() {
            return Ok(true);
        }
        for (persist_id, record) in &self.staged_storage {
            let Some(offset) = self.mappings.get(persist_id) else {
                return Ok(true);
            };
            let offset = usize::try_from(*offset)
                .map_err(|_err| Error::Corrupted("persist offset exceeds usize".into()))?;
            if rewrite::slice(&self.document, offset)? != record.as_slice() {
                return Ok(true);
            }
        }
        Ok(false)
    }

    /// Emits the staged changes as one PPT incremental edit.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub(crate) fn finish(self) -> Result<Vec<u8>> {
        transaction::finish(self)
    }
}
