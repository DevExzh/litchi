//! Failure-atomic semantic edits over one OLE-object snapshot.

use crate::error::Result;

use super::super::package::Editor;
use super::super::{
    EmbeddedObjectDraft, EmbeddedPayload, FormControl, ObjectMetadataChange, OleObjectRecord,
};
use super::{Commit, Diagnostics, Patch, Snapshot};

/// A detached transaction over typed XLS OLE objects and form controls.
///
/// Every operation edits a private candidate package. The source transaction
/// is replaced only after the package editor has validated and published that
/// candidate, so failed edits leave all typed views unchanged.
#[derive(Clone)]
pub struct Edit {
    base: crate::scoped::Base,
    package: Option<crate::Package>,
    source: Option<Snapshot>,
    editor: Option<Editor>,
    limits: litchi_ole_common::object::Limits,
}

impl Edit {
    #[cfg(test)]
    pub(super) fn new(base: crate::scoped::Base, source: Snapshot) -> Self {
        let limits = source.limits();
        Self {
            base,
            package: None,
            editor: Some(source.editor().clone()),
            source: Some(source),
            limits,
        }
    }

    pub(super) fn lazy(
        base: crate::scoped::Base,
        package: crate::Package,
        limits: litchi_ole_common::object::Limits,
    ) -> Self {
        Self {
            base,
            package: Some(package),
            source: None,
            editor: None,
            limits,
        }
    }

    fn ensure_open(&mut self) -> Result<()> {
        if self.source.is_some() {
            return Ok(());
        }
        let package = self.package.clone().ok_or_else(|| {
            crate::Error::InvalidData("OLE-object edit lost its source package".into())
        })?;
        let (_, bytes) = crate::scoped::capture(&package)?;
        let source = Snapshot::open(bytes, self.limits)?;
        let editor = source.editor().clone();
        self.source = Some(source);
        self.editor = Some(editor);
        Ok(())
    }

    /// Returns the immutable source snapshot used for publication checks.
    #[must_use]
    pub fn before(&mut self) -> Result<&Snapshot> {
        self.ensure_open()?;
        self.source.as_ref().ok_or_else(|| {
            crate::Error::InvalidData("OLE-object edit source was not retained".into())
        })
    }

    /// Returns the current worksheet count without materializing the CFB.
    #[must_use]
    pub fn worksheet_count(&mut self) -> Result<usize> {
        self.ensure_open()?;
        Ok(self
            .editor
            .as_ref()
            .ok_or_else(|| crate::Error::InvalidData("OLE-object editor was not retained".into()))?
            .worksheet_count())
    }

    /// Returns the current typed embedded-OLE records for one worksheet.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn objects(&mut self, worksheet: usize) -> Result<&[OleObjectRecord]> {
        self.ensure_open()?;
        self.editor
            .as_ref()
            .ok_or_else(|| crate::Error::InvalidData("OLE-object editor was not retained".into()))?
            .objects(worksheet)
    }

    /// Returns the current typed form-control records for one worksheet.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn form_controls(&mut self, worksheet: usize) -> Result<&[FormControl]> {
        self.ensure_open()?;
        self.editor
            .as_ref()
            .ok_or_else(|| crate::Error::InvalidData("OLE-object editor was not retained".into()))?
            .form_controls(worksheet)
    }

    /// Adds a typed form-control `Obj` record.
    ///
    /// Existing controls and their unknown subrecords remain byte-identical;
    /// the new control is appended in worksheet record order.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn add_form_control(&mut self, worksheet: usize, control: FormControl) -> Result<()> {
        self.with_candidate(|editor| editor.add_form_control(worksheet, control))
    }

    /// Adds a typed embedded-OLE `Obj` record and its inert CFB payload.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn add_object(
        &mut self,
        worksheet: usize,
        object: OleObjectRecord,
        compound_file: Vec<u8>,
    ) -> Result<()> {
        self.with_candidate(|editor| editor.add(worksheet, object, compound_file))
    }

    /// Adds a bounded, inert storage-backed embedded payload.
    ///
    /// The typed request authors the matching `Obj` and `MBDxxxxxxxx` storage
    /// as one identity closure.  The CFB remains opaque and no OLE server,
    /// macro, control, link, or native payload is activated.
    /// # Errors
    ///
    /// Returns an error if the request conflicts with workbook identity or
    /// the payload fails bounded CFB admission.
    pub fn add_embedded_payload(
        &mut self,
        worksheet: usize,
        draft: EmbeddedObjectDraft,
    ) -> Result<()> {
        self.with_candidate(|editor| editor.add_embedded_payload(worksheet, draft))
    }

    /// Removes one embedded-OLE `Obj` record and an unreferenced storage.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn remove_object(&mut self, worksheet: usize, object_id: u16) -> Result<OleObjectRecord> {
        self.ensure_open()?;
        let mut candidate = self
            .editor
            .as_ref()
            .ok_or_else(|| crate::Error::InvalidData("OLE-object editor was not retained".into()))?
            .clone();
        let removed = candidate.remove(worksheet, object_id)?;
        self.editor = Some(candidate);
        Ok(removed)
    }

    /// Removes a selected storage-backed embedded payload and its unreferenced
    /// MBD storage.
    ///
    /// Linked/DDE and controls-stream objects are outside this API and are
    /// rejected before the candidate is changed.
    /// # Errors
    ///
    /// Returns an error if the selected object is absent or is not a valid
    /// storage-backed embedded object.
    pub fn remove_embedded_payload(&mut self, worksheet: usize, object_id: u16) -> Result<()> {
        self.ensure_open()?;
        let mut candidate = self
            .editor
            .as_ref()
            .ok_or_else(|| crate::Error::InvalidData("OLE-object editor was not retained".into()))?
            .clone();
        candidate.remove_embedded_payload(worksheet, object_id)?;
        self.editor = Some(candidate);
        Ok(())
    }

    /// Reorders embedded-OLE records while preserving their raw subrecords.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn reorder_objects(&mut self, worksheet: usize, ids: &[u16]) -> Result<()> {
        self.with_candidate(|editor| editor.reorder(worksheet, ids))
    }

    /// Updates bounded metadata on one existing embedded OLE object.
    ///
    /// `object_id` is the source identity check for the operation. The edit
    /// only changes `FtCmo` identity/flags and `FtPioGrbit`; the `FtPictFmla`
    /// storage reference, unknown subrecords, text object, and inert payload
    /// remain untouched.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn update_object_metadata(
        &mut self,
        worksheet: usize,
        object_id: u16,
        edit: ObjectMetadataChange,
    ) -> Result<()> {
        self.with_candidate(|editor| editor.update_object_metadata(worksheet, object_id, edit))
    }

    /// Replaces one referenced storage with a validated inert CFB payload.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn replace_storage(&mut self, storage_name: &str, compound_file: Vec<u8>) -> Result<()> {
        self.with_candidate(|editor| editor.replace_storage(storage_name, compound_file))
    }

    /// Replaces one selected storage-backed embedded payload while retaining
    /// its object ID and MBD storage identity.
    ///
    /// The supplied CFB owns the selected storage's complete inert subtree;
    /// unrelated workbook records and streams remain retained.  The common
    /// editor validates the replacement under the snapshot limits and refuses
    /// protected containers before publication.
    /// # Errors
    ///
    /// Returns an error if the selected object is absent or is not an MBD
    /// embedded object, or if the payload is malformed, protected, or too
    /// large.
    pub fn replace_embedded_payload(
        &mut self,
        worksheet: usize,
        object_id: u16,
        payload: EmbeddedPayload,
    ) -> Result<()> {
        self.with_candidate(|editor| editor.replace_embedded_payload(worksheet, object_id, payload))
    }

    /// Validates and publishes the candidate as a new package plus a
    /// reversible source-checked patch.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn commit(mut self) -> Result<Commit> {
        self.ensure_open()?;
        let base = self.base.clone();
        let record = self.commit_record()?;
        let package = crate::scoped::publish(base, record.patch.before(), record.patch.after())?;
        Ok(Commit::new(package, record.patch, record.diagnostics))
    }

    /// The record-level result: the reopened candidate snapshot, its patch
    /// and diagnostics, before the package is republished.
    pub(crate) fn commit_record(mut self) -> Result<RecordCommit> {
        self.ensure_open()?;
        let before = self.source.take().ok_or_else(|| {
            crate::Error::InvalidData("OLE-object edit source was not retained".into())
        })?;
        let editor = self.editor.take().ok_or_else(|| {
            crate::Error::InvalidData("OLE-object editor was not retained".into())
        })?;
        let after = editor.finish()?;
        let patch = Patch::new(before.bytes().to_vec(), after);
        // Reopen the candidate through the object owner: the commit proves
        // the rendered artifact still validates under its bounds.
        let snapshot = if patch.is_noop() {
            before
        } else {
            Snapshot::open(patch.after().to_vec(), before.limits())?
        };
        let diagnostics = Diagnostics::new(!patch.is_noop());
        Ok(RecordCommit {
            snapshot,
            patch,
            diagnostics,
        })
    }

    fn with_candidate<F>(&mut self, edit: F) -> Result<()>
    where
        F: FnOnce(&mut Editor) -> Result<()>,
    {
        self.ensure_open()?;
        let mut candidate = self
            .editor
            .as_ref()
            .ok_or_else(|| crate::Error::InvalidData("OLE-object editor was not retained".into()))?
            .clone();
        edit(&mut candidate)?;
        self.editor = Some(candidate);
        Ok(())
    }
}

/// The record-level result of an OLE-object edit, before publication.
#[derive(Clone)]
pub(crate) struct RecordCommit {
    #[cfg_attr(not(test), allow(dead_code, reason = "the record-level tests read it"))]
    snapshot: Snapshot,
    patch: Patch,
    diagnostics: Diagnostics,
}

#[cfg(test)]
#[allow(dead_code, reason = "not every record-level test reads every part")]
impl RecordCommit {
    pub(crate) const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    pub(crate) fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }

    pub(crate) const fn patch(&self) -> &Patch {
        &self.patch
    }

    pub(crate) const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}
