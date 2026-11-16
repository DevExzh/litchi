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

    /// Adds a typed form control shown at `anchor`; returns its `Obj`
    /// identifier.
    ///
    /// The control is written as Excel writes one: a host-control shape at
    /// the anchor and the control's `Obj` right after it. The identifier is
    /// allocated unique within the worksheet (the control's own `FtCmo.id`
    /// is replaced). Existing controls and their unknown
    /// subrecords remain byte-identical. A label (`TxO`) is not written: a
    /// control with a text object is refused.
    /// # Errors
    ///
    /// Returns an error if the control is invalid or carries a text object,
    /// the workbook's object identities are not proven, or the worksheet
    /// drawing cannot take another shape.
    pub fn add_form_control(
        &mut self,
        worksheet: usize,
        control: FormControl,
        anchor: crate::sheet_features::edit::Anchor,
    ) -> Result<u16> {
        self.ensure_open()?;
        self.editor()?.check_add_form_control(worksheet, &control)?;
        self.with_drawing(worksheet, |features, sheet| {
            features.add_form_control_object(sheet, anchor, |object_id| {
                control.with_object_id(object_id).to_record_bytes()
            })
        })
    }

    /// Adds a bounded, inert storage-backed embedded payload; returns the
    /// new object's `Obj` identifier.
    ///
    /// The object is written as Excel writes one: a picture-frame shape at
    /// the draft's anchor showing its preview, the matching `Obj` right
    /// after it, and a new `MBDxxxxxxxx` storage holding the payload. The
    /// CFB remains opaque and no OLE server, macro, control, link, or
    /// native payload is activated.
    /// # Errors
    ///
    /// Returns an error if the workbook's object identities are not proven,
    /// the payload fails bounded CFB admission, the preview is not PNG or
    /// JPEG, or the worksheet drawing cannot take another shape.
    pub fn add_embedded_payload(
        &mut self,
        worksheet: usize,
        draft: EmbeddedObjectDraft,
    ) -> Result<u16> {
        self.ensure_open()?;
        self.editor()?
            .check_add_embedded_payload(worksheet, &draft)?;
        let spec = draft.spec();
        self.with_drawing(worksheet, |features, sheet| {
            features.add_embedded_object(sheet, &spec)
        })
    }

    /// Removes one storage-backed embedded object: its drawing shape, its
    /// `Obj`, its preview picture reference and its storage when no other
    /// `Obj` names it. Returns the removed record.
    /// # Errors
    ///
    /// Returns an error if the object is absent, is a linked or control
    /// object, or the removal would orphan an opaque storage reference.
    pub fn remove_object(&mut self, worksheet: usize, object_id: u16) -> Result<OleObjectRecord> {
        self.ensure_open()?;
        let removed = self.editor()?.check_remove(worksheet, object_id)?;
        self.with_drawing(worksheet, |features, sheet| {
            features.remove_embedded_object(sheet, object_id)
        })?;
        Ok(removed)
    }

    /// Removes a selected storage-backed embedded payload with its drawing
    /// shape, preview picture reference and unreferenced MBD storage.
    ///
    /// Linked/DDE and controls-stream objects are outside this API and are
    /// rejected before the candidate is changed.
    /// # Errors
    ///
    /// Returns an error if the selected object is absent or is not a valid
    /// storage-backed embedded object.
    pub fn remove_embedded_payload(&mut self, worksheet: usize, object_id: u16) -> Result<()> {
        self.ensure_open()?;
        self.editor()?.check_embedded(worksheet, object_id)?;
        self.remove_object(worksheet, object_id).map(|_| ())
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

    fn editor(&self) -> Result<&Editor> {
        self.editor
            .as_ref()
            .ok_or_else(|| crate::Error::InvalidData("OLE-object editor was not retained".into()))
    }

    /// Runs a drawing-aware worksheet-feature edit over the current
    /// candidate (with the worksheet's catalog position) and adopts its
    /// result. The worksheet-feature edit owns the drawing layer, so a
    /// shape and its `Obj` are always added and removed together.
    fn with_drawing<T>(
        &mut self,
        worksheet: usize,
        edit: impl FnOnce(
            &mut crate::sheet_features::edit::Edit,
            usize,
        ) -> crate::sheet_features::edit::Result<T>,
    ) -> Result<T> {
        let editor = self.editor()?;
        let sheet = editor.sheet_position(worksheet)?;
        let bytes = editor.clone().finish()?;
        let package = crate::Package::from_bytes(&bytes).map_err(crate::scoped::package_error)?;
        let mut features = package.edit_sheet_features().map_err(feature_error)?;
        let value = edit(&mut features, sheet).map_err(feature_error)?;
        let commit = features.commit().map_err(feature_error)?;
        let candidate = Editor::new(commit.patch().after().to_vec(), self.limits)?;
        self.editor = Some(candidate);
        Ok(value)
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

/// The workbook error of a worksheet-feature edit run for an OLE-object
/// edit.
fn feature_error(error: crate::sheet_features::edit::Error) -> crate::Error {
    use crate::sheet_features::edit::Error as Feature;
    match error {
        Feature::Package(error) => crate::scoped::package_error(error),
        Feature::Records(error) => error,
        Feature::Refused(refusal) => crate::Error::Refused(refusal),
        Feature::PatchConflict => crate::Error::PatchConflict {
            operation: "edit OLE objects",
            reason: "the workbook changed during the edit",
        },
        Feature::Signed => crate::Error::Signed,
        other => crate::Error::InvalidData(other.to_string()),
    }
}
