//! Failure-atomic semantic edits over one OLE-object snapshot.

use crate::error::Result;

use super::super::package::Editor;
use super::super::{FormControl, OleObjectRecord};
use super::{Commit, Patch, Snapshot};

/// A detached transaction over typed XLS OLE objects and form controls.
///
/// Every operation edits a private candidate package. The source transaction
/// is replaced only after the package editor has validated and published that
/// candidate, so failed edits leave all typed views unchanged.
#[derive(Clone)]
pub struct Transaction {
    source: Snapshot,
    editor: Editor,
}

impl Transaction {
    pub(super) fn new(source: Snapshot) -> Self {
        Self {
            editor: source.editor().clone(),
            source,
        }
    }

    /// Returns the immutable source snapshot used for publication checks.
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        &self.source
    }

    /// Materializes the current candidate as a validated immutable snapshot.
    pub fn snapshot(&self) -> Result<Snapshot> {
        Snapshot::open(self.editor.clone().finish()?, self.source.limits())
    }

    /// Returns the current worksheet count without materializing the CFB.
    #[must_use]
    pub fn worksheet_count(&self) -> usize {
        self.editor.worksheet_count()
    }

    /// Returns the current typed embedded-OLE records for one worksheet.
    pub fn objects(&self, worksheet: usize) -> Result<&[OleObjectRecord]> {
        self.editor.objects(worksheet)
    }

    /// Returns the current typed form-control records for one worksheet.
    pub fn form_controls(&self, worksheet: usize) -> Result<&[FormControl]> {
        self.editor.form_controls(worksheet)
    }

    /// Adds a typed form-control `Obj` record.
    ///
    /// Existing controls and their unknown subrecords remain byte-identical;
    /// the new control is appended in worksheet record order.
    pub fn add_form_control(&mut self, worksheet: usize, control: FormControl) -> Result<()> {
        self.with_candidate(|editor| editor.add_form_control(worksheet, control))
    }

    /// Adds a typed embedded-OLE `Obj` record and its inert CFB payload.
    pub fn add_object(
        &mut self,
        worksheet: usize,
        object: OleObjectRecord,
        compound_file: Vec<u8>,
    ) -> Result<()> {
        self.with_candidate(|editor| editor.add(worksheet, object, compound_file))
    }

    /// Removes one embedded-OLE `Obj` record and an unreferenced storage.
    pub fn remove_object(&mut self, worksheet: usize, object_id: u16) -> Result<OleObjectRecord> {
        let mut candidate = self.editor.clone();
        let removed = candidate.remove(worksheet, object_id)?;
        self.editor = candidate;
        Ok(removed)
    }

    /// Reorders embedded-OLE records while preserving their raw subrecords.
    pub fn reorder_objects(&mut self, worksheet: usize, ids: &[u16]) -> Result<()> {
        self.with_candidate(|editor| editor.reorder(worksheet, ids))
    }

    /// Replaces one referenced storage with a validated inert CFB payload.
    pub fn replace_storage(&mut self, storage_name: &str, compound_file: Vec<u8>) -> Result<()> {
        self.with_candidate(|editor| editor.replace_storage(storage_name, compound_file))
    }

    /// Whether the current candidate serializes differently from its source.
    pub fn is_changed(&self) -> Result<bool> {
        Ok(self.editor.clone().finish()? != self.source.bytes())
    }

    /// Discards the candidate and returns the original snapshot.
    #[must_use]
    pub fn rollback(self) -> Snapshot {
        self.source
    }

    /// Validates and publishes the candidate as a snapshot plus reversible
    /// source-checked patch.
    pub fn commit(self) -> Result<Commit> {
        let before = self.source;
        let after = self.editor.finish()?;
        let patch = Patch::new(before.bytes().to_vec(), after.clone());
        let snapshot = if patch.is_noop() {
            before
        } else {
            Snapshot::open(after, before.limits())?
        };
        Ok(Commit::new(snapshot, patch))
    }

    fn with_candidate<F>(&mut self, edit: F) -> Result<()>
    where
        F: FnOnce(&mut Editor) -> Result<()>,
    {
        let mut candidate = self.editor.clone();
        edit(&mut candidate)?;
        self.editor = candidate;
        Ok(())
    }
}
