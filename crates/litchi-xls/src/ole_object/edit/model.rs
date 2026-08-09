//! Immutable OLE-object snapshots and reversible artifact patches.

use std::fmt;
use std::sync::Arc;

use crate::error::{Error, Result};

use super::super::package::Editor;
use super::super::{FormControl, Limits, OleObjectRecord};

/// An immutable typed view of one exact legacy XLS CFB artifact.
///
/// The snapshot retains the source bytes for exact no-op publication and
/// delegates semantic BIFF views to the existing package editor. Embedded
/// storages, controls, unknown `Obj` subrecords, and `TxO` records remain
/// inert data; opening a snapshot never activates them.
#[derive(Clone)]
pub struct Snapshot {
    source: Arc<[u8]>,
    limits: Limits,
    editor: Editor,
}

impl Snapshot {
    /// Opens a bounded XLS CFB artifact and captures its typed OLE-object and
    /// form-control views.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn open(bytes: impl Into<Vec<u8>>, limits: Limits) -> Result<Self> {
        let bytes = bytes.into();
        let editor = Editor::new(bytes.clone(), limits)?;
        Ok(Self {
            source: Arc::from(bytes),
            limits,
            editor,
        })
    }

    /// Returns the exact artifact used to create this snapshot.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.source
    }

    /// Returns the number of worksheet BIFF substreams represented by this
    /// snapshot.
    #[must_use]
    pub fn worksheet_count(&self) -> usize {
        self.editor.worksheet_count()
    }

    /// Returns typed embedded-OLE `Obj` records for one worksheet.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn objects(&self, worksheet: usize) -> Result<&[OleObjectRecord]> {
        self.editor.objects(worksheet)
    }

    /// Returns typed form-control `Obj` records for one worksheet.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn form_controls(&self, worksheet: usize) -> Result<&[FormControl]> {
        self.editor.form_controls(worksheet)
    }

    /// Starts an independent failure-atomic edit.
    #[must_use]
    pub fn edit(&self) -> super::Transaction {
        super::Transaction::new(self.clone())
    }

    /// Publishes this immutable snapshot without rebuilding the CFB.
    ///
    /// Because a snapshot always owns an already-validated artifact, this
    /// returns the exact source bytes, including producer-specific padding
    /// and unknown records.
    #[must_use]
    pub fn finish(&self) -> Vec<u8> {
        self.source.as_ref().to_vec()
    }

    pub(super) fn editor(&self) -> &Editor {
        &self.editor
    }

    pub(super) fn limits(&self) -> Limits {
        self.limits
    }
}

impl fmt::Debug for Snapshot {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Snapshot")
            .field("bytes", &self.source.len())
            .field("worksheets", &self.worksheet_count())
            .finish()
    }
}

impl PartialEq for Snapshot {
    fn eq(&self, other: &Self) -> bool {
        self.source == other.source
    }
}

impl Eq for Snapshot {}

/// A source-checked replacement of one complete XLS CFB artifact.
///
/// The patch boundary is deliberately the complete artifact: BIFF offsets,
/// CFB allocation, inert payloads, and unknown subrecords are all validated by
/// the package editor before a changed snapshot is published.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Patch {
    before: Arc<[u8]>,
    after: Arc<[u8]>,
}

impl Patch {
    pub(super) fn new(before: Vec<u8>, after: Vec<u8>) -> Self {
        Self {
            before: Arc::from(before),
            after: Arc::from(after),
        }
    }

    /// Returns the exact artifact required as the patch source.
    #[must_use]
    pub fn before(&self) -> &[u8] {
        &self.before
    }

    /// Returns the exact artifact produced by the patch.
    #[must_use]
    pub fn after(&self) -> &[u8] {
        &self.after
    }

    /// Whether the edit is a byte-for-byte no-op.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after
    }

    /// Applies this patch only to a snapshot with the exact expected source.
    ///
    /// The returned snapshot is built before it is published to the caller,
    /// so source mismatches and malformed patch output cannot partially apply.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        if source.source.as_ref() != self.before.as_ref() {
            return Err(Error::UnsafeEdit(
                "OLE-object patch source does not match its base snapshot".into(),
            ));
        }
        if self.is_noop() {
            return Ok(source.clone());
        }
        Snapshot::open(self.after.as_ref().to_vec(), source.limits)
    }

    /// Returns the exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: Arc::clone(&self.after),
            after: Arc::clone(&self.before),
        }
    }
}

/// A validated immutable publication containing the resulting snapshot and
/// its reversible patch.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    pub(super) fn new(snapshot: Snapshot, patch: Patch) -> Self {
        Self { snapshot, patch }
    }

    /// Returns the post-edit typed snapshot.
    #[must_use]
    pub fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Returns the source-checked reversible patch.
    #[must_use]
    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Whether publication changed any artifact byte.
    #[must_use]
    pub fn changed(&self) -> bool {
        !self.patch.is_noop()
    }

    /// Consumes the publication into its post-edit snapshot.
    #[must_use]
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }

    /// Consumes the publication into its reversible patch.
    #[must_use]
    pub fn into_patch(self) -> Patch {
        self.patch
    }
}
