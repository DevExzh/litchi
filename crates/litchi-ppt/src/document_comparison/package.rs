//! OLE package discovery and atomic publication for review metadata.
//!
//! The semantic owner edits only the live `DocumentContainer` persisted
//! record. Publication goes through the existing incremental PPT record
//! editor, so old stream bytes, opaque records, and their record framing stay
//! intact. No comparison, merge, accept, or reject operation is performed.

use super::model::Limits;
use super::transaction::{
    Change, Editor as DocumentEditor, Patch as DocumentPatch, Revision,
    Snapshot as DocumentSnapshot,
};
use crate::package::{Error, Package as PresentationPackage, Result};
use crate::presentation::Presentation;
use crate::records::Record;
use std::io::Cursor;
use std::sync::Arc;

/// Maximum complete OLE artifact accepted by the bounded package facade.
pub const MAX_PACKAGE_BYTES: usize = 256 * 1024 * 1024;

/// An immutable PPT package snapshot with its live document-comparison owner.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Snapshot {
    bytes: Arc<[u8]>,
    document: DocumentSnapshot,
    document_persist_id: u32,
    limits: Limits,
}

impl Snapshot {
    /// Parse a complete OLE2 PowerPoint package using default limits.
    pub fn parse(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes_with_limits(bytes.to_vec(), Limits::default())
    }

    /// Capture a complete OLE2 PowerPoint package without another caller copy.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        Self::from_bytes_with_limits(bytes, Limits::default())
    }

    /// Parse a package with the document-comparison resource limits supplied
    /// by the caller.
    pub fn from_bytes_with_limits(bytes: Vec<u8>, limits: Limits) -> Result<Self> {
        if bytes.len() > MAX_PACKAGE_BYTES {
            return Err(Error::InvalidFormat(
                "PowerPoint package exceeds the bounded package byte limit".into(),
            ));
        }
        let mut package = PresentationPackage::from_reader(Cursor::new(bytes.clone()))?;
        let presentation = package.presentation()?;
        let document_persist_id = presentation.slide_directory().document_persist_id();
        let document = snapshot_from_presentation(&presentation, limits)?;
        Ok(Self {
            bytes: Arc::from(bytes.into_boxed_slice()),
            document,
            document_persist_id,
            limits,
        })
    }

    /// Exact source bytes of the complete OLE2 artifact.
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Immutable live `DocumentContainer` snapshot.
    pub const fn document(&self) -> &DocumentSnapshot {
        &self.document
    }

    /// Native persist identifier used by the current `UserEditAtom`.
    pub const fn document_persist_id(&self) -> u32 {
        self.document_persist_id
    }

    /// Limits captured by this package snapshot.
    pub const fn limits(&self) -> Limits {
        self.limits
    }

    /// Content revision of the complete package artifact.
    #[must_use]
    pub fn revision(&self) -> Revision {
        Revision::from_bytes(&self.bytes)
    }

    /// Begin an isolated package edit over the review-owned metadata.
    pub fn edit(&self) -> Editor {
        Editor {
            source: self.clone(),
            document: self.document.edit(),
        }
    }
}

/// An isolated package transaction over inert document-comparison metadata.
#[derive(Debug, Clone)]
pub struct Editor {
    source: Snapshot,
    document: DocumentEditor,
}

impl Editor {
    /// Read the transaction-local review records.
    pub fn review(&self) -> Result<super::model::Review> {
        self.document.review()
    }

    /// Read the transaction-local toolbar state.
    pub fn toolbar(&self) -> Result<Option<super::model::ReviewingToolbarStates>> {
        self.document.toolbar()
    }

    /// Read one transaction-local reviewer tree.
    pub fn diff_tree(&self, index: usize) -> Result<Option<super::model::DiffTree10>> {
        self.document.diff_tree(index)
    }

    /// Return the current transaction-local document snapshot.
    pub fn document_snapshot(&self) -> Result<DocumentSnapshot> {
        self.document.snapshot()
    }

    /// Return semantic changes staged in this package transaction.
    pub fn changes(&self) -> &[Change] {
        self.document.changes()
    }

    /// Whether any review metadata has been staged.
    pub const fn is_changed(&self) -> bool {
        self.document.is_changed()
    }

    /// Set or replace the reviewing toolbar state.
    pub fn set_toolbar(&mut self, value: super::model::ReviewingToolbarStates) -> Result<()> {
        self.document.set_toolbar(value)
    }

    /// Remove the reviewing toolbar state, preserving all other PP10 records.
    pub fn clear_toolbar(&mut self) -> Result<bool> {
        self.document.clear_toolbar()
    }

    /// Set one reviewer name in source order.
    pub fn set_reviewer_name(&mut self, tree_index: usize, value: impl Into<String>) -> Result<()> {
        self.document.set_reviewer_name(tree_index, value)
    }

    /// Set one reviewer tree's document-level display flags.
    pub fn set_document_flags(
        &mut self,
        tree_index: usize,
        value: super::model::DiffFlags,
    ) -> Result<()> {
        self.document.set_document_flags(tree_index, value)
    }

    /// Publish the package edit as a fresh OLE artifact and source-checked
    /// reversible patch.
    pub fn commit(self) -> Result<Commit> {
        let document_commit = self.document.commit()?;
        let source = self.source;
        let target_document = document_commit.snapshot().clone();

        if target_document.bytes() == source.document.bytes() {
            let patch = Patch {
                base: source.revision(),
                target: source.revision(),
                before: source.bytes.clone(),
                after: source.bytes.clone(),
                document: document_commit.patch().clone(),
                limits: source.limits,
            };
            return Ok(Commit {
                snapshot: source,
                patch,
            });
        }

        let mut editor = crate::embedded::object::Editor::open_records(source.bytes.to_vec())?;
        let current = editor.persisted_record(source.document_persist_id)?;
        if current.as_slice() != source.document.bytes() {
            return Err(Error::InvalidFormat(
                "live PowerPoint document changed during review staging".into(),
            ));
        }
        editor.replace_persisted_record(
            source.document_persist_id,
            target_document.bytes().to_vec(),
        )?;
        let bytes = editor.finish()?;
        let snapshot = Snapshot::from_bytes_with_limits(bytes, source.limits)?;
        if snapshot.document_persist_id != source.document_persist_id
            || snapshot.document.bytes() != target_document.bytes()
        {
            return Err(Error::InvalidFormat(
                "published review metadata did not round-trip through the live document".into(),
            ));
        }
        let patch = Patch {
            base: source.revision(),
            target: snapshot.revision(),
            before: source.bytes,
            after: snapshot.bytes.clone(),
            document: document_commit.patch().clone(),
            limits: snapshot.limits,
        };
        Ok(Commit { snapshot, patch })
    }

    /// Discard the package candidate and recover the source snapshot.
    pub fn rollback(self) -> Snapshot {
        self.source
    }
}

/// A committed package snapshot and reversible source-checked patch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    pub fn document(&self) -> &DocumentSnapshot {
        self.snapshot.document()
    }

    pub fn undo(&self, current: &Snapshot) -> Result<Snapshot> {
        self.patch.undo(current)
    }

    pub fn redo(&self, current: &Snapshot) -> Result<Snapshot> {
        self.patch.redo(current)
    }

    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

/// A reversible whole-artifact package patch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Patch {
    base: Revision,
    target: Revision,
    before: Arc<[u8]>,
    after: Arc<[u8]>,
    document: DocumentPatch,
    limits: Limits,
}

impl Patch {
    pub const fn base(&self) -> Revision {
        self.base
    }

    pub const fn target(&self) -> Revision {
        self.target
    }

    pub fn before_bytes(&self) -> &[u8] {
        &self.before
    }

    pub fn after_bytes(&self) -> &[u8] {
        &self.after
    }

    pub const fn document(&self) -> &DocumentPatch {
        &self.document
    }

    pub fn undo(&self, current: &Snapshot) -> Result<Snapshot> {
        if current.revision() != self.target || current.bytes() != self.after.as_ref() {
            return Err(Error::InvalidFormat(
                "cannot undo review package edits against a different source".into(),
            ));
        }
        Snapshot::from_bytes_with_limits(self.before.to_vec(), self.limits)
    }

    pub fn redo(&self, current: &Snapshot) -> Result<Snapshot> {
        if current.revision() != self.base || current.bytes() != self.before.as_ref() {
            return Err(Error::InvalidFormat(
                "cannot redo review package edits against a different source".into(),
            ));
        }
        Snapshot::from_bytes_with_limits(self.after.to_vec(), self.limits)
    }
}

/// Load the live document-comparison snapshot from an already opened PPT.
pub(crate) fn from_presentation(presentation: &Presentation) -> Result<DocumentSnapshot> {
    snapshot_from_presentation(presentation, Limits::default())
}

fn snapshot_from_presentation(
    presentation: &Presentation,
    limits: Limits,
) -> Result<DocumentSnapshot> {
    let offset = presentation.slide_directory().document_offset();
    let stream = presentation.document_stream();
    let (record, consumed) = Record::parse_strict(stream, offset)?;
    let end = offset
        .checked_add(consumed)
        .ok_or_else(|| Error::Corrupted("live document record offset overflow".into()))?;
    let source = stream
        .get(offset..end)
        .ok_or_else(|| Error::Corrupted("live document record exceeds its stream".into()))?;
    if record.record_type != crate::consts::RecordType::Document {
        return Err(Error::Corrupted(
            "live persist record is not a DocumentContainer".into(),
        ));
    }
    DocumentSnapshot::parse_with_limits(source, limits)
}
