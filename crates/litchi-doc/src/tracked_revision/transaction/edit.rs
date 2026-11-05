//! Clone-first semantic edits over DOC tracked-revision snapshots.

use super::super::{Revision, RevisionEditor, RevisionKind, RevisionMetadata};
use super::snapshot::Snapshot;
use super::{Commit, Error, Patch, Result};

/// A staged, failure-atomic edit over one source tracked-revision snapshot.
#[derive(Clone)]
pub(crate) struct Transaction {
    source: Snapshot,
    editor: RevisionEditor,
}

impl Transaction {
    pub(super) fn new(source: Snapshot) -> Result<Self> {
        let editor = source.editor().map_err(Error::Invalid)?;
        Ok(Self { source, editor })
    }

    /// Returns the immutable source snapshot used for stale-source checks.
    #[must_use]
    pub(crate) const fn before(&self) -> &Snapshot {
        &self.source
    }

    /// Returns the current candidate revisions in source order.
    pub(crate) fn revisions(&self) -> Result<Vec<Revision>> {
        self.editor.revisions().map_err(Error::Invalid)
    }

    /// Returns the current candidate revision-author table.
    #[must_use]
    pub(crate) fn authors(&self) -> &[String] {
        self.editor.authors()
    }

    /// Materializes the current candidate as a validated snapshot.
    ///
    /// A byte-identical candidate returns the original snapshot, preserving
    /// exact source bytes and avoiding an unnecessary CFB parse and allocation.
    pub(crate) fn snapshot(&self) -> Result<Snapshot> {
        let bytes = self.editor.clone().finish().map_err(Error::Invalid)?;
        if bytes == self.source.bytes() {
            return Ok(self.source.clone());
        }
        Snapshot::open_with_policy(
            bytes,
            self.source.limits(),
            self.source.protection_policy.clone(),
        )
        .map_err(Error::Invalid)
    }

    /// Adds a revision mark to an existing main-story range.
    pub(crate) fn add(
        &mut self,
        start_cp: u32,
        end_cp: u32,
        kind: RevisionKind,
        metadata: RevisionMetadata,
    ) -> Result<Revision> {
        self.editor
            .add(start_cp, end_cp, kind, metadata)
            .map_err(Error::Invalid)
    }

    /// Inserts inert plain text and marks it as an insertion or move target.
    ///
    /// `cp` must lie before the main story's final paragraph mark (MS-DOC
    /// 2.3.1); see [`RevisionEditor::add_text`].
    pub(crate) fn add_text(
        &mut self,
        cp: u32,
        text: &str,
        kind: RevisionKind,
        metadata: RevisionMetadata,
    ) -> Result<Revision> {
        self.editor
            .add_text(cp, text, kind, metadata)
            .map_err(Error::Invalid)
    }

    /// Replaces one revision's typed metadata without touching unrelated SPRMs.
    ///
    /// Replacing with the already-published metadata is an exact no-op and does
    /// not rebuild the FKP or append a new table block.
    pub(crate) fn replace(&mut self, index: usize, metadata: RevisionMetadata) -> Result<Revision> {
        let current = self
            .editor
            .revisions()
            .map_err(Error::Invalid)?
            .get(index)
            .cloned()
            .ok_or_else(|| invalid("revision index is out of range"))?;
        if same_metadata(&current, &metadata) {
            return Ok(current);
        }
        self.editor.update(index, metadata).map_err(Error::Invalid)
    }

    /// Removes one revision mark while retaining its text/current formatting.
    pub(crate) fn remove(&mut self, index: usize) -> Result<Revision> {
        self.editor.remove(index).map_err(Error::Invalid)
    }

    /// Accepts one revision using Word redline semantics.
    pub(crate) fn accept(&mut self, index: usize) -> Result<Revision> {
        self.editor.accept(index).map_err(Error::Invalid)
    }

    /// Rejects one revision using Word redline semantics.
    pub(crate) fn reject(&mut self, index: usize) -> Result<Revision> {
        self.editor.reject(index).map_err(Error::Invalid)
    }

    /// Commits the candidate as a validated snapshot and reversible patch.
    pub(crate) fn commit(self) -> Result<Commit> {
        let snapshot = self.snapshot()?;
        let patch = Patch::new(self.source, snapshot.clone());
        Ok(Commit { snapshot, patch })
    }
}

fn same_metadata(revision: &Revision, metadata: &RevisionMetadata) -> bool {
    revision.author == metadata.author
        && revision.timestamp == metadata.timestamp
        && revision.reason == metadata.reason
        && revision.revision_save_id == metadata.revision_save_id
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(crate::package::Error::InvalidFormat(message.into()))
}
