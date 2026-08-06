//! Failure-atomic edits for one inert PowerPoint storage payload.

use super::model::{Kind, Storage};
use super::snapshot::Snapshot;
use crate::package::Result;

/// A transaction-local editor for one persisted `ExOleObjStg` payload.
///
/// Edits operate on a private candidate. Every replacement validates the
/// complete bounded candidate before publishing it into the editor, and
/// [`Self::commit`] revalidates before returning the new immutable storage.
/// The source passed to [`Storage::edit`] is never mutated; the editor never
/// opens, interprets, or executes the embedded OLE2 content.
#[derive(Clone, Debug)]
pub struct Editor {
    candidate: Storage,
}

impl Editor {
    pub(super) const fn from_storage(candidate: Storage) -> Self {
        Self { candidate }
    }

    /// Borrow the editor's current candidate without copying its payload.
    #[must_use]
    pub fn snapshot(&self) -> Snapshot<'_> {
        self.candidate.snapshot()
    }

    /// Return the candidate's payload-free metadata.
    #[must_use]
    pub fn metadata(&self) -> super::snapshot::Metadata {
        self.candidate.metadata()
    }

    /// Change the PowerPoint context associated with the payload.
    ///
    /// This does not alter the embedded bytes or activate the storage. The
    /// context is written by the host when the candidate is serialized.
    pub fn set_kind(&mut self, kind: Kind) {
        self.candidate.kind = kind;
    }

    /// Replace the candidate with an uncompressed payload.
    ///
    /// If validation fails, the editor retains its previous candidate.
    pub fn replace_uncompressed(&mut self, data: Vec<u8>) -> Result<()> {
        let candidate = Storage::uncompressed(self.candidate.kind, data)?;
        self.candidate = candidate;
        Ok(())
    }

    /// Replace the candidate with a zlib payload and its declared size.
    ///
    /// `data` excludes PowerPoint's four-byte decompressed-size prefix. If
    /// validation fails, the editor retains its previous candidate.
    pub fn replace_compressed(&mut self, uncompressed_len: u32, data: Vec<u8>) -> Result<()> {
        let candidate = Storage::compressed(self.candidate.kind, uncompressed_len, data)?;
        self.candidate = candidate;
        Ok(())
    }

    /// Publish the fully validated candidate as a new immutable storage.
    pub fn commit(self) -> Result<Storage> {
        Storage::from_parts(
            self.candidate.kind,
            self.candidate.compression,
            self.candidate.declared_uncompressed_len,
            self.candidate.data,
        )
    }

    /// Return the current outer encoding without copying the payload.
    #[must_use]
    pub const fn compression(&self) -> super::model::Compression {
        self.candidate.compression
    }

    /// Return the current PowerPoint storage context.
    #[must_use]
    pub const fn kind(&self) -> Kind {
        self.candidate.kind
    }
}

impl Storage {
    /// Start an editor with an owned candidate cloned from this snapshot.
    #[must_use]
    pub fn edit(&self) -> Editor {
        Editor::from_storage(self.clone())
    }

    /// Start an editor by moving this snapshot without copying its payload.
    #[must_use]
    pub fn into_edit(self) -> Editor {
        Editor::from_storage(self)
    }
}
