//! A failure-atomic draft of one inert `PowerPoint` storage value.

use super::model::{Kind, Storage};
use super::snapshot::Snapshot;
use crate::package::Result;

/// A draft of one `ExOleObjStg` storage value.
///
/// The draft edits a value, not a presentation: every replacement validates
/// the complete bounded candidate before adopting it, and
/// [`Self::into_storage`] revalidates before returning the new immutable
/// storage, which [`crate::embedded::object::Edit::add`] and
/// [`crate::embedded::object::Edit::replace_storage`] stage into a package.
/// The source passed to [`Storage::draft`] is never mutated; the draft never
/// opens, interprets, or executes the embedded OLE2 content.
#[derive(Clone, Debug)]
pub struct Draft {
    candidate: Storage,
}

impl Draft {
    pub(super) const fn from_storage(candidate: Storage) -> Self {
        Self { candidate }
    }

    /// Borrow the draft's current candidate without copying its payload.
    #[must_use]
    pub fn view(&self) -> Snapshot<'_> {
        self.candidate.snapshot()
    }

    /// Return the candidate's payload-free metadata.
    #[must_use]
    pub fn metadata(&self) -> super::snapshot::Metadata {
        self.candidate.metadata()
    }

    /// Change the `PowerPoint` context associated with the payload.
    ///
    /// This does not alter the embedded bytes or activate the storage. The
    /// context is written by the host when the candidate is serialized.
    pub fn set_kind(&mut self, kind: Kind) {
        self.candidate.kind = kind;
    }

    /// Replace the candidate with an uncompressed payload.
    ///
    /// If validation fails, the draft retains its previous candidate.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn replace_uncompressed(&mut self, data: Vec<u8>) -> Result<()> {
        let candidate = Storage::uncompressed(self.candidate.kind, data)?;
        self.candidate = candidate;
        Ok(())
    }

    /// Replace the candidate with a zlib payload and its declared size.
    ///
    /// `data` excludes `PowerPoint`'s four-byte decompressed-size prefix. If
    /// validation fails, the draft retains its previous candidate.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn replace_compressed(&mut self, uncompressed_len: u32, data: Vec<u8>) -> Result<()> {
        let candidate = Storage::compressed(self.candidate.kind, uncompressed_len, data)?;
        self.candidate = candidate;
        Ok(())
    }

    /// The fully validated candidate as a new immutable storage.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn into_storage(self) -> Result<Storage> {
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

    /// Return the current `PowerPoint` storage context.
    #[must_use]
    pub const fn kind(&self) -> Kind {
        self.candidate.kind
    }
}

impl Storage {
    /// Start a draft with an owned candidate cloned from this storage.
    #[must_use]
    pub fn draft(&self) -> Draft {
        Draft::from_storage(self.clone())
    }

    /// Start a draft by moving this storage without copying its payload.
    #[must_use]
    pub fn into_draft(self) -> Draft {
        Draft::from_storage(self)
    }
}
