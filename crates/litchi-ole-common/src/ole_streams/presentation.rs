//! The presentation-stream snapshot and its edit family: [`Snapshot`] is the
//! immutable parsed stream, [`Edit`] stages a change, [`Commit`] publishes
//! the new snapshot with its reversible [`Patch`] and [`Diagnostics`].

use super::*;

pub use super::{
    ClipboardFormat, Limits, OlePresentationStream, StreamRevision, TargetDevice, TocEntry,
};

/// A cheap immutable snapshot of an OlePresentationStream.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Snapshot {
    stream: OlePresentationStream,
    limits: Limits,
    revision: StreamRevision,
}

impl Snapshot {
    /// Parses a source presentation stream under explicit limits.
    pub fn parse(bytes: &[u8], limits: Limits) -> Result<Self, Error> {
        let limits = limits.validate()?;
        if bytes.len() > limits.max_bytes {
            return Err(limit(
                crate::LimitKind::StreamBytes,
                bytes.len(),
                limits.max_bytes,
            ));
        }
        Self::parse_shared(Arc::<[u8]>::from(bytes), limits)
    }

    /// Parses a source presentation stream without copying its allocation.
    pub fn parse_shared(bytes: Arc<[u8]>, limits: Limits) -> Result<Self, Error> {
        let limits = limits.validate()?;
        let stream = OlePresentationStream::parse_shared(bytes, limits)?;
        let revision = StreamRevision::from(Revision::of(stream.bytes()));
        Ok(Self {
            stream,
            limits,
            revision,
        })
    }

    /// Captures an already parsed stream as a source snapshot.
    pub fn from_stream(stream: OlePresentationStream, limits: Limits) -> Result<Self, Error> {
        let limits = limits.validate()?;
        stream.validate(limits)?;
        if stream.bytes().len() > limits.max_bytes {
            return Err(limit(
                crate::LimitKind::StreamBytes,
                stream.bytes().len(),
                limits.max_bytes,
            ));
        }
        if stream.dirty {
            let bytes = stream.to_bytes(limits)?;
            return Self::parse_shared(bytes.into(), limits);
        }
        let revision = StreamRevision::from(Revision::of(stream.bytes()));
        Ok(Self {
            stream,
            limits,
            revision,
        })
    }

    /// Borrows the typed presentation projection.
    #[must_use]
    pub const fn stream(&self) -> &OlePresentationStream {
        &self.stream
    }

    /// Exact source bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        self.stream.bytes()
    }

    /// Shared ownership of the exact source allocation.
    #[must_use]
    pub fn bytes_shared(&self) -> Arc<[u8]> {
        Arc::clone(&self.stream.wire)
    }

    /// Source identity used by reversible patches.
    #[must_use]
    pub const fn revision(&self) -> StreamRevision {
        self.revision
    }

    /// Source fingerprint used by reversible patches.
    #[must_use]
    pub const fn fingerprint(&self) -> u64 {
        self.revision.value()
    }

    /// Limits retained for subsequent edits.
    #[must_use]
    pub const fn limits(&self) -> Limits {
        self.limits
    }

    /// Starts an isolated transaction.
    #[must_use]
    pub fn edit(&self) -> Edit {
        Edit {
            source: self.clone(),
            candidate: self.stream.clone(),
        }
    }

    fn patch_to(&self, after: &Self) -> Patch {
        Patch::new(self.clone(), after.clone())
    }
}

impl std::ops::Deref for Snapshot {
    type Target = OlePresentationStream;

    fn deref(&self) -> &Self::Target {
        self.stream()
    }
}

/// A failure-atomic edit over one presentation snapshot.
#[derive(Debug, Clone)]
pub struct Edit {
    source: Snapshot,
    candidate: OlePresentationStream,
}

impl Edit {
    /// Borrows the immutable source snapshot.
    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        &self.source
    }

    /// Borrows the current typed candidate.
    #[must_use]
    pub const fn stream(&self) -> &OlePresentationStream {
        &self.candidate
    }

    /// Whether the candidate differs from its source projection.
    #[must_use]
    fn has_changes(&self) -> bool {
        !self.candidate.content_eq(&self.source.stream)
    }

    /// Applies one callback to a cloned candidate and validates it before
    /// publication.
    pub fn update<F>(&mut self, edit: F) -> Result<&mut Self, Error>
    where
        F: FnOnce(&mut OlePresentationStream) -> Result<(), Error>,
    {
        let mut candidate = self.candidate.clone();
        edit(&mut candidate)?;
        candidate.validate(self.source.limits)?;
        self.candidate = candidate;
        Ok(self)
    }

    /// Stages an opaque data replacement.
    pub fn set_data(&mut self, value: impl Into<Vec<u8>>) -> Result<&mut Self, Error> {
        let limits = self.source.limits;
        self.update(|stream| stream.set_data(value, limits))
    }

    /// Stages a primary clipboard-format replacement.
    pub fn set_clipboard_format(&mut self, value: ClipboardFormat) -> Result<&mut Self, Error> {
        let limits = self.source.limits;
        self.update(|stream| stream.set_clipboard_format(value, limits))
    }

    /// Stages a target-device replacement.
    pub fn set_target_device(&mut self, value: impl Into<Vec<u8>>) -> Result<&mut Self, Error> {
        let limits = self.source.limits;
        self.update(|stream| stream.set_target_device(value, limits))
    }

    /// Stages an aspect-hint replacement.
    pub fn set_aspect(&mut self, value: u32) -> Result<&mut Self, Error> {
        self.update(|stream| {
            stream.set_aspect(value);
            Ok(())
        })
    }

    /// Stages an Lindex-hint replacement.
    pub fn set_lindex(&mut self, value: u32) -> Result<&mut Self, Error> {
        self.update(|stream| {
            stream.set_lindex(value);
            Ok(())
        })
    }

    /// Stages an Advf-hint replacement.
    pub fn set_advf(&mut self, value: u32) -> Result<&mut Self, Error> {
        self.update(|stream| {
            stream.set_advf(value);
            Ok(())
        })
    }

    /// Stages a width replacement.
    pub fn set_width(&mut self, value: i32) -> Result<&mut Self, Error> {
        self.update(|stream| {
            stream.set_width(value);
            Ok(())
        })
    }

    /// Stages a height replacement.
    pub fn set_height(&mut self, value: i32) -> Result<&mut Self, Error> {
        self.update(|stream| {
            stream.set_height(value);
            Ok(())
        })
    }

    /// Stages a TOC entry edit.
    pub fn update_toc_entry<F>(&mut self, index: usize, edit: F) -> Result<&mut Self, Error>
    where
        F: FnOnce(&mut TocEntry) -> Result<(), Error>,
    {
        let limits = self.source.limits;
        self.update(|stream| stream.update_toc_entry(index, limits, edit))
    }

    /// Stages insertion of a TOC entry.
    pub fn insert_toc_entry(&mut self, index: usize, entry: TocEntry) -> Result<&mut Self, Error> {
        let limits = self.source.limits;
        self.update(|stream| stream.insert_toc_entry(index, entry, limits))
    }

    /// Stages removal of a TOC entry.
    pub fn remove_toc_entry(&mut self, index: usize) -> Result<TocEntry, Error> {
        let limits = self.source.limits;
        let mut removed = None;
        self.update(|stream| {
            removed = Some(stream.remove_toc_entry(index, limits)?);
            Ok(())
        })?;
        removed.ok_or_else(|| invalid("OlePresentationStream TOC removal produced no entry"))
    }

    /// Validates and publishes this transaction.
    pub fn commit(self) -> Result<Commit, Error> {
        let snapshot = self.materialize()?;
        let patch = self.source.patch_to(&snapshot);
        let diagnostics = Diagnostics::new(!patch.is_noop());
        Ok(Commit::new(snapshot, patch, diagnostics))
    }

    fn materialize(&self) -> Result<Snapshot, Error> {
        self.candidate.validate(self.source.limits)?;
        if !self.has_changes() {
            return Ok(self.source.clone());
        }
        let bytes = self.candidate.to_bytes(self.source.limits)?;
        Snapshot::parse_shared(bytes.into(), self.source.limits)
    }
}

/// A published edit: the new snapshot, its reversible patch and
/// diagnostics.
pub type Commit = litchi_core::edit::Commit<Snapshot, Patch>;

/// Content-free diagnostics of one commit.
pub use litchi_core::edit::Diagnostics;

/// A reversible, exact-source-checked presentation replacement.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Patch {
    base: StreamRevision,
    target: StreamRevision,
    before: Snapshot,
    after: Snapshot,
    changed: bool,
}

impl Patch {
    fn new(before: Snapshot, after: Snapshot) -> Self {
        let changed = before.bytes() != after.bytes();
        Self {
            base: before.revision,
            target: after.revision,
            before,
            after,
            changed,
        }
    }

    /// Expected source revision.
    #[must_use]
    pub const fn base(&self) -> StreamRevision {
        self.base
    }

    /// Resulting revision.
    #[must_use]
    pub const fn target(&self) -> StreamRevision {
        self.target
    }

    /// Exact source bytes required by this patch.
    #[must_use]
    pub fn before_bytes(&self) -> &[u8] {
        self.before.bytes()
    }

    /// Exact target bytes produced by this patch.
    #[must_use]
    pub fn after_bytes(&self) -> &[u8] {
        self.after.bytes()
    }

    /// Borrows the exact source snapshot retained by this patch.
    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        &self.before
    }

    /// Borrows the exact replacement snapshot retained by this patch.
    #[must_use]
    pub const fn replacement(&self) -> &Snapshot {
        &self.after
    }

    /// Whether the replacement is byte-for-byte empty.
    #[must_use]
    pub const fn is_noop(&self) -> bool {
        !self.changed
    }

    /// Applies this patch only to its exact source snapshot.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot, Error> {
        if source.revision != self.base || source.bytes() != self.before.bytes() {
            return Err(Error::PatchConflict);
        }
        Ok(self.after.clone())
    }

    /// Returns the exact inverse replacement.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            base: self.target,
            target: self.base,
            before: self.after.clone(),
            after: self.before.clone(),
            changed: self.changed,
        }
    }
}
