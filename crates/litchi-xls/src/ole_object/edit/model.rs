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
    /// A codec-level edit of this bare snapshot (tests only); package
    /// callers use [`crate::Package::edit_ole_objects`].
    #[cfg(test)]
    #[must_use]
    pub(crate) fn edit(&self) -> super::Edit {
        super::Edit::new(None, self.clone())
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
            return Err(Error::PatchConflict {
                operation: "edit OLE objects",
                reason: "OLE-object patch source does not match its base snapshot",
            });
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

/// A published OLE-object edit: the new package, its exact artifact patch
/// and diagnostics.
pub type Commit = litchi_core::edit::Commit<crate::Package, Patch>;

/// Content-free diagnostics of one commit.
pub use litchi_core::edit::Diagnostics;

/// A configured OLE-object capture or publication resource dimension.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Selected embedded object count.
    Objects,
    /// Bytes retained for one standalone embedded compound file.
    ObjectBytes,
    /// Directory entries visited during bounded capture.
    DirectoryEntries,
    /// Captured descendant storage count.
    Storages,
    /// Nested storage depth.
    StorageDepth,
    /// Captured stream count, bounded by both object and package policies.
    Streams,
    /// Bytes in one stream, before reading its payload.
    StreamBytes,
    /// Aggregate captured stream bytes.
    AggregateBytes,
}

impl LimitKind {
    pub(crate) const fn description(self) -> &'static str {
        match self {
            Self::Objects => "OLE object count",
            Self::ObjectBytes => "OLE object bytes",
            Self::DirectoryEntries => "OLE directory entries",
            Self::Storages => "OLE storage count",
            Self::StorageDepth => "OLE storage depth",
            Self::Streams => "OLE stream count",
            Self::StreamBytes => "OLE stream bytes",
            Self::AggregateBytes => "OLE aggregate stream bytes",
        }
    }

    pub(crate) const fn exceeded(self, observed: u64, maximum: u64) -> Error {
        Error::Limit {
            kind: crate::LimitKind::OleObject(self),
            observed,
            maximum,
        }
    }

    pub(crate) const fn invalid_zero(self) -> Error {
        Error::InvalidLimit {
            kind: crate::LimitKind::OleObject(self),
            value: 0,
            maximum: match self {
                Self::ObjectBytes | Self::StreamBytes | Self::AggregateBytes => u64::MAX,
                _ => usize::MAX as u64,
            },
        }
    }
}

/// Per-edit configuration of an OLE-object edit.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct EditOptions {
    limits: Limits,
}

impl EditOptions {
    /// Default options: the default object capture bounds.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Select the object capture bounds.
    #[must_use]
    pub const fn limits(mut self, limits: Limits) -> Self {
        self.limits = limits;
        self
    }

    /// The selected object capture bounds.
    #[must_use]
    pub const fn limits_value(&self) -> Limits {
        self.limits
    }
}

impl crate::Package {
    /// The OLE-object and form-control read view of this package.
    ///
    /// # Errors
    ///
    /// Returns an error when the package is not a supported BIFF8 workbook.
    pub fn ole_objects(&self) -> Result<Snapshot> {
        let (_, bytes) = crate::scoped::capture(self)?;
        Snapshot::open(bytes, Limits::default())
    }

    /// Start an OLE-object edit of this package with the default bounds.
    ///
    /// # Errors
    ///
    /// Returns an error for an encrypted, protected or unsupported workbook.
    pub fn edit_ole_objects(&self) -> Result<super::Edit> {
        self.edit_ole_objects_with_options(EditOptions::default())
    }

    /// Start an OLE-object edit of this package under explicit bounds.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::edit_ole_objects`], or a limit
    /// error when the package exceeds `options`.
    pub fn edit_ole_objects_with_options(&self, options: EditOptions) -> Result<super::Edit> {
        if self.is_encrypted() {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::Protected,
                "edit OLE objects",
                "an encrypted workbook cannot be republished by a scoped edit",
            ));
        }
        Ok(super::Edit::lazy(
            Some(self.clone()),
            self.clone(),
            options.limits_value(),
        ))
    }

    /// Replay an OLE-object patch onto the exact package it was made from.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// source, or [`Error::Signed`] for a changed signed workbook.
    pub fn apply_ole_objects(&self, patch: &Patch) -> Result<Commit> {
        crate::scoped::replay_source(self, patch.before(), "edit OLE objects")?;
        let package = crate::scoped::publish(Some(self.clone()), patch.before(), patch.after())?;
        Ok(Commit::new(
            package,
            patch.clone(),
            Diagnostics::new(!patch.is_noop()),
        ))
    }
}
