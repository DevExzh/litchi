//! DOC/OLE package ownership for the `fcMsoEnvelope` FIB range.

#![allow(
    dead_code,
    reason = "the byte-level editor keeps its record-level API for its unit tests; the package family drives a subset"
)]

use super::codec;
use super::model::{Envelope, Message};
use super::transaction::{
    Edit as RecordEdit, Error as TransactionError, Patch, Snapshot as TransactionSnapshot,
};
use super::validation;
use crate::package::{Error as PackageError, Result};
use crate::parts::fib::FileInformationBlock;
use crate::parts::protection::{EditProtection, PackagePatch, ProtectionPolicy, classify};
use litchi_ole_common::object::{
    Edit as ObjectEditor, Limits, OpenOptions, Package as ObjectPackage, Targets,
};

/// Immutable package bytes plus the decoded envelope state.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Snapshot {
    bytes: Vec<u8>,
    envelope: TransactionSnapshot,
}

impl Snapshot {
    fn new(bytes: Vec<u8>, envelope: TransactionSnapshot) -> Self {
        Self { bytes, envelope }
    }

    /// The immutable semantic envelope snapshot.
    #[must_use]
    pub(crate) fn envelope(&self) -> &TransactionSnapshot {
        &self.envelope
    }

    /// The decoded envelope, when its FIB range is present.
    #[must_use]
    pub(crate) fn value(&self) -> Option<&Envelope> {
        self.envelope.envelope()
    }

    /// The supported Office message body, when present.
    #[must_use]
    pub(crate) fn message(&self) -> Option<&Message> {
        self.envelope.message()
    }

    /// Return the rendered package bytes.
    pub(crate) fn finish(&self) -> Result<Vec<u8>> {
        Ok(self.bytes.clone())
    }

    /// Start a semantic transaction from this package snapshot.
    #[must_use]
    pub(crate) fn edit(&self) -> RecordEdit {
        self.envelope.edit()
    }
}

/// A package commit containing a semantic patch and a reversible CFB patch.
#[derive(Debug, Clone)]
pub(crate) struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    package_patch: PackagePatch,
}

impl Commit {
    /// The immutable post-edit package snapshot.
    #[must_use]
    pub(crate) fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// The reversible semantic envelope patch.
    #[must_use]
    pub(crate) fn patch(&self) -> &Patch {
        &self.patch
    }

    /// The reversible whole-CFB byte patch.
    #[must_use]
    pub(crate) fn package_patch(&self) -> &PackagePatch {
        &self.package_patch
    }

    /// Split the commit into its snapshot and both reversible patches.
    #[must_use]
    pub(crate) fn into_parts(self) -> (Snapshot, Patch, PackagePatch) {
        (self.snapshot, self.patch, self.package_patch)
    }
}

/// Transactional editor for the `WordDocument` and selected table stream.
#[derive(Debug, Clone)]
pub(crate) struct Editor {
    package: ObjectEditor,
    table_name: String,
    original: TransactionSnapshot,
    envelope: TransactionSnapshot,
    protection: EditProtection,
    protection_policy: ProtectionPolicy,
    changed: bool,
}

impl Editor {
    /// Open a DOC package from complete CFB bytes.
    pub(crate) fn open(bytes: Vec<u8>) -> Result<Self> {
        Self::open_bounded(bytes, Limits::default(), ProtectionPolicy::default())
    }

    /// Open a package with an explicit policy for changed protected content.
    pub(crate) fn open_with_policy(bytes: Vec<u8>, policy: ProtectionPolicy) -> Result<Self> {
        Self::open_bounded(bytes, Limits::default(), policy)
    }

    /// Opens a package under one edit profile: its revision limits bound the
    /// compound-file parse and its protection policy governs changed
    /// protected content.
    pub(crate) fn open_with_options(bytes: Vec<u8>, options: &crate::EditOptions) -> Result<Self> {
        Self::open_bounded(
            bytes,
            options.revision_limits_value(),
            options.protection_policy_value(),
        )
    }

    /// Open a package with both bounded resources and an explicit protection policy.
    fn open_bounded(
        bytes: Vec<u8>,
        limits: Limits,
        protection_policy: ProtectionPolicy,
    ) -> Result<Self> {
        let targets = Targets::new([]).map_err(PackageError::from)?;
        let package = ObjectPackage::open_with_options(
            bytes,
            OpenOptions::new().targets(targets).limits(limits),
        )
        .map_err(PackageError::from)?
        .edit()
        .map_err(PackageError::from)?;
        let word_path = vec!["WordDocument".to_owned()];
        let word = package
            .stream(&word_path)
            .ok_or_else(|| PackageError::StreamNotFound("WordDocument".into()))?;
        let fib = FileInformationBlock::parse(word)?;
        validation::package_fib(&fib)?;
        let table_name = if fib.which_table_stream() {
            "1Table"
        } else {
            "0Table"
        };
        let table_path = vec![table_name.to_owned()];
        let table = package
            .stream(&table_path)
            .ok_or_else(|| PackageError::StreamNotFound(table_name.into()))?;
        let protection = classify(&fib, table)?;
        let envelope = TransactionSnapshot::from_option(codec::parse_fib(&fib, table)?)?;
        Ok(Self {
            package,
            table_name: table_name.to_owned(),
            original: envelope.clone(),
            envelope,
            protection,
            protection_policy,
            changed: false,
        })
    }

    /// The current immutable semantic envelope snapshot.
    #[must_use]
    pub(crate) fn envelope(&self) -> &TransactionSnapshot {
        &self.envelope
    }

    /// The current decoded envelope, when present.
    #[must_use]
    pub(crate) fn value(&self) -> Option<&Envelope> {
        self.envelope.envelope()
    }

    /// The current supported Office message body, when present.
    #[must_use]
    pub(crate) fn message(&self) -> Option<&Message> {
        self.envelope.message()
    }

    /// Whether a package edit has changed any stream bytes.
    #[must_use]
    pub(crate) fn is_changed(&self) -> bool {
        self.changed
    }

    /// Start an independent semantic transaction.
    #[must_use]
    pub(crate) fn edit(&self) -> RecordEdit {
        self.envelope.edit()
    }

    /// Set or replace the complete envelope FIB range.
    pub(crate) fn set(&mut self, envelope: Envelope) -> Result<Commit> {
        self.install(TransactionSnapshot::new(envelope)?)
    }

    /// Set or replace the supported Office message body.
    pub(crate) fn set_message(&mut self, message: Message) -> Result<Commit> {
        self.set(Envelope::from_message(message)?)
    }

    /// Replace an existing envelope, rejecting a missing FIB range.
    pub(crate) fn replace(&mut self, envelope: Envelope) -> Result<Commit> {
        if !self.envelope.is_present() {
            return Err(PackageError::InvalidFormat(
                "cannot replace missing MsoEnvelope metadata".into(),
            ));
        }
        self.set(envelope)
    }

    /// Apply and publish a clone-first semantic transaction atomically.
    pub(crate) fn apply(&mut self, transaction: RecordEdit) -> Result<Commit> {
        let commit = transaction.commit().map_err(transaction_error)?;
        if commit.patch().before() != &self.envelope {
            return Err(PackageError::InvalidFormat(
                "MsoEnvelope transaction snapshot conflict".into(),
            ));
        }
        self.install(commit.package().clone())
    }

    /// Clone-first update of the supported Office message body.
    pub(crate) fn update_message<F>(&mut self, edit: F) -> Result<Commit>
    where
        F: FnOnce(&mut Message),
    {
        let mut transaction = self.edit();
        transaction
            .update_message(edit)
            .map_err(transaction_error)?;
        self.apply(transaction)
    }

    /// Clear the complete envelope FIB range.
    pub(crate) fn clear(&mut self) -> Result<Commit> {
        let mut transaction = self.edit();
        transaction.clear();
        self.apply(transaction)
    }

    /// Capture the current package as an immutable snapshot.
    pub(crate) fn snapshot(&self) -> Result<Snapshot> {
        let bytes = crate::package::object_edit_bytes(&self.package)?;
        Ok(Snapshot::new(bytes, self.envelope.clone()))
    }

    /// Finish the edit and return rendered DOC bytes.
    pub(crate) fn finish(self) -> Result<Vec<u8>> {
        if self.changed {
            self.protection_policy.authorize(self.protection)?;
        }
        crate::package::object_edit_bytes(&self.package)
    }

    /// Commit the package as an immutable snapshot with reversible patches.
    pub(crate) fn commit(self) -> Result<Commit> {
        if self.changed {
            self.protection_policy.authorize(self.protection)?;
        }
        let patch = Patch::new(self.original, self.envelope.clone());
        let object_commit = self.package.commit().map_err(PackageError::from)?;
        let bytes = object_commit.patch().after().to_vec();
        Ok(Commit {
            snapshot: Snapshot::new(bytes, self.envelope),
            patch,
            package_patch: PackagePatch::new(object_commit.into_parts().1, self.protection),
        })
    }

    fn install(&mut self, snapshot: TransactionSnapshot) -> Result<Commit> {
        let patch = Patch::new(self.envelope.clone(), snapshot.clone());
        if patch.is_noop() {
            let package_patch = self
                .package
                .clone()
                .commit()
                .map_err(PackageError::from)?
                .into_parts()
                .1;
            let package_patch = PackagePatch::new(package_patch, self.protection);
            return Ok(Commit {
                snapshot: self.snapshot()?,
                patch,
                package_patch,
            });
        }

        self.protection_policy.authorize(self.protection)?;

        // All package mutation happens on a clone. The source editor is only
        // replaced after the candidate has been reparsed and round-tripped.
        let mut candidate = self.clone();
        candidate.write_snapshot(&snapshot)?;
        candidate.envelope = snapshot;
        candidate.changed = true;
        let object_commit = candidate
            .package
            .clone()
            .commit()
            .map_err(PackageError::from)?;
        let bytes = object_commit.patch().after().to_vec();
        let package_patch = PackagePatch::new(object_commit.into_parts().1, candidate.protection);
        let package_snapshot = Snapshot::new(bytes, candidate.envelope.clone());
        *self = candidate;
        Ok(Commit {
            snapshot: package_snapshot,
            patch,
            package_patch,
        })
    }

    fn write_snapshot(&mut self, snapshot: &TransactionSnapshot) -> Result<()> {
        let word_path = vec!["WordDocument".to_owned()];
        let table_path = vec![self.table_name.clone()];
        let mut word = self
            .package
            .stream(&word_path)
            .ok_or_else(|| PackageError::StreamNotFound("WordDocument".into()))?
            .to_vec();
        let mut table = self
            .package
            .stream(&table_path)
            .ok_or_else(|| PackageError::StreamNotFound(self.table_name.clone()))?
            .to_vec();
        let fib = FileInformationBlock::parse(&word)?;
        let pointer = validation::pointer_location(&fib)?;
        if let Some(envelope) = snapshot.envelope() {
            let payload = codec::write(envelope)?;
            let offset = u32::try_from(table.len())
                .map_err(|_| PackageError::Corrupted("table stream exceeds u32::MAX".into()))?;
            let end = table
                .len()
                .checked_add(payload.len())
                .ok_or_else(|| PackageError::Corrupted("table stream size overflows".into()))?;
            if end > u32::MAX as usize {
                return Err(PackageError::Corrupted(
                    "table stream exceeds u32::MAX after appending MsoEnvelope".into(),
                ));
            }
            let length = u32::try_from(payload.len()).map_err(|_| {
                PackageError::Corrupted("MsoEnvelope payload exceeds u32::MAX".into())
            })?;
            table.extend_from_slice(&payload);
            word[pointer..pointer + 4].copy_from_slice(&offset.to_le_bytes());
            word[pointer + 4..pointer + 8].copy_from_slice(&length.to_le_bytes());
        } else {
            word[pointer..pointer + 8].fill(0);
        }

        let reparsed_fib = FileInformationBlock::parse(&word)?;
        let reparsed = TransactionSnapshot::from_option(codec::parse_fib(&reparsed_fib, &table)?)?;
        if &reparsed != snapshot {
            return Err(PackageError::Corrupted(
                "MsoEnvelope snapshot failed FIB/table-stream round-trip validation".into(),
            ));
        }

        let mut package = self.package.clone();
        package
            .put_stream(&table_path, table)
            .map_err(PackageError::from)?;
        package
            .put_stream(&word_path, word)
            .map_err(PackageError::from)?;
        self.package = package;
        Ok(())
    }
}

fn transaction_error(error: TransactionError) -> PackageError {
    PackageError::InvalidFormat(format!("MsoEnvelope transaction failed: {error}"))
}
