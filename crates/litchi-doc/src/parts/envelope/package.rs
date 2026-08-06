//! DOC/OLE package ownership for the `fcMsoEnvelope` FIB range.

use super::codec;
use super::model::{Envelope, Message};
use super::transaction::{
    Error as TransactionError, Patch, Snapshot as TransactionSnapshot, Transaction,
};
use super::validation;
use crate::package::{Error as PackageError, Result};
use crate::parts::fib::FileInformationBlock;
use litchi_ole_common::object::{Editor as ObjectEditor, Limits, Patch as ObjectPatch, Targets};

/// Immutable package bytes plus the decoded envelope state.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Snapshot {
    bytes: Vec<u8>,
    envelope: TransactionSnapshot,
}

impl Snapshot {
    fn new(bytes: Vec<u8>, envelope: TransactionSnapshot) -> Self {
        Self { bytes, envelope }
    }

    /// The immutable semantic envelope snapshot.
    #[must_use]
    pub fn envelope(&self) -> &TransactionSnapshot {
        &self.envelope
    }

    /// The decoded envelope, when its FIB range is present.
    #[must_use]
    pub fn value(&self) -> Option<&Envelope> {
        self.envelope.envelope()
    }

    /// The supported Office message body, when present.
    #[must_use]
    pub fn message(&self) -> Option<&Message> {
        self.envelope.message()
    }

    /// Return the rendered package bytes.
    pub fn finish(&self) -> Result<Vec<u8>> {
        Ok(self.bytes.clone())
    }

    /// Start a semantic transaction from this package snapshot.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        self.envelope.edit()
    }
}

/// A package commit containing a semantic patch and a reversible CFB patch.
#[derive(Debug, Clone)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    package_patch: ObjectPatch,
}

impl Commit {
    /// The immutable post-edit package snapshot.
    #[must_use]
    pub fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// The reversible semantic envelope patch.
    #[must_use]
    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    /// The reversible whole-CFB byte patch.
    #[must_use]
    pub fn package_patch(&self) -> &ObjectPatch {
        &self.package_patch
    }

    /// Split the commit into its snapshot and both reversible patches.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch, ObjectPatch) {
        (self.snapshot, self.patch, self.package_patch)
    }
}

/// Transactional editor for the WordDocument and selected table stream.
#[derive(Debug, Clone)]
pub struct Editor {
    package: ObjectEditor,
    table_name: String,
    original: TransactionSnapshot,
    envelope: TransactionSnapshot,
    changed: bool,
}

impl Editor {
    /// Open a DOC package from complete CFB bytes.
    pub fn open(bytes: Vec<u8>) -> Result<Self> {
        Self::open_with_limits(bytes, Limits::default())
    }

    /// Open a DOC package with an explicit bounded OLE resource profile.
    pub fn open_with_limits(bytes: Vec<u8>, limits: Limits) -> Result<Self> {
        let targets = Targets::new([]).map_err(PackageError::from)?;
        let package = ObjectEditor::open(bytes, targets, limits).map_err(PackageError::from)?;
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
        let envelope = TransactionSnapshot::from_option(codec::parse_fib(&fib, table)?)?;
        Ok(Self {
            package,
            table_name: table_name.to_owned(),
            original: envelope.clone(),
            envelope,
            changed: false,
        })
    }

    /// The current immutable semantic envelope snapshot.
    #[must_use]
    pub fn envelope(&self) -> &TransactionSnapshot {
        &self.envelope
    }

    /// The current decoded envelope, when present.
    #[must_use]
    pub fn value(&self) -> Option<&Envelope> {
        self.envelope.envelope()
    }

    /// The current supported Office message body, when present.
    #[must_use]
    pub fn message(&self) -> Option<&Message> {
        self.envelope.message()
    }

    /// Whether a package edit has changed any stream bytes.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.changed
    }

    /// Start an independent semantic transaction.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        self.envelope.edit()
    }

    /// Set or replace the complete envelope FIB range.
    pub fn set(&mut self, envelope: Envelope) -> Result<Commit> {
        self.install(TransactionSnapshot::new(envelope)?)
    }

    /// Set or replace the supported Office message body.
    pub fn set_message(&mut self, message: Message) -> Result<Commit> {
        self.set(Envelope::from_message(message)?)
    }

    /// Replace an existing envelope, rejecting a missing FIB range.
    pub fn replace(&mut self, envelope: Envelope) -> Result<Commit> {
        if !self.envelope.is_present() {
            return Err(PackageError::InvalidFormat(
                "cannot replace missing MsoEnvelope metadata".into(),
            ));
        }
        self.set(envelope)
    }

    /// Apply and publish a clone-first semantic transaction atomically.
    pub fn apply(&mut self, transaction: Transaction) -> Result<Commit> {
        let commit = transaction.commit().map_err(transaction_error)?;
        if commit.patch().before() != &self.envelope {
            return Err(PackageError::InvalidFormat(
                "MsoEnvelope transaction snapshot conflict".into(),
            ));
        }
        self.install(commit.snapshot().clone())
    }

    /// Clone-first update of the supported Office message body.
    pub fn update_message<F>(&mut self, edit: F) -> Result<Commit>
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
    pub fn clear(&mut self) -> Result<Commit> {
        let mut transaction = self.edit();
        transaction.clear();
        self.apply(transaction)
    }

    /// Capture the current package as an immutable snapshot.
    pub fn snapshot(&self) -> Result<Snapshot> {
        let bytes = self.package.clone().finish().map_err(PackageError::from)?;
        Ok(Snapshot::new(bytes, self.envelope.clone()))
    }

    /// Finish the edit and return rendered DOC bytes.
    pub fn finish(self) -> Result<Vec<u8>> {
        self.package.finish().map_err(PackageError::from)
    }

    /// Commit the package as an immutable snapshot with reversible patches.
    pub fn commit(self) -> Result<Commit> {
        let patch = Patch::new(self.original, self.envelope.clone());
        let object_commit = self.package.commit().map_err(PackageError::from)?;
        let bytes = object_commit.patch().after().to_vec();
        Ok(Commit {
            snapshot: Snapshot::new(bytes, self.envelope),
            patch,
            package_patch: object_commit.into_patch(),
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
                .into_patch();
            return Ok(Commit {
                snapshot: self.snapshot()?,
                patch,
                package_patch,
            });
        }

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
        let package_patch = object_commit.into_patch();
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
