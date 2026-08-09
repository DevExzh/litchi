//! DOC/OLE package ownership for `SttbfAtnBkmk`.

use super::codec;
use super::model::{Tag, Tags};
use super::transaction::{
    Error as TransactionError, Patch, Snapshot as TransactionSnapshot, Transaction,
};
use super::validation;
use crate::package::{Error as PackageError, Result};
use crate::parts::fib::FileInformationBlock;
use litchi_ole_common::object::{Editor as ObjectEditor, Limits, Patch as ObjectPatch, Targets};

/// Immutable package bytes plus the decoded annotation-tag state.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Snapshot {
    bytes: Vec<u8>,
    tags: TransactionSnapshot,
}

impl Snapshot {
    fn new(bytes: Vec<u8>, tags: TransactionSnapshot) -> Self {
        Self { bytes, tags }
    }

    /// The immutable semantic table state.
    #[must_use]
    pub fn tags(&self) -> &TransactionSnapshot {
        &self.tags
    }

    /// The decoded tags, when the FIB range is present.
    #[must_use]
    pub fn value(&self) -> Option<&Tags> {
        self.tags.tags()
    }

    /// Return rendered package bytes.
    pub fn finish(&self) -> Result<Vec<u8>> {
        Ok(self.bytes.clone())
    }

    /// Start an edit from this package snapshot.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        self.tags.edit()
    }
}

/// A package commit containing semantic and whole-CFB patches.
#[derive(Debug, Clone)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    package_patch: ObjectPatch,
}

impl Commit {
    /// Immutable post-edit package snapshot.
    #[must_use]
    pub fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Reversible semantic patch.
    #[must_use]
    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Reversible whole-CFB patch.
    #[must_use]
    pub fn package_patch(&self) -> &ObjectPatch {
        &self.package_patch
    }
}

/// Transactional editor for the `WordDocument` and selected table stream.
#[derive(Debug, Clone)]
pub struct Editor {
    package: ObjectEditor,
    table_name: String,
    original: TransactionSnapshot,
    tags: TransactionSnapshot,
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
            .ok_or_else(|| PackageError::StreamNotFound(table_name.to_owned()))?;
        let tags = TransactionSnapshot::from_option(codec::parse(&fib, table)?)?;
        Ok(Self {
            package,
            table_name: table_name.to_owned(),
            original: tags.clone(),
            tags,
            changed: false,
        })
    }

    /// Current immutable semantic table state.
    #[must_use]
    pub fn tags(&self) -> &TransactionSnapshot {
        &self.tags
    }

    /// Decoded tags, when present.
    #[must_use]
    pub fn value(&self) -> Option<&Tags> {
        self.tags.tags()
    }

    /// Whether this editor has published changed package bytes.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.changed
    }

    /// Start an independent semantic transaction.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        self.tags.edit()
    }

    /// Set or replace the complete present table.
    pub fn set(&mut self, tags: Tags) -> Result<Commit> {
        self.install(TransactionSnapshot::new(tags)?)
    }

    /// Replace the table, requiring an existing FIB range.
    pub fn replace(&mut self, tags: Tags) -> Result<Commit> {
        if !self.tags.is_present() {
            return Err(PackageError::InvalidFormat(
                "cannot replace missing SttbfAtnBkmk metadata".into(),
            ));
        }
        self.set(tags)
    }

    /// Apply and publish a semantic transaction atomically.
    pub fn apply(&mut self, transaction: Transaction) -> Result<Commit> {
        let commit = transaction.commit().map_err(transaction_error)?;
        if commit.patch().before() != &self.tags {
            return Err(PackageError::InvalidFormat(
                "SttbfAtnBkmk transaction snapshot conflict".into(),
            ));
        }
        self.install(commit.snapshot().clone())
    }

    /// Replace one annotation tag and publish atomically.
    pub fn replace_entry(&mut self, index: usize, tag: Tag) -> Result<Commit> {
        let mut transaction = self.edit();
        transaction
            .replace_entry(index, tag)
            .map_err(transaction_error)?;
        self.apply(transaction)
    }

    /// Insert one annotation tag and publish atomically.
    pub fn insert(&mut self, index: usize, tag: Tag) -> Result<Commit> {
        let mut transaction = self.edit();
        transaction.insert(index, tag).map_err(transaction_error)?;
        self.apply(transaction)
    }

    /// Remove one annotation tag and publish atomically.
    pub fn remove(&mut self, index: usize) -> Result<Commit> {
        let mut transaction = self.edit();
        transaction.remove(index).map_err(transaction_error)?;
        self.apply(transaction)
    }

    /// Remove the complete FIB range.
    pub fn clear(&mut self) -> Result<Commit> {
        let mut transaction = self.edit();
        transaction.clear();
        self.apply(transaction)
    }

    /// Capture the current package as an immutable snapshot.
    pub fn snapshot(&self) -> Result<Snapshot> {
        let bytes = self.package.clone().finish().map_err(PackageError::from)?;
        Ok(Snapshot::new(bytes, self.tags.clone()))
    }

    /// Finish and return DOC bytes.
    pub fn finish(self) -> Result<Vec<u8>> {
        self.package.finish().map_err(PackageError::from)
    }

    /// Commit semantic and byte-level reversible patches.
    pub fn commit(self) -> Result<Commit> {
        let patch = Patch::new(self.original, self.tags.clone());
        let object_commit = self.package.commit().map_err(PackageError::from)?;
        let bytes = object_commit.patch().after().to_vec();
        Ok(Commit {
            snapshot: Snapshot::new(bytes, self.tags),
            patch,
            package_patch: object_commit.into_patch(),
        })
    }

    fn install(&mut self, snapshot: TransactionSnapshot) -> Result<Commit> {
        let patch = Patch::new(self.tags.clone(), snapshot.clone());
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

        let mut candidate = self.clone();
        candidate.write_snapshot(&snapshot)?;
        candidate.tags = snapshot;
        candidate.changed = true;
        let object_commit = candidate
            .package
            .clone()
            .commit()
            .map_err(PackageError::from)?;
        let bytes = object_commit.patch().after().to_vec();
        let package_patch = object_commit.into_patch();
        let package_snapshot = Snapshot::new(bytes, candidate.tags.clone());
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
        if let Some(tags) = snapshot.tags() {
            let payload = codec::to_bytes(tags)?;
            let offset = u32::try_from(table.len())
                .map_err(|_| PackageError::Corrupted("table stream exceeds u32::MAX".into()))?;
            let length = u32::try_from(payload.len()).map_err(|_| {
                PackageError::Corrupted("SttbfAtnBkmk payload exceeds u32::MAX".into())
            })?;
            table.extend_from_slice(&payload);
            word[pointer..pointer + 4].copy_from_slice(&offset.to_le_bytes());
            word[pointer + 4..pointer + 8].copy_from_slice(&length.to_le_bytes());
        } else {
            word[pointer..pointer + 8].fill(0);
        }

        let reparsed_fib = FileInformationBlock::parse(&word)?;
        let reparsed = codec::parse(&reparsed_fib, &table)?;
        if reparsed.as_ref() != snapshot.tags() {
            return Err(PackageError::Corrupted(
                "SttbfAtnBkmk snapshot failed FIB/table-stream round-trip validation".into(),
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
    PackageError::InvalidFormat(format!("SttbfAtnBkmk transaction failed: {error}"))
}
