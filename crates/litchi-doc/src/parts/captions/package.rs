//! DOC/OLE package ownership for caption and `AutoCaption` metadata.

use super::codec::{AUTO_CAPTION_FIB_INDEX, CAPTION_FIB_INDEX};
use super::semantic::Tables;
use super::transaction::{
    Commit as TransactionCommit, Error as TransactionError, Snapshot as TransactionSnapshot,
    Transaction,
};
use super::validation;
use super::{AutoTable, LabelTable};
use crate::package::{Error as PackageError, Result};
use crate::parts::fib::FileInformationBlock;
use litchi_ole_common::object::{Editor as ObjectEditor, Limits, Patch as ObjectPatch, Targets};

/// An immutable DOC package snapshot carrying its caption metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Snapshot {
    bytes: Vec<u8>,
    captions: TransactionSnapshot,
}

impl Snapshot {
    fn new(bytes: Vec<u8>, captions: TransactionSnapshot) -> Self {
        Self { bytes, captions }
    }

    /// Returns the immutable caption snapshot.
    #[must_use]
    pub fn captions(&self) -> &TransactionSnapshot {
        &self.captions
    }

    /// Returns the paired caption tables.
    #[must_use]
    pub fn tables(&self) -> &Tables {
        self.captions.tables()
    }

    /// Returns the rendered package bytes.
    pub fn finish(&self) -> Result<Vec<u8>> {
        Ok(self.bytes.clone())
    }

    /// Starts an independent semantic transaction.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        self.captions.edit()
    }
}

/// A package commit containing semantic and reversible OLE byte patches.
#[derive(Debug, Clone)]
pub struct Commit {
    snapshot: Snapshot,
    patch: TransactionCommit,
    package_patch: ObjectPatch,
}

impl Commit {
    /// The immutable post-edit package snapshot.
    #[must_use]
    pub fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// The reversible semantic caption patch.
    #[must_use]
    pub fn patch(&self) -> &TransactionCommit {
        &self.patch
    }

    /// The reversible whole-CFB byte patch.
    #[must_use]
    pub fn package_patch(&self) -> &ObjectPatch {
        &self.package_patch
    }

    /// Splits the commit into its snapshot, semantic patch, and byte patch.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, TransactionCommit, ObjectPatch) {
        (self.snapshot, self.patch, self.package_patch)
    }
}

/// Transactional package editor for the two Normal-template caption ranges.
#[derive(Debug, Clone)]
pub struct Editor {
    package: ObjectEditor,
    table_name: String,
    original: TransactionSnapshot,
    captions: TransactionSnapshot,
    changed: bool,
}

impl Editor {
    /// Opens an unencrypted Word 97+ Normal template from CFB bytes.
    pub fn open(bytes: Vec<u8>) -> Result<Self> {
        Self::open_with_limits(bytes, Limits::default())
    }

    /// Opens a package with an explicit bounded OLE resource profile.
    pub fn open_with_limits(bytes: Vec<u8>, limits: Limits) -> Result<Self> {
        let package =
            ObjectEditor::open(bytes, Targets::new([]).map_err(PackageError::from)?, limits)
                .map_err(PackageError::from)?;
        let word = package
            .stream(&["WordDocument".to_owned()])
            .ok_or_else(|| PackageError::StreamNotFound("WordDocument".into()))?;
        let fib = FileInformationBlock::parse(word)?;
        validation::package_fib(&fib)?;
        let table_name = if fib.which_table_stream() {
            "1Table"
        } else {
            "0Table"
        };
        let table = package
            .stream(&[table_name.to_owned()])
            .ok_or_else(|| PackageError::StreamNotFound(table_name.to_owned()))?;
        let captions = TransactionSnapshot::new(Tables::parse(&fib, table)?)?;
        Ok(Self {
            package,
            table_name: table_name.to_owned(),
            original: captions.clone(),
            captions,
            changed: false,
        })
    }

    /// Returns the current immutable caption snapshot.
    #[must_use]
    pub fn captions(&self) -> &TransactionSnapshot {
        &self.captions
    }

    /// Returns the current paired caption tables.
    #[must_use]
    pub fn tables(&self) -> &Tables {
        self.captions.tables()
    }

    /// Whether this editor has published a changed package state.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.changed
    }

    /// Starts an independent semantic transaction.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        self.captions.edit()
    }

    /// Sets or replaces both caption tables atomically.
    pub fn set(&mut self, tables: Tables) -> Result<Commit> {
        self.install(TransactionSnapshot::new(tables)?)
    }

    /// Replaces existing caption metadata, rejecting an absent pair.
    pub fn replace(&mut self, tables: Tables) -> Result<Commit> {
        if !self.captions.is_present() {
            return Err(PackageError::InvalidFormat(
                "cannot replace missing caption metadata".into(),
            ));
        }
        self.set(tables)
    }

    /// Applies a semantic transaction and publishes it atomically.
    pub fn apply(&mut self, transaction: Transaction) -> Result<Commit> {
        let commit = transaction.commit().map_err(transaction_error)?;
        if commit.before() != &self.captions {
            return Err(PackageError::InvalidFormat(
                "caption transaction snapshot conflict".into(),
            ));
        }
        self.install(commit.snapshot().clone())
    }

    /// Replaces the label table while retaining automatic-caption rules.
    pub fn replace_labels(&mut self, labels: LabelTable) -> Result<Commit> {
        let mut transaction = self.edit();
        transaction
            .replace_labels(labels)
            .map_err(transaction_error)?;
        self.apply(transaction)
    }

    /// Replaces automatic-caption rules while retaining labels.
    pub fn replace_auto(&mut self, auto: AutoTable) -> Result<Commit> {
        let mut transaction = self.edit();
        transaction.replace_auto(auto).map_err(transaction_error)?;
        self.apply(transaction)
    }

    /// Removes only the label range, rejecting rules that would dangle.
    pub fn clear_labels(&mut self) -> Result<Commit> {
        let mut transaction = self.edit();
        transaction.clear_labels().map_err(transaction_error)?;
        self.apply(transaction)
    }

    /// Removes only the automatic-caption range while retaining labels.
    pub fn clear_auto(&mut self) -> Result<Commit> {
        let mut transaction = self.edit();
        transaction.clear_auto().map_err(transaction_error)?;
        self.apply(transaction)
    }

    /// Removes both caption table ranges while preserving their old bytes.
    pub fn clear(&mut self) -> Result<Commit> {
        let mut transaction = self.edit();
        transaction.clear();
        self.apply(transaction)
    }

    /// Captures the current package as an immutable snapshot.
    pub fn snapshot(&self) -> Result<Snapshot> {
        let bytes = self.package.clone().finish().map_err(PackageError::from)?;
        Ok(Snapshot::new(bytes, self.captions.clone()))
    }

    /// Finishes the edit and returns rendered DOC bytes.
    pub fn finish(self) -> Result<Vec<u8>> {
        self.package.finish().map_err(PackageError::from)
    }

    /// Commits the package as an immutable snapshot with reversible patches.
    pub fn commit(self) -> Result<Commit> {
        let patch = TransactionCommit::new(self.original, self.captions.clone());
        let object_commit = self.package.commit().map_err(PackageError::from)?;
        let bytes = object_commit.patch().after().to_vec();
        Ok(Commit {
            snapshot: Snapshot::new(bytes, self.captions),
            patch,
            package_patch: object_commit.into_patch(),
        })
    }

    fn install(&mut self, snapshot: TransactionSnapshot) -> Result<Commit> {
        let patch = TransactionCommit::new(self.captions.clone(), snapshot.clone());
        if patch.is_noop() {
            let package_patch = self
                .package
                .clone()
                .commit()
                .map_err(PackageError::from)?
                .into_patch();
            let package_snapshot = self.snapshot()?;
            return Ok(Commit {
                snapshot: package_snapshot,
                patch,
                package_patch,
            });
        }

        let mut candidate = self.clone();
        candidate.write_snapshot(&snapshot)?;
        candidate.captions = snapshot;
        candidate.changed = true;
        let package_commit = candidate
            .package
            .clone()
            .commit()
            .map_err(PackageError::from)?;
        let bytes = package_commit.patch().after().to_vec();
        let package_patch = package_commit.into_patch();
        let package_snapshot = Snapshot::new(bytes, candidate.captions.clone());
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
        let word = self
            .package
            .stream(&word_path)
            .ok_or_else(|| PackageError::StreamNotFound("WordDocument".into()))?
            .to_vec();
        let table = self
            .package
            .stream(&table_path)
            .ok_or_else(|| PackageError::StreamNotFound(self.table_name.clone()))?
            .to_vec();
        let fib = FileInformationBlock::parse(&word)?;
        let mut next_word = word;
        let mut next_table = table;

        append_range(
            &mut next_word,
            &mut next_table,
            validation::pointer_location(&fib, CAPTION_FIB_INDEX)?,
            snapshot.labels().map(LabelTable::to_bytes).transpose()?,
            "SttbfCaption",
        )?;
        append_range(
            &mut next_word,
            &mut next_table,
            validation::pointer_location(&fib, AUTO_CAPTION_FIB_INDEX)?,
            snapshot.auto().map(AutoTable::to_bytes).transpose()?,
            "SttbfAutoCaption",
        )?;

        let reparsed_fib = FileInformationBlock::parse(&next_word)?;
        let reparsed = Tables::parse(&reparsed_fib, &next_table)?;
        if &reparsed != snapshot.tables() {
            return Err(PackageError::Corrupted(
                "caption snapshot failed FIB/table-stream round-trip validation".into(),
            ));
        }

        let mut package = self.package.clone();
        package
            .put_stream(&table_path, next_table)
            .map_err(PackageError::from)?;
        package
            .put_stream(&word_path, next_word)
            .map_err(PackageError::from)?;
        self.package = package;
        Ok(())
    }
}

fn append_range(
    word: &mut [u8],
    table: &mut Vec<u8>,
    pointer: usize,
    payload: Option<Vec<u8>>,
    name: &str,
) -> Result<()> {
    let end = pointer
        .checked_add(8)
        .ok_or_else(|| PackageError::Corrupted(format!("{name} FIB pointer overflows")))?;
    if end > word.len() {
        return Err(PackageError::Corrupted(format!(
            "{name} FIB pointer extends beyond WordDocument"
        )));
    }
    let Some(payload) = payload else {
        word[pointer..end].fill(0);
        return Ok(());
    };
    let offset = u32::try_from(table.len())
        .map_err(|_| PackageError::Corrupted(format!("{name} offset exceeds u32::MAX")))?;
    let length = u32::try_from(payload.len())
        .map_err(|_| PackageError::Corrupted(format!("{name} length exceeds u32::MAX")))?;
    table.extend_from_slice(&payload);
    word[pointer..pointer + 4].copy_from_slice(&offset.to_le_bytes());
    word[pointer + 4..end].copy_from_slice(&length.to_le_bytes());
    Ok(())
}

fn transaction_error(error: TransactionError) -> PackageError {
    PackageError::InvalidFormat(format!("caption transaction failed: {error}"))
}
