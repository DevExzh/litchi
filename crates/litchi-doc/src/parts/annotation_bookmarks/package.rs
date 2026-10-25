//! DOC/OLE package ownership for `SttbfAtnBkmk`.

#![allow(
    dead_code,
    reason = "the byte-level editor keeps its record-level API for its unit tests; the package family drives a subset"
)]

use super::codec;
use super::model::{Tag, Tags};
use super::transaction::{
    Edit as RecordEdit, Error as TransactionError, Patch, Snapshot as TransactionSnapshot,
};
use super::validation;
use crate::package::{Error as PackageError, Result, SourceLimits};
use crate::parts::fib::FileInformationBlock;
use crate::parts::protection::{EditProtection, PackagePatch, ProtectionPolicy, classify};
use litchi_ole_common::object::{
    Edit as ObjectEditor, Limits, OpenOptions, Package as ObjectPackage, Targets,
};

/// Immutable package bytes plus the decoded annotation-tag state.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Snapshot {
    bytes: Vec<u8>,
    tags: TransactionSnapshot,
}

impl Snapshot {
    fn new(bytes: Vec<u8>, tags: TransactionSnapshot) -> Self {
        Self { bytes, tags }
    }

    /// The immutable semantic table state.
    #[must_use]
    pub(crate) fn tags(&self) -> &TransactionSnapshot {
        &self.tags
    }

    /// The decoded tags, when the FIB range is present.
    #[must_use]
    pub(crate) fn value(&self) -> Option<&Tags> {
        self.tags.tags()
    }

    /// Return rendered package bytes.
    pub(crate) fn finish(&self) -> Result<Vec<u8>> {
        Ok(self.bytes.clone())
    }

    /// Start an edit from this package snapshot.
    #[must_use]
    pub(crate) fn edit(&self) -> RecordEdit {
        self.tags.edit()
    }
}

/// A package commit containing semantic and whole-CFB patches.
#[derive(Debug, Clone)]
pub(crate) struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    package_patch: PackagePatch,
}

impl Commit {
    /// Immutable post-edit package snapshot.
    #[must_use]
    pub(crate) fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Reversible semantic patch.
    #[must_use]
    pub(crate) fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Reversible whole-CFB patch.
    #[must_use]
    pub(crate) fn package_patch(&self) -> &PackagePatch {
        &self.package_patch
    }
}

/// Transactional editor for the `WordDocument` and selected table stream.
#[derive(Debug, Clone)]
pub(crate) struct Editor {
    package: ObjectEditor,
    table_name: String,
    source_limits: SourceLimits,
    original: TransactionSnapshot,
    tags: TransactionSnapshot,
    protection: EditProtection,
    protection_policy: ProtectionPolicy,
    changed: bool,
}

impl Editor {
    /// Open a DOC package from complete CFB bytes.
    pub(crate) fn open(bytes: Vec<u8>) -> Result<Self> {
        Self::open_bounded(
            bytes,
            SourceLimits::default(),
            Limits::default(),
            ProtectionPolicy::default(),
        )
    }

    /// Open a package with an explicit policy for changed protected content.
    pub(crate) fn open_with_policy(bytes: Vec<u8>, policy: ProtectionPolicy) -> Result<Self> {
        Self::open_bounded(bytes, SourceLimits::default(), Limits::default(), policy)
    }

    /// Opens a package under one edit profile: its revision limits bound the
    /// compound-file parse and its protection policy governs changed
    /// protected content.
    pub(crate) fn open_with_options(
        bytes: Vec<u8>,
        source_limits: SourceLimits,
        options: &crate::EditOptions,
    ) -> Result<Self> {
        Self::open_bounded(
            bytes,
            source_limits,
            options.revision_limits_value(),
            options.protection_policy_value(),
        )
    }

    /// Open a package with both bounded resources and an explicit protection policy.
    fn open_bounded(
        bytes: Vec<u8>,
        source_limits: SourceLimits,
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
            .ok_or_else(|| PackageError::StreamNotFound(table_name.to_owned()))?;
        let protection = classify(&fib, table)?;
        let tags = TransactionSnapshot::from_option(codec::parse(&fib, table, &source_limits)?)?;
        Ok(Self {
            package,
            table_name: table_name.to_owned(),
            source_limits,
            original: tags.clone(),
            tags,
            protection,
            protection_policy,
            changed: false,
        })
    }

    /// Current immutable semantic table state.
    #[must_use]
    pub(crate) fn tags(&self) -> &TransactionSnapshot {
        &self.tags
    }

    /// Decoded tags, when present.
    #[must_use]
    pub(crate) fn value(&self) -> Option<&Tags> {
        self.tags.tags()
    }

    /// Whether this editor has published changed package bytes.
    #[must_use]
    pub(crate) fn is_changed(&self) -> bool {
        self.changed
    }

    /// Start an independent semantic transaction.
    #[must_use]
    pub(crate) fn edit(&self) -> RecordEdit {
        self.tags.edit()
    }

    /// Set or replace the complete present table.
    pub(crate) fn set(&mut self, tags: Tags) -> Result<Commit> {
        self.install(TransactionSnapshot::new(tags)?)
    }

    /// Replace the table, requiring an existing FIB range.
    pub(crate) fn replace(&mut self, tags: Tags) -> Result<Commit> {
        if !self.tags.is_present() {
            return Err(PackageError::InvalidFormat(
                "cannot replace missing SttbfAtnBkmk metadata".into(),
            ));
        }
        self.set(tags)
    }

    /// Apply and publish a semantic transaction atomically.
    pub(crate) fn apply(&mut self, transaction: RecordEdit) -> Result<Commit> {
        let commit = transaction.commit().map_err(transaction_error)?;
        if commit.patch().before() != &self.tags {
            return Err(PackageError::InvalidFormat(
                "SttbfAtnBkmk transaction snapshot conflict".into(),
            ));
        }
        self.install(commit.package().clone())
    }

    /// Replace one annotation tag and publish atomically.
    pub(crate) fn replace_entry(&mut self, index: usize, tag: Tag) -> Result<Commit> {
        let mut transaction = self.edit();
        transaction
            .replace_entry(index, tag)
            .map_err(transaction_error)?;
        self.apply(transaction)
    }

    /// Insert one annotation tag and publish atomically.
    pub(crate) fn insert(&mut self, index: usize, tag: Tag) -> Result<Commit> {
        let mut transaction = self.edit();
        transaction.insert(index, tag).map_err(transaction_error)?;
        self.apply(transaction)
    }

    /// Remove one annotation tag and publish atomically.
    pub(crate) fn remove(&mut self, index: usize) -> Result<Commit> {
        let mut transaction = self.edit();
        transaction.remove(index).map_err(transaction_error)?;
        self.apply(transaction)
    }

    /// Remove the complete FIB range.
    pub(crate) fn clear(&mut self) -> Result<Commit> {
        let mut transaction = self.edit();
        transaction.clear();
        self.apply(transaction)
    }

    /// Capture the current package as an immutable snapshot.
    pub(crate) fn snapshot(&self) -> Result<Snapshot> {
        let bytes = crate::package::object_edit_bytes(&self.package)?;
        Ok(Snapshot::new(bytes, self.tags.clone()))
    }

    /// Finish and return DOC bytes.
    pub(crate) fn finish(self) -> Result<Vec<u8>> {
        if self.changed {
            self.protection_policy.authorize(self.protection)?;
        }
        crate::package::object_edit_bytes(&self.package)
    }

    /// Commit semantic and byte-level reversible patches.
    pub(crate) fn commit(self) -> Result<Commit> {
        if self.changed {
            self.protection_policy.authorize(self.protection)?;
        }
        let patch = Patch::new(self.original, self.tags.clone());
        let object_commit = self.package.commit().map_err(PackageError::from)?;
        let bytes = object_commit.patch().after().to_vec();
        Ok(Commit {
            snapshot: Snapshot::new(bytes, self.tags),
            patch,
            package_patch: PackagePatch::new(object_commit.into_parts().1, self.protection),
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
        let package_patch = PackagePatch::new(object_commit.into_parts().1, candidate.protection);
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
            let payload = codec::to_bytes(tags, &self.source_limits)?;
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
        let reparsed = codec::parse(&reparsed_fib, &table, &self.source_limits)?;
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
