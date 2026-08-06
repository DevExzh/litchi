//! DOC/OLE package ownership for route-slip metadata.

use super::codec;
use super::model::Metadata;
use super::transaction::{Patch, RecipientSelector, Snapshot, Transaction, TransactionError};
use super::validation;
use crate::package::{Error as PackageError, Result};
use crate::parts::fib::FileInformationBlock;
use litchi_ole_common::object::{Editor as ObjectEditor, Limits, Patch as ObjectPatch, Targets};

/// An immutable DOC snapshot carrying the route-slip semantic state and bytes.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PackageSnapshot {
    bytes: Vec<u8>,
    route_slip: Snapshot,
}

impl PackageSnapshot {
    fn new(bytes: Vec<u8>, route_slip: Snapshot) -> Self {
        Self { bytes, route_slip }
    }

    /// Returns route-slip metadata, when the package has an active range.
    #[must_use]
    pub fn metadata(&self) -> Option<&Metadata> {
        self.route_slip.metadata()
    }

    /// Returns the immutable semantic route-slip snapshot.
    #[must_use]
    pub fn route_slip(&self) -> &Snapshot {
        &self.route_slip
    }

    /// Returns the rendered package bytes.
    pub fn finish(&self) -> Result<Vec<u8>> {
        Ok(self.bytes.clone())
    }

    /// Starts a semantic route-slip transaction from this package snapshot.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        self.route_slip.edit()
    }
}

/// A package commit containing a semantic patch and the common OLE byte patch.
#[derive(Debug, Clone)]
pub struct PackageCommit {
    snapshot: PackageSnapshot,
    patch: Patch,
    package_patch: ObjectPatch,
}

impl PackageCommit {
    /// The immutable post-edit package snapshot.
    #[must_use]
    pub fn snapshot(&self) -> &PackageSnapshot {
        &self.snapshot
    }

    /// The reversible route-slip semantic patch.
    #[must_use]
    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    /// The reversible whole-CFB patch produced by the common OLE editor.
    #[must_use]
    pub fn package_patch(&self) -> &ObjectPatch {
        &self.package_patch
    }

    /// Splits the package commit into its snapshot, semantic patch, and byte patch.
    #[must_use]
    pub fn into_parts(self) -> (PackageSnapshot, Patch, ObjectPatch) {
        (self.snapshot, self.patch, self.package_patch)
    }
}

/// Transactional DOC editor for the WordDocument and selected table stream.
#[derive(Debug, Clone)]
pub struct Editor {
    package: ObjectEditor,
    table_name: String,
    original_route_slip: Snapshot,
    route_slip: Snapshot,
    changed: bool,
}

impl Editor {
    /// Opens a DOC package from bytes and reads its selected route-slip range.
    ///
    /// Only Word 97 and later, unencrypted packages with a complete FIB table
    /// pointer array are accepted for editing.
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
        let route_slip =
            Snapshot::from_option(codec::parse(&fib, table)?).map_err(PackageError::from)?;
        Ok(Self {
            package,
            table_name: table_name.to_owned(),
            original_route_slip: route_slip.clone(),
            route_slip,
            changed: false,
        })
    }

    /// Returns the current route-slip metadata, when present.
    #[must_use]
    pub fn metadata(&self) -> Option<&Metadata> {
        self.route_slip.metadata()
    }

    /// Returns the current immutable route-slip semantic snapshot.
    #[must_use]
    pub fn route_slip(&self) -> &Snapshot {
        &self.route_slip
    }

    /// Starts an independent semantic transaction.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        self.route_slip.edit()
    }

    /// Whether a package edit has changed any stream bytes.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.changed
    }

    /// Sets or replaces the complete route-slip metadata.
    pub fn set(&mut self, metadata: Metadata) -> Result<PackageCommit> {
        let snapshot = Snapshot::new(metadata)?;
        self.install(snapshot)
    }

    /// Replaces existing route-slip metadata, rejecting an absent range.
    pub fn replace(&mut self, metadata: Metadata) -> Result<PackageCommit> {
        if !self.route_slip.is_present() {
            return Err(PackageError::InvalidFormat(
                "cannot replace missing route-slip metadata".into(),
            ));
        }
        self.set(metadata)
    }

    /// Applies and publishes a semantic transaction atomically.
    pub fn apply(&mut self, transaction: Transaction) -> Result<PackageCommit> {
        let commit = transaction.commit().map_err(transaction_error)?;
        if commit.patch().before() != &self.route_slip {
            return Err(PackageError::InvalidFormat(
                "route-slip transaction snapshot conflict".into(),
            ));
        }
        self.install(commit.snapshot().clone())
    }

    /// Sets the current recipient stage and publishes the edit atomically.
    pub fn set_stage(&mut self, selector: RecipientSelector<'_>) -> Result<PackageCommit> {
        let mut transaction = self.edit();
        transaction.set_stage(selector).map_err(transaction_error)?;
        self.apply(transaction)
    }

    /// Advances to the next recipient and publishes the lifecycle edit.
    pub fn advance_stage(&mut self) -> Result<PackageCommit> {
        let mut transaction = self.edit();
        transaction.advance_stage().map_err(transaction_error)?;
        self.apply(transaction)
    }

    /// Adds a recipient and publishes the edit atomically.
    pub fn add_recipient(&mut self, recipient: super::model::Recipient) -> Result<PackageCommit> {
        let mut transaction = self.edit();
        transaction
            .add_recipient(recipient)
            .map_err(transaction_error)?;
        self.apply(transaction)
    }

    /// Replaces a selected recipient and publishes the edit atomically.
    pub fn replace_recipient(
        &mut self,
        selector: RecipientSelector<'_>,
        recipient: super::model::Recipient,
    ) -> Result<PackageCommit> {
        let mut transaction = self.edit();
        transaction
            .replace_recipient(selector, recipient)
            .map_err(transaction_error)?;
        self.apply(transaction)
    }

    /// Removes a selected recipient and publishes the edit atomically.
    pub fn remove_recipient(&mut self, selector: RecipientSelector<'_>) -> Result<PackageCommit> {
        let mut transaction = self.edit();
        transaction
            .remove_recipient(selector)
            .map_err(transaction_error)?;
        self.apply(transaction)
    }

    /// Clears the route-slip range and its semantic metadata.
    pub fn clear(&mut self) -> Result<PackageCommit> {
        let mut transaction = self.edit();
        transaction.clear().map_err(transaction_error)?;
        self.apply(transaction)
    }

    /// Completes routing by clearing the route-slip FIB range.
    pub fn complete(&mut self) -> Result<PackageCommit> {
        let mut transaction = self.edit();
        transaction.complete().map_err(transaction_error)?;
        self.apply(transaction)
    }

    /// Captures the current package as an immutable snapshot.
    pub fn snapshot(&self) -> Result<PackageSnapshot> {
        let bytes = self.package.clone().finish().map_err(PackageError::from)?;
        Ok(PackageSnapshot::new(bytes, self.route_slip.clone()))
    }

    /// Finishes the edit and returns the rendered DOC bytes.
    pub fn finish(self) -> Result<Vec<u8>> {
        self.package.finish().map_err(PackageError::from)
    }

    /// Commits the package as an immutable snapshot with reversible patches.
    pub fn commit(self) -> Result<PackageCommit> {
        let semantic_patch = Patch::new(self.original_route_slip, self.route_slip.clone());
        let object_commit = self.package.commit().map_err(PackageError::from)?;
        let bytes = object_commit.patch().after().to_vec();
        let snapshot = PackageSnapshot::new(bytes, self.route_slip);
        Ok(PackageCommit {
            snapshot,
            patch: semantic_patch,
            package_patch: object_commit.into_patch(),
        })
    }

    fn install(&mut self, snapshot: Snapshot) -> Result<PackageCommit> {
        let semantic_patch = Patch::new(self.route_slip.clone(), snapshot.clone());
        if semantic_patch.is_noop() {
            let package_patch = self
                .package
                .clone()
                .commit()
                .map_err(PackageError::from)?
                .into_patch();
            let package_snapshot = self.snapshot()?;
            return Ok(PackageCommit {
                snapshot: package_snapshot,
                patch: semantic_patch,
                package_patch,
            });
        }

        let mut candidate = self.clone();
        candidate.write_snapshot(&snapshot)?;
        candidate.route_slip = snapshot;
        candidate.changed = true;

        let package_commit = candidate
            .package
            .clone()
            .commit()
            .map_err(PackageError::from)?;
        let bytes = package_commit.patch().after().to_vec();
        let package_snapshot = PackageSnapshot::new(bytes, candidate.route_slip.clone());
        let package_patch = package_commit.into_patch();
        *self = candidate;
        Ok(PackageCommit {
            snapshot: package_snapshot,
            patch: semantic_patch,
            package_patch,
        })
    }

    fn write_snapshot(&mut self, snapshot: &Snapshot) -> Result<()> {
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
        let pointer = validation::route_pointer_location(&fib)?;
        if let Some(metadata) = snapshot.metadata() {
            let payload = codec::to_bytes(metadata)?;
            let offset = u32::try_from(table.len())
                .map_err(|_| PackageError::Corrupted("table stream exceeds u32::MAX".into()))?;
            let length = u32::try_from(payload.len()).map_err(|_| {
                PackageError::Corrupted("route-slip payload exceeds u32::MAX".into())
            })?;
            table.extend_from_slice(&payload);
            word[pointer..pointer + 4].copy_from_slice(&offset.to_le_bytes());
            word[pointer + 4..pointer + 8].copy_from_slice(&length.to_le_bytes());
        } else {
            word[pointer..pointer + 8].fill(0);
        }

        let reparsed_fib = FileInformationBlock::parse(&word)?;
        let reparsed = codec::parse(&reparsed_fib, &table)?;
        if reparsed.as_ref() != snapshot.metadata() {
            return Err(PackageError::Corrupted(
                "route-slip snapshot failed FIB/table-stream round-trip validation".into(),
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
    PackageError::InvalidFormat(format!("route-slip transaction failed: {error}"))
}
