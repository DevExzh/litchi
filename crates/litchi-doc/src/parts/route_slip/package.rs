//! DOC/OLE package ownership for route-slip metadata.

#![allow(
    dead_code,
    reason = "the byte-level editor keeps its record-level API for its unit tests; the package family drives a subset"
)]

use super::codec;
use super::model::Metadata;
use super::transaction::{
    Edit as RecordEdit, Error as TransactionError, Patch, RecipientSelector,
    Snapshot as TransactionSnapshot,
};
use super::validation;
use crate::package::{Error as PackageError, Result, SourceLimits};
use crate::parts::fib::FileInformationBlock;
use crate::parts::protection::{EditProtection, PackagePatch, ProtectionPolicy, classify};
use litchi_ole_common::object::{
    Edit as ObjectEditor, Limits, OpenOptions, Package as ObjectPackage, Targets,
};

/// An immutable DOC snapshot carrying the route-slip semantic state and bytes.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Snapshot {
    bytes: Vec<u8>,
    route_slip: TransactionSnapshot,
}

impl Snapshot {
    fn new(bytes: Vec<u8>, route_slip: TransactionSnapshot) -> Self {
        Self { bytes, route_slip }
    }

    /// Returns route-slip metadata, when the package has an active range.
    #[must_use]
    pub(crate) fn metadata(&self) -> Option<&Metadata> {
        self.route_slip.metadata()
    }

    /// Returns the immutable semantic route-slip snapshot.
    #[must_use]
    pub(crate) fn route_slip(&self) -> &TransactionSnapshot {
        &self.route_slip
    }

    /// Returns the rendered package bytes.
    pub(crate) fn finish(&self) -> Result<Vec<u8>> {
        Ok(self.bytes.clone())
    }

    /// Starts a semantic route-slip transaction from this package snapshot.
    #[must_use]
    pub(crate) fn edit(&self) -> RecordEdit {
        self.route_slip.edit()
    }
}

/// A package commit containing a semantic patch and the common OLE byte patch.
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

    /// The reversible route-slip semantic patch.
    #[must_use]
    pub(crate) fn patch(&self) -> &Patch {
        &self.patch
    }

    /// The reversible whole-CFB patch produced by the common OLE editor.
    #[must_use]
    pub(crate) fn package_patch(&self) -> &PackagePatch {
        &self.package_patch
    }

    /// Splits the package commit into its snapshot, semantic patch, and byte patch.
    #[must_use]
    pub(crate) fn into_parts(self) -> (Snapshot, Patch, PackagePatch) {
        (self.snapshot, self.patch, self.package_patch)
    }
}

/// Transactional DOC editor for the `WordDocument` and selected table stream.
#[derive(Debug, Clone)]
pub(crate) struct Editor {
    package: ObjectEditor,
    table_name: String,
    source_limits: SourceLimits,
    original_route_slip: TransactionSnapshot,
    route_slip: TransactionSnapshot,
    protection: EditProtection,
    protection_policy: ProtectionPolicy,
    changed: bool,
}

impl Editor {
    /// Opens a DOC package from bytes and reads its selected route-slip range.
    ///
    /// Only Word 97 and later, unencrypted packages with a complete FIB table
    /// pointer array are accepted for editing.
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
        let route_slip =
            TransactionSnapshot::from_option(codec::parse(&fib, table, &source_limits)?)?;
        Ok(Self {
            package,
            table_name: table_name.to_owned(),
            source_limits,
            original_route_slip: route_slip.clone(),
            route_slip,
            protection,
            protection_policy,
            changed: false,
        })
    }

    /// Returns the current route-slip metadata, when present.
    #[must_use]
    pub(crate) fn metadata(&self) -> Option<&Metadata> {
        self.route_slip.metadata()
    }

    /// Returns the current immutable route-slip semantic snapshot.
    #[must_use]
    pub(crate) fn route_slip(&self) -> &TransactionSnapshot {
        &self.route_slip
    }

    /// Starts an independent semantic transaction.
    #[must_use]
    pub(crate) fn edit(&self) -> RecordEdit {
        self.route_slip.edit()
    }

    /// Whether a package edit has changed any stream bytes.
    #[must_use]
    pub(crate) fn is_changed(&self) -> bool {
        self.changed
    }

    /// Sets or replaces the complete route-slip metadata.
    pub(crate) fn set(&mut self, metadata: Metadata) -> Result<Commit> {
        let snapshot = TransactionSnapshot::new(metadata)?;
        self.install(snapshot)
    }

    /// Replaces existing route-slip metadata, rejecting an absent range.
    pub(crate) fn replace(&mut self, metadata: Metadata) -> Result<Commit> {
        if !self.route_slip.is_present() {
            return Err(PackageError::InvalidFormat(
                "cannot replace missing route-slip metadata".into(),
            ));
        }
        self.set(metadata)
    }

    /// Applies and publishes a semantic transaction atomically.
    pub(crate) fn apply(&mut self, transaction: RecordEdit) -> Result<Commit> {
        let commit = transaction.commit().map_err(transaction_error)?;
        if commit.patch().before() != &self.route_slip {
            return Err(PackageError::InvalidFormat(
                "route-slip transaction snapshot conflict".into(),
            ));
        }
        self.install(commit.package().clone())
    }

    /// Sets the current recipient stage and publishes the edit atomically.
    pub(crate) fn set_stage(&mut self, selector: RecipientSelector<'_>) -> Result<Commit> {
        let mut transaction = self.edit();
        transaction.set_stage(selector).map_err(transaction_error)?;
        self.apply(transaction)
    }

    /// Advances to the next recipient and publishes the lifecycle edit.
    pub(crate) fn advance_stage(&mut self) -> Result<Commit> {
        let mut transaction = self.edit();
        transaction.advance_stage().map_err(transaction_error)?;
        self.apply(transaction)
    }

    /// Adds a recipient and publishes the edit atomically.
    pub(crate) fn add_recipient(&mut self, recipient: super::model::Recipient) -> Result<Commit> {
        let mut transaction = self.edit();
        transaction
            .add_recipient(recipient)
            .map_err(transaction_error)?;
        self.apply(transaction)
    }

    /// Replaces a selected recipient and publishes the edit atomically.
    pub(crate) fn replace_recipient(
        &mut self,
        selector: RecipientSelector<'_>,
        recipient: super::model::Recipient,
    ) -> Result<Commit> {
        let mut transaction = self.edit();
        transaction
            .replace_recipient(selector, recipient)
            .map_err(transaction_error)?;
        self.apply(transaction)
    }

    /// Removes a selected recipient and publishes the edit atomically.
    pub(crate) fn remove_recipient(&mut self, selector: RecipientSelector<'_>) -> Result<Commit> {
        let mut transaction = self.edit();
        transaction
            .remove_recipient(selector)
            .map_err(transaction_error)?;
        self.apply(transaction)
    }

    /// Clears the route-slip range and its semantic metadata.
    pub(crate) fn clear(&mut self) -> Result<Commit> {
        let mut transaction = self.edit();
        transaction.clear().map_err(transaction_error)?;
        self.apply(transaction)
    }

    /// Completes routing by clearing the route-slip FIB range.
    pub(crate) fn complete(&mut self) -> Result<Commit> {
        let mut transaction = self.edit();
        transaction.complete().map_err(transaction_error)?;
        self.apply(transaction)
    }

    /// Captures the current package as an immutable snapshot.
    pub(crate) fn snapshot(&self) -> Result<Snapshot> {
        let bytes = crate::package::object_edit_bytes(&self.package)?;
        Ok(Snapshot::new(bytes, self.route_slip.clone()))
    }

    /// Finishes the edit and returns the rendered DOC bytes.
    pub(crate) fn finish(self) -> Result<Vec<u8>> {
        if self.changed {
            self.protection_policy.authorize(self.protection)?;
        }
        crate::package::object_edit_bytes(&self.package)
    }

    /// Commits the package as an immutable snapshot with reversible patches.
    pub(crate) fn commit(self) -> Result<Commit> {
        if self.changed {
            self.protection_policy.authorize(self.protection)?;
        }
        let semantic_patch = Patch::new(self.original_route_slip, self.route_slip.clone());
        let object_commit = self.package.commit().map_err(PackageError::from)?;
        let bytes = object_commit.patch().after().to_vec();
        let snapshot = Snapshot::new(bytes, self.route_slip);
        Ok(Commit {
            snapshot,
            patch: semantic_patch,
            package_patch: PackagePatch::new(object_commit.into_parts().1, self.protection),
        })
    }

    fn install(&mut self, snapshot: TransactionSnapshot) -> Result<Commit> {
        let semantic_patch = Patch::new(self.route_slip.clone(), snapshot.clone());
        if semantic_patch.is_noop() {
            let package_patch = self
                .package
                .clone()
                .commit()
                .map_err(PackageError::from)?
                .into_parts()
                .1;
            let package_patch = PackagePatch::new(package_patch, self.protection);
            let package_snapshot = self.snapshot()?;
            return Ok(Commit {
                snapshot: package_snapshot,
                patch: semantic_patch,
                package_patch,
            });
        }

        self.protection_policy.authorize(self.protection)?;

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
        let package_snapshot = Snapshot::new(bytes, candidate.route_slip.clone());
        let package_patch = PackagePatch::new(package_commit.into_parts().1, candidate.protection);
        *self = candidate;
        Ok(Commit {
            snapshot: package_snapshot,
            patch: semantic_patch,
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
        let pointer = validation::route_pointer_location(&fib)?;
        if let Some(metadata) = snapshot.metadata() {
            let payload = codec::to_bytes(metadata, &self.source_limits)?;
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
        let reparsed = codec::parse(&reparsed_fib, &table, &self.source_limits)?;
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
