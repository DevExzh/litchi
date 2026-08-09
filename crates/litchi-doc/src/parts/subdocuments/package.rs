//! DOC/OLE package ownership for `PlcfWKB` and `SttbFnm`.
//!
//! The detached owner edits only validated table images. This layer owns the
//! `WordDocument` FIB and the selected `0Table`/`1Table` stream, appending new
//! table images before atomically publishing both FIB pointer pairs. The old
//! table images and every unrelated stream remain intact. The common OLE
//! editor supplies exact no-op bytes, source checks, CLSID preservation, and
//! candidate reparse validation.

use super::model::Collection;
use super::transaction::{
    Commit as TransactionCommit, Error as TransactionError, Patch, Snapshot as TransactionSnapshot,
    Transaction,
};
use super::validation;
use crate::package::{Error as PackageError, Result};
use crate::parts::fib::FileInformationBlock;
use litchi_ole_common::object::{Editor as ObjectEditor, Limits, Patch as ObjectPatch, Targets};

/// Immutable package bytes plus the contextual subdocument snapshot.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Snapshot {
    bytes: Vec<u8>,
    subdocuments: TransactionSnapshot,
}

impl Snapshot {
    fn new(bytes: Vec<u8>, subdocuments: TransactionSnapshot) -> Self {
        Self {
            bytes,
            subdocuments,
        }
    }

    /// The immutable semantic and detached-wire snapshot.
    #[must_use]
    pub fn subdocuments(&self) -> &TransactionSnapshot {
        &self.subdocuments
    }

    /// The decoded master-document subdocument metadata.
    #[must_use]
    pub fn value(&self) -> &Collection {
        self.subdocuments.collection()
    }

    /// Return the rendered package bytes.
    pub fn finish(&self) -> Result<Vec<u8>> {
        Ok(self.bytes.clone())
    }

    /// Start an independent semantic transaction from this package snapshot.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        self.subdocuments.edit()
    }
}

/// A package commit containing a semantic patch and a reversible whole-CFB
/// patch. The semantic patch remains detached; [`Self::package_patch`] is the
/// publication boundary that includes FIB pointer relocation.
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

    /// The reversible semantic and detached table patch.
    #[must_use]
    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    /// The source-checked whole-CFB patch, including `WordDocument` FIB edits.
    #[must_use]
    pub fn package_patch(&self) -> &ObjectPatch {
        &self.package_patch
    }

    /// Split the package snapshot and both reversible patches.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch, ObjectPatch) {
        (self.snapshot, self.patch, self.package_patch)
    }
}

/// Transactional owner for the `WordDocument` and selected table stream.
#[derive(Debug, Clone)]
pub struct Editor {
    package: ObjectEditor,
    table_name: String,
    original: TransactionSnapshot,
    /// The semantic/detached wire lineage. The physical package may contain
    /// older table images because publication is append-only, so this value
    /// intentionally remains the detached candidate lineage.
    subdocuments: TransactionSnapshot,
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
            .ok_or_else(|| PackageError::StreamNotFound("WordDocument".to_owned()))?;
        let fib = FileInformationBlock::parse(word)?;
        let table_name = if fib.which_table_stream() {
            "1Table"
        } else {
            "0Table"
        };
        let table_path = vec![table_name.to_owned()];
        let table = package
            .stream(&table_path)
            .ok_or_else(|| PackageError::StreamNotFound(table_name.to_owned()))?;
        validation::package_fib(&fib, table)?;
        let subdocuments = TransactionSnapshot::parse(&fib, table)?.ok_or_else(|| {
            PackageError::InvalidFormat(
                "DOC package has no FIB-addressed PlcfWKB or SttbFnm owner".to_owned(),
            )
        })?;
        Ok(Self {
            package,
            table_name: table_name.to_owned(),
            original: subdocuments.clone(),
            subdocuments,
            changed: false,
        })
    }

    /// The current contextual subdocument snapshot.
    #[must_use]
    pub fn subdocuments(&self) -> &TransactionSnapshot {
        &self.subdocuments
    }

    /// The decoded master-document subdocument metadata.
    #[must_use]
    pub fn value(&self) -> &Collection {
        self.subdocuments.collection()
    }

    /// Whether this editor has published changed package bytes.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.changed
    }

    /// Start an independent semantic transaction.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        self.subdocuments.edit()
    }

    /// Apply and publish a validated semantic transaction atomically.
    pub fn apply(&mut self, transaction: Transaction) -> Result<Commit> {
        let commit = transaction.commit().map_err(transaction_error)?;
        if commit.patch().before() != &self.subdocuments {
            return Err(PackageError::InvalidFormat(
                "subdocument transaction snapshot conflict".to_owned(),
            ));
        }
        self.install(commit)
    }

    /// Capture the current package as an immutable snapshot.
    pub fn snapshot(&self) -> Result<Snapshot> {
        let bytes = self.package.clone().finish().map_err(PackageError::from)?;
        Ok(Snapshot::new(bytes, self.subdocuments.clone()))
    }

    /// Finish the edit and return rendered DOC bytes.
    pub fn finish(self) -> Result<Vec<u8>> {
        self.package.finish().map_err(PackageError::from)
    }

    /// Commit the current package as a semantic snapshot plus reversible
    /// semantic and whole-CFB patches.
    pub fn commit(self) -> Result<Commit> {
        let patch = Patch::new(self.original, self.subdocuments.clone());
        let object_commit = self.package.commit().map_err(PackageError::from)?;
        let bytes = object_commit.patch().after().to_vec();
        Ok(Commit {
            snapshot: Snapshot::new(bytes, self.subdocuments),
            patch,
            package_patch: object_commit.into_patch(),
        })
    }

    fn install(&mut self, commit: TransactionCommit) -> Result<Commit> {
        let (snapshot, patch) = commit.into_parts();
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

        // Candidate rendering and reparsing happen on a clone. The source
        // editor is replaced only after both FIB pointers and both table
        // payloads have passed validation.
        let mut candidate = self.clone();
        candidate.write_snapshot(&snapshot)?;
        candidate.subdocuments = snapshot;
        candidate.changed = true;
        let object_commit = candidate
            .package
            .clone()
            .commit()
            .map_err(PackageError::from)?;
        let bytes = object_commit.patch().after().to_vec();
        let package_patch = object_commit.into_patch();
        let package_snapshot = Snapshot::new(bytes, candidate.subdocuments.clone());
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
            .ok_or_else(|| PackageError::StreamNotFound("WordDocument".to_owned()))?
            .to_vec();
        let mut table = self
            .package
            .stream(&table_path)
            .ok_or_else(|| PackageError::StreamNotFound(self.table_name.clone()))?
            .to_vec();
        let fib = FileInformationBlock::parse(&word)?;
        validation::package_fib(&fib, &table)?;
        let current = TransactionSnapshot::parse(&fib, &table)?.ok_or_else(|| {
            PackageError::Corrupted(
                "the current DOC package lost its subdocument table owner".to_owned(),
            )
        })?;
        if !same_source(&current, &self.subdocuments) {
            return Err(PackageError::InvalidFormat(
                "subdocument package source changed before publication".to_owned(),
            ));
        }

        let referenced_files = if snapshot.collection().referenced_files()
            == self.subdocuments.collection().referenced_files()
        {
            current.referenced_files_bytes().map(ToOwned::to_owned)
        } else {
            snapshot.encode_referenced_files()?
        };
        let subdocuments = if snapshot.collection().subdocuments()
            == self.subdocuments.collection().subdocuments()
            && snapshot.main_document_chars() == self.subdocuments.main_document_chars()
        {
            current.subdocuments_bytes().map(ToOwned::to_owned)
        } else {
            snapshot.encode_subdocuments()?
        };
        if let Some(payload) = referenced_files
            && current.referenced_files_bytes() != Some(payload.as_slice())
        {
            let (offset, length) = append_table(&mut table, &payload, "SttbFnm")?;
            set_pointer(&mut word, &fib, validation::STTB_FNM, offset, length)?;
        }
        if let Some(payload) = subdocuments
            && current.subdocuments_bytes() != Some(payload.as_slice())
        {
            let (offset, length) = append_table(&mut table, &payload, "PlcfWKB")?;
            set_pointer(&mut word, &fib, validation::PLCF_WKB, offset, length)?;
        }

        let reparsed_fib = FileInformationBlock::parse(&word)?;
        validation::package_fib(&reparsed_fib, &table)?;
        let reparsed = TransactionSnapshot::parse(&reparsed_fib, &table)?.ok_or_else(|| {
            PackageError::Corrupted(
                "published DOC package lost its subdocument table owner".to_owned(),
            )
        })?;
        if !same_source(&reparsed, snapshot) {
            return Err(PackageError::Corrupted(
                "subdocument snapshot failed FIB/table publication validation".to_owned(),
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

fn same_source(left: &TransactionSnapshot, right: &TransactionSnapshot) -> bool {
    left.collection() == right.collection()
        && left.main_document_chars() == right.main_document_chars()
        && left.referenced_files_bytes() == right.referenced_files_bytes()
        && left.subdocuments_bytes() == right.subdocuments_bytes()
}

fn append_table(table: &mut Vec<u8>, payload: &[u8], name: &str) -> Result<(u32, u32)> {
    let offset = u32::try_from(table.len())
        .map_err(|_| PackageError::Corrupted(format!("{name} offset exceeds u32::MAX")))?;
    let length = u32::try_from(payload.len())
        .map_err(|_| PackageError::Corrupted(format!("{name} length exceeds u32::MAX")))?;
    let end = table
        .len()
        .checked_add(payload.len())
        .ok_or_else(|| PackageError::Corrupted(format!("{name} range overflows")))?;
    if end > u32::MAX as usize {
        return Err(PackageError::Corrupted(format!(
            "table stream exceeds u32::MAX after appending {name}"
        )));
    }
    table.extend_from_slice(payload);
    Ok((offset, length))
}

fn set_pointer(
    word: &mut [u8],
    fib: &FileInformationBlock,
    index: usize,
    offset: u32,
    length: u32,
) -> Result<()> {
    let pointer = validation::pointer_location(fib, index)?;
    let bytes = word
        .get_mut(pointer..pointer + 8)
        .ok_or_else(|| PackageError::Corrupted("FIB pointer exceeds WordDocument".to_owned()))?;
    bytes[..4].copy_from_slice(&offset.to_le_bytes());
    bytes[4..].copy_from_slice(&length.to_le_bytes());
    Ok(())
}

fn transaction_error(error: TransactionError) -> PackageError {
    PackageError::InvalidFormat(format!("subdocument transaction failed: {error}"))
}
