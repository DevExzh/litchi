//! Immutable snapshots and failure-atomic semantic edits for master-document
//! subdocument metadata.

use super::codec::{self, PLCF_WKB, STTB_FNM, WKB_FLAGS_REQUIRED, WKB_OUTLINE_LEVEL};
use super::model::{
    Collection, FileNameKey, FileNameKeyError, FileNameMetadata, Kind, Name, Reference,
};
use super::patch::{PatchError, SourceContext, SourceRanges, TablePatch, WireImage};
use super::validation;
use crate::package::{Error as PackageError, Result as PackageResult};
use crate::parts::fib::FileInformationBlock;
use std::fmt;

/// An immutable, source-capturing subdocument metadata snapshot.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Snapshot {
    pub(super) collection: Collection,
    pub(super) main_document_chars: u32,
    pub(super) wire: WireImage,
}

impl Snapshot {
    /// Parse and capture the exact FIB-addressed `SttbFnm`/`PlcfWKB` slices.
    ///
    /// `None` means neither owner table is present. A present snapshot owns
    /// the table bytes and the relevant main-document context; it never holds
    /// a live reference to the caller's FIB or table stream.
    pub fn parse(fib: &FileInformationBlock, table_stream: &[u8]) -> PackageResult<Option<Self>> {
        let Some(collection) = Collection::parse(fib, table_stream)? else {
            return Ok(None);
        };
        let referenced_files = codec::table_slice(fib, table_stream, STTB_FNM, "SttbFnm")?;
        let subdocuments = codec::table_slice(fib, table_stream, PLCF_WKB, "PlcfWKB")?;
        let ranges = SourceRanges {
            referenced_files: referenced_files
                .map(|(offset, data)| super::patch::TableRange::new(offset, data.len())),
            subdocuments: subdocuments
                .map(|(offset, data)| super::patch::TableRange::new(offset, data.len())),
        };
        let main_document_chars = fib.get_main_doc_range().1;
        validation::collection(&collection, main_document_chars)?;
        let wire = WireImage::capture(
            table_stream,
            main_document_chars,
            ranges,
            referenced_files.map(|(_, data)| data),
            subdocuments.map(|(_, data)| data),
        )
        .map_err(patch_error)?;
        Ok(Some(Self {
            collection,
            main_document_chars,
            wire,
        }))
    }

    /// The typed immutable collection.
    #[must_use]
    pub fn collection(&self) -> &Collection {
        &self.collection
    }

    /// The main-document `ccpText` captured with this snapshot.
    pub const fn main_document_chars(&self) -> u32 {
        self.main_document_chars
    }

    /// The source stream/range context required by the bounded patch.
    pub const fn source_context(&self) -> SourceContext {
        self.wire.context
    }

    /// The exact source `SttbFnm` bytes, when present.
    pub fn referenced_files_bytes(&self) -> Option<&[u8]> {
        self.wire.referenced_files.as_deref()
    }

    /// The exact source `PlcfWKB` bytes, when present.
    pub fn subdocuments_bytes(&self) -> Option<&[u8]> {
        self.wire.subdocuments.as_deref()
    }

    /// Encode the candidate's complete `SttbFnm` independently, with its
    /// exact wire length.
    pub fn encode_referenced_files(&self) -> PackageResult<Option<Vec<u8>>> {
        validation::collection(&self.collection, self.main_document_chars)?;
        if self.wire.referenced_files.is_some() {
            Ok(Some(codec::encode_sttb_fnm(
                &self.collection.referenced_files,
            )?))
        } else if self.collection.referenced_files.is_empty() {
            Ok(None)
        } else {
            Err(PackageError::Corrupted(
                "SttbFnm is absent, so a referenced file cannot be encoded".to_string(),
            ))
        }
    }

    /// Encode the candidate's complete `PlcfWKB` independently, with its
    /// exact wire length and terminal CP.
    pub fn encode_subdocuments(&self) -> PackageResult<Option<Vec<u8>>> {
        validation::collection(&self.collection, self.main_document_chars)?;
        if self.wire.subdocuments.is_some() {
            Ok(Some(codec::encode_plcf_wkb(
                &self.collection.subdocuments,
                self.main_document_chars,
                &self.collection.referenced_files,
            )?))
        } else if self.collection.subdocuments.is_empty() {
            Ok(None)
        } else {
            Err(PackageError::Corrupted(
                "PlcfWKB is absent, so a subdocument reference cannot be encoded".to_string(),
            ))
        }
    }

    /// Starts a failure-atomic semantic transaction.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        Transaction {
            before: self.clone(),
            working: self.clone(),
        }
    }
}

/// A checked selection of one `SttbFnm` entry.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FileNameSelector {
    /// Select by zero-based `SttbFnm` order.
    Index(usize),
    /// Select by its typed `FNPI` key.
    Key(FileNameKey),
}

/// A checked selection of one `PlcfWKB` reference by its ordered index.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ReferenceSelector(pub usize);

impl ReferenceSelector {
    /// Select a zero-based `PlcfWKB` reference index.
    pub const fn index(index: usize) -> Self {
        Self(index)
    }

    pub const fn get(self) -> usize {
        self.0
    }
}

/// A staged candidate based on one immutable source snapshot.
#[derive(Debug, Clone)]
pub struct Transaction {
    before: Snapshot,
    working: Snapshot,
}

impl Transaction {
    /// Creates a transaction from an immutable snapshot.
    #[must_use]
    pub fn new(snapshot: Snapshot) -> Self {
        Self {
            before: snapshot.clone(),
            working: snapshot,
        }
    }

    /// The current candidate snapshot.
    #[must_use]
    pub fn snapshot(&self) -> &Snapshot {
        &self.working
    }

    /// The immutable source snapshot.
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Whether the candidate differs semantically or in its terminal-CP
    /// context from the source.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.before.collection != self.working.collection
            || self.before.main_document_chars != self.working.main_document_chars
    }

    /// Restores the candidate to its source snapshot.
    pub fn rollback(&mut self) {
        self.working = self.before.clone();
    }

    /// Changes the `ccpText` context used for the WKB terminal CP.
    ///
    /// This does not edit main-document text. The caller must update the
    /// document's FIB/text owner separately when applying the resulting table
    /// patch.
    pub fn set_main_document_chars(&mut self, value: u32) {
        self.working.main_document_chars = value;
    }

    /// Adds one external file and allocates the smallest unused identifier for
    /// its `FNPI.fnpt` kind.
    pub fn add_file_name(
        &mut self,
        kind: Kind,
        path: impl Into<String>,
        metadata: FileNameMetadata,
    ) -> Result<FileNameKey, TransactionError> {
        let mut candidate = self.working.collection.clone();
        let key = allocate_key(&candidate, kind)?;
        candidate.referenced_files.push(Name {
            fnpi: key.fnpi(),
            raw_fnfb: raw_fnfb(metadata),
            fnif_unused: [0; 4],
            path: path.into(),
            relative_path_offset: metadata.relative_path_offset,
            valid_on_fat: metadata.valid_on_fat,
            valid_on_ntfs: metadata.valid_on_ntfs,
            is_non_file_system_path: metadata.is_non_file_system_path,
        });
        self.publish(candidate)?;
        Ok(key)
    }

    /// Adds a subdocument file and an ordered WKB reference as one atomic
    /// operation. A source without a `PlcfWKB` slice cannot safely grow one,
    /// so this operation rejects that case rather than pretending to update
    /// FIB pointers.
    pub fn add_subdocument(
        &mut self,
        start: u32,
        path: impl Into<String>,
        metadata: FileNameMetadata,
    ) -> Result<FileNameKey, TransactionError> {
        if self.working.wire.context.ranges().subdocuments().is_none() {
            return Err(TransactionError::Unsupported(
                "the source has no PlcfWKB slice; adding one requires a FIB/package writer",
            ));
        }
        let mut candidate = self.working.collection.clone();
        let key = allocate_key(&candidate, Kind::Subdocument)?;
        let file_name_index = candidate.referenced_files.len();
        candidate.referenced_files.push(Name {
            fnpi: key.fnpi(),
            raw_fnfb: raw_fnfb(metadata),
            fnif_unused: [0; 4],
            path: path.into(),
            relative_path_offset: metadata.relative_path_offset,
            valid_on_fat: metadata.valid_on_fat,
            valid_on_ntfs: metadata.valid_on_ntfs,
            is_non_file_system_path: metadata.is_non_file_system_path,
        });
        candidate.subdocuments.push(Reference {
            start,
            outline_level: WKB_OUTLINE_LEVEL,
            file_name: key.fnpi(),
            file_name_index,
            raw_flags: WKB_FLAGS_REQUIRED,
            raw_wkb: [0; 12],
        });
        candidate
            .subdocuments
            .sort_by_key(|reference| reference.start);
        self.publish(candidate)?;
        Ok(key)
    }

    /// Replaces a path and its typed `FNIF` metadata while preserving unknown
    /// source `fnfb` bits and `unused` bytes.
    pub fn update_file_name(
        &mut self,
        selector: FileNameSelector,
        path: impl Into<String>,
        metadata: FileNameMetadata,
    ) -> Result<(), TransactionError> {
        let index = self.resolve_file_name(selector)?;
        let mut candidate = self.working.collection.clone();
        let file_count = candidate.referenced_files.len();
        let file = candidate
            .referenced_files
            .get_mut(index)
            .ok_or(TransactionError::Selection(SelectionError::FileNameIndex {
                index,
                len: file_count,
            }))?;
        file.path = path.into();
        file.relative_path_offset = metadata.relative_path_offset;
        file.valid_on_fat = metadata.valid_on_fat;
        file.valid_on_ntfs = metadata.valid_on_ntfs;
        file.is_non_file_system_path = metadata.is_non_file_system_path;
        self.publish(candidate)
    }

    /// Replaces only a file path, retaining its current `FNIF` metadata.
    pub fn update_file_path(
        &mut self,
        selector: FileNameSelector,
        path: impl Into<String>,
    ) -> Result<(), TransactionError> {
        let index = self.resolve_file_name(selector)?;
        let metadata = self.working.collection.referenced_files[index].metadata();
        self.update_file_name(selector, path, metadata)
    }

    /// Allocates a new identifier within the selected file's kind and updates
    /// every WKB reference to that file atomically.
    pub fn renumber_file_name(
        &mut self,
        selector: FileNameSelector,
        identifier: u16,
    ) -> Result<FileNameKey, TransactionError> {
        let index = self.resolve_file_name(selector)?;
        let mut candidate = self.working.collection.clone();
        let old_key = candidate.referenced_files[index].key();
        let key = FileNameKey::try_new(old_key.kind(), identifier).map_err(key_error)?;
        if candidate
            .referenced_files
            .iter()
            .enumerate()
            .any(|(other, file)| other != index && file.key() == key)
        {
            return Err(TransactionError::Invalid(PackageError::Corrupted(
                "the requested FNPI key is already allocated".to_string(),
            )));
        }
        candidate.referenced_files[index].fnpi = key.fnpi();
        for reference in &mut candidate.subdocuments {
            if reference.file_name == old_key.fnpi() {
                reference.file_name = key.fnpi();
            }
        }
        self.publish(candidate)?;
        Ok(key)
    }

    /// Moves one subdocument start CP while retaining WKB order.
    pub fn set_subdocument_start(
        &mut self,
        selector: ReferenceSelector,
        start: u32,
    ) -> Result<(), TransactionError> {
        let index = self.resolve_reference(selector)?;
        let mut candidate = self.working.collection.clone();
        candidate.subdocuments[index].start = start;
        candidate
            .subdocuments
            .sort_by_key(|reference| reference.start);
        self.publish(candidate)
    }

    /// Repoints one WKB entry to an existing subdocument file name.
    pub fn set_subdocument_file(
        &mut self,
        selector: ReferenceSelector,
        file: FileNameSelector,
    ) -> Result<(), TransactionError> {
        let file_index = self.resolve_file_name(file)?;
        let candidate_file = self.working.collection.referenced_files[file_index].clone();
        if candidate_file.kind() != Kind::Subdocument {
            return Err(TransactionError::Invalid(PackageError::Corrupted(
                "WKB references must use a subdocument FNPI".to_string(),
            )));
        }
        let index = self.resolve_reference(selector)?;
        let mut candidate = self.working.collection.clone();
        candidate.subdocuments[index].file_name = candidate_file.fnpi;
        candidate.subdocuments[index].file_name_index = file_index;
        self.publish(candidate)
    }

    /// Removes one WKB entry. Referenced file names remain in `SttbFnm` so
    /// unrelated mail-merge or future owner references are not guessed at.
    pub fn remove_subdocument(
        &mut self,
        selector: ReferenceSelector,
    ) -> Result<(), TransactionError> {
        let index = self.resolve_reference(selector)?;
        let mut candidate = self.working.collection.clone();
        candidate.subdocuments.remove(index);
        self.publish(candidate)
    }

    /// Commits the candidate and creates both a semantic patch and a bounded
    /// exact table-stream patch.
    pub fn commit(self) -> Result<Commit, TransactionError> {
        validation::collection(&self.working.collection, self.working.main_document_chars)
            .map_err(TransactionError::Invalid)?;
        let after_wire = WireImage::reencode(
            &self.before.wire,
            &self.working.collection,
            self.working.main_document_chars,
        )
        .map_err(TransactionError::Invalid)?;
        let after = Snapshot {
            collection: self.working.collection,
            main_document_chars: self.working.main_document_chars,
            wire: after_wire.clone(),
        };
        let table_patch = TablePatch::new(self.before.wire.clone(), after_wire);
        let patch = Patch {
            before: self.before,
            after: after.clone(),
            table: table_patch,
        };
        Ok(Commit {
            snapshot: after,
            patch,
        })
    }

    fn publish(&mut self, collection: Collection) -> Result<(), TransactionError> {
        validation::collection(&collection, self.working.main_document_chars)
            .map_err(TransactionError::Invalid)?;
        self.working.collection = collection;
        Ok(())
    }

    fn resolve_file_name(&self, selector: FileNameSelector) -> Result<usize, TransactionError> {
        match selector {
            FileNameSelector::Index(index) => {
                if index < self.working.collection.referenced_files.len() {
                    Ok(index)
                } else {
                    Err(TransactionError::Selection(SelectionError::FileNameIndex {
                        index,
                        len: self.working.collection.referenced_files.len(),
                    }))
                }
            },
            FileNameSelector::Key(key) => self
                .working
                .collection
                .referenced_files
                .iter()
                .position(|file| file.key() == key)
                .ok_or(TransactionError::Selection(SelectionError::FileNameKey(
                    key,
                ))),
        }
    }

    fn resolve_reference(&self, selector: ReferenceSelector) -> Result<usize, TransactionError> {
        let index = selector.get();
        if index < self.working.collection.subdocuments.len() {
            Ok(index)
        } else {
            Err(TransactionError::Selection(
                SelectionError::ReferenceIndex {
                    index,
                    len: self.working.collection.subdocuments.len(),
                },
            ))
        }
    }
}

/// A reversible semantic and wire-level subdocument change.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
    table: TablePatch,
}

impl Patch {
    pub(crate) fn new(before: Snapshot, after: Snapshot) -> Self {
        Self {
            table: TablePatch::new(before.wire.clone(), after.wire.clone()),
            before,
            after,
        }
    }

    /// The exact semantic source snapshot.
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// The exact semantic result snapshot.
    #[must_use]
    pub fn after(&self) -> &Snapshot {
        &self.after
    }

    /// The source-checked table-stream replacement.
    #[must_use]
    pub fn table_patch(&self) -> &TablePatch {
        &self.table
    }

    /// Whether semantic and wire state are unchanged.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after && self.table.is_noop()
    }

    /// Applies this semantic patch only to the exact source snapshot.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot, TransactionError> {
        if source != &self.before {
            return Err(TransactionError::Conflict);
        }
        Ok(self.after.clone())
    }

    /// Returns the exact inverse semantic and table-stream patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
            table: self.table.inverse(),
        }
    }
}

/// The result of a committed transaction.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    /// The immutable candidate snapshot.
    #[must_use]
    pub fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// The reversible semantic/wire patch.
    #[must_use]
    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Splits the committed result into its snapshot and patch.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

/// A selector failed to resolve against the candidate.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum SelectionError {
    /// An index was outside `SttbFnm` order.
    FileNameIndex { index: usize, len: usize },
    /// A key was not present in `SttbFnm`.
    FileNameKey(FileNameKey),
    /// An index was outside `PlcfWKB` order.
    ReferenceIndex { index: usize, len: usize },
}

impl fmt::Display for SelectionError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::FileNameIndex { index, len } => {
                write!(
                    formatter,
                    "file-name index {index} is outside {len} entries"
                )
            },
            Self::FileNameKey(key) => write!(formatter, "file-name key {key:?} was not found"),
            Self::ReferenceIndex { index, len } => {
                write!(
                    formatter,
                    "subdocument index {index} is outside {len} entries"
                )
            },
        }
    }
}

impl std::error::Error for SelectionError {}

/// Errors produced while staging or applying a semantic subdocument edit.
#[derive(Debug)]
pub enum TransactionError {
    /// The candidate violates an MS-DOC invariant or wire limit.
    Invalid(PackageError),
    /// A selector could not be resolved.
    Selection(SelectionError),
    /// The patch was applied to a different semantic snapshot.
    Conflict,
    /// The requested operation would require a package/FIB writer outside
    /// this bounded owner.
    Unsupported(&'static str),
    /// A bounded table patch could not be applied.
    Patch(PatchError),
}

impl fmt::Display for TransactionError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Invalid(error) => error.fmt(formatter),
            Self::Selection(error) => error.fmt(formatter),
            Self::Conflict => formatter.write_str("subdocument transaction snapshot conflict"),
            Self::Unsupported(message) => formatter.write_str(message),
            Self::Patch(error) => error.fmt(formatter),
        }
    }
}

impl std::error::Error for TransactionError {}

impl From<PackageError> for TransactionError {
    fn from(error: PackageError) -> Self {
        Self::Invalid(error)
    }
}

impl From<PatchError> for TransactionError {
    fn from(error: PatchError) -> Self {
        Self::Patch(error)
    }
}

fn allocate_key(collection: &Collection, kind: Kind) -> Result<FileNameKey, TransactionError> {
    for identifier in 0..=0x0FFEu16 {
        let key = FileNameKey::try_new(kind, identifier).map_err(key_error)?;
        if !collection
            .referenced_files
            .iter()
            .any(|file| file.key() == key)
        {
            return Ok(key);
        }
    }
    Err(TransactionError::Invalid(PackageError::Corrupted(
        "no unused FNPI identifier remains".to_string(),
    )))
}

fn raw_fnfb(metadata: FileNameMetadata) -> u8 {
    u8::from(metadata.valid_on_fat) * codec::FNFB_FAT
        | u8::from(metadata.valid_on_ntfs) * codec::FNFB_NTFS
        | u8::from(metadata.is_non_file_system_path) * codec::FNFB_NON_FILE_SYS
}

fn key_error(error: FileNameKeyError) -> TransactionError {
    TransactionError::Invalid(PackageError::Corrupted(error.to_string()))
}

fn patch_error(error: PatchError) -> PackageError {
    PackageError::Corrupted(error.to_string())
}
