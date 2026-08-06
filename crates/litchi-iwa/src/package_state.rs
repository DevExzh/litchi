//! Shared indexed storage for immutable iWork package snapshots.

use std::sync::Arc;

use crate::archive::{Archive, ArchiveLimits};
use crate::{Error, Result};
use litchi_iwa_cache::{GetOrInsertError, ParseError, WeightedCache};
use litchi_iwa_package::{Entry, EntryStore, Error as EntryStoreError};

/// Package entries plus an index for validated name lookups.
///
/// The state is kept behind an `Arc` by both the mutable package and immutable
/// snapshot types. Structural edits clone this state once, then rebuild the
/// small name index; read-only clones never duplicate either the entry bytes,
/// the index, or the completed values in the bounded parsed-archive cache.
/// Active parser flights remain owned by their source generation.
#[derive(Debug)]
pub(crate) struct PackageState {
    pub(crate) entries: EntryStore,
    archive_limits: ArchiveLimits,
    parsed_archive: WeightedCache<String, Archive>,
}

impl Default for PackageState {
    fn default() -> Self {
        Self::from_store(EntryStore::default(), ArchiveLimits::default())
    }
}

pub(crate) fn entry_store_error(error: EntryStoreError) -> Error {
    match error {
        EntryStoreError::DuplicateEntry(name) => {
            Error::Bundle(format!("Duplicate package entry is ambiguous: {name}"))
        },
        EntryStoreError::InvalidPosition { position, len } => Error::Bundle(format!(
            "Package entry position {position} is outside a table of length {len}"
        )),
        EntryStoreError::Allocation { requested } => Error::Bundle(format!(
            "Failed to allocate package entry index for {requested} entries"
        )),
        error => Error::Bundle(format!("Package entry storage error: {error}")),
    }
}

pub(crate) fn package_patch_error(error: EntryStoreError) -> Error {
    if matches!(error, EntryStoreError::PatchSourceMismatch) {
        Error::InvalidFormat("iWork package patch source does not match".to_owned())
    } else {
        entry_store_error(error)
    }
}

impl Clone for PackageState {
    fn clone(&self) -> Self {
        Self {
            entries: self.entries.clone(),
            archive_limits: self.archive_limits,
            // Completed values share their immutable Arcs, while active
            // flights are tied to the source generation's exact entry bytes
            // and must never be carried into the detached state.
            parsed_archive: self.parsed_archive.fork(),
        }
    }
}

impl PackageState {
    pub(crate) fn from_entries(entries: Vec<Entry>, archive_limits: ArchiveLimits) -> Result<Self> {
        let entries = EntryStore::try_from_entries(entries).map_err(entry_store_error)?;
        Ok(Self::from_store(entries, archive_limits))
    }

    pub(crate) fn from_store(entries: EntryStore, archive_limits: ArchiveLimits) -> Self {
        Self {
            entries,
            archive_limits,
            parsed_archive: new_archive_cache(archive_limits),
        }
    }

    pub(crate) fn position(&self, name: &str) -> Option<usize> {
        self.entries.position(name)
    }

    /// Parse or wait for one archive without holding the cache lock during
    /// decompression or protobuf parsing. The decompressed stream byte length
    /// is charged as the retained cache weight.
    pub(crate) fn get_or_parse_archive<F>(&self, name: &str, parse: F) -> Result<Arc<Archive>>
    where
        F: FnOnce(ArchiveLimits) -> Result<(Archive, usize)>,
    {
        self.parsed_archive
            .get_or_try_insert_with_weight(name.to_owned(), || {
                parse(self.archive_limits).map_err(|error| Box::new(error) as ParseError)
            })
            .map_err(cache_error)
    }

    pub(crate) fn invalidate_archive(&mut self, name: &str) {
        let _removed = self.parsed_archive.invalidate(&name.to_owned());
    }
}

fn new_archive_cache(archive_limits: ArchiveLimits) -> WeightedCache<String, Archive> {
    WeightedCache::new(archive_limits.max_archive_bytes()).unwrap_or_else(|error| {
        unreachable!("validated IWA archive limits cannot create a zero-weight cache: {error}")
    })
}

fn cache_error(error: GetOrInsertError) -> Error {
    if let GetOrInsertError::Parse(shared) = error {
        if let Some(error) = shared.downcast_ref::<Error>() {
            return clone_error(error);
        }
        return Error::ParseError(shared.to_string());
    }
    Error::Archive(format!("IWA archive cache lookup failed: {error}"))
}

fn clone_error(error: &Error) -> Error {
    match error {
        Error::Io(error) => Error::Io(std::io::Error::new(error.kind(), error.to_string())),
        Error::SourceChanged { expected, observed } => Error::SourceChanged {
            expected: *expected,
            observed: *observed,
        },
        Error::IwaCore(error) => Error::IwaCore(error.clone()),
        Error::IwaCommon(error) => Error::IwaCommon(error.clone()),
        Error::PagesSemantic(error) => Error::PagesSemantic(error.clone()),
        Error::TextHyperlink(error) => Error::TextHyperlink(*error),
        Error::TextNumberAttachment(error) => Error::TextNumberAttachment(*error),
        Error::TextComment(error) => Error::TextComment(*error),
        Error::ParagraphStyle(error) => Error::ParagraphStyle(*error),
        Error::InvalidFormat(message) => Error::InvalidFormat(message.clone()),
        Error::Snappy(message) => Error::Snappy(message.clone()),
        Error::ProtobufDecode(error) => Error::ProtobufDecode(error.clone()),
        Error::UnsupportedMessageType(type_) => Error::UnsupportedMessageType(*type_),
        Error::Archive(message) => Error::Archive(message.clone()),
        Error::Bundle(message) => Error::Bundle(message.clone()),
        Error::ParseError(message) => Error::ParseError(message.clone()),
    }
}
