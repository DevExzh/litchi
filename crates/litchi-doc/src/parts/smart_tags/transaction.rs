//! Source-checked, failure-atomic smart-tag transactions.

use super::codec;
use super::semantic::Topology;
use super::validation;
use super::{DocumentSmartTag, DocumentSmartTags, FileInformationBlock, SmartTagBookmarkInfo};
use crate::package::{Error as PackageError, Result as PackageResult};
use litchi_codepage::Ansi;
use litchi_ole_common::smart_tags::{Property, PropertyBag, PropertyBagString, Type};
use std::sync::Arc;

use super::Limits;

type TxResult<T> = Result<T, Error>;

/// An immutable snapshot of the selected Word table stream and its FIB.
///
/// The snapshot keeps the complete source allocations, not only the five
/// decoded smart-tag ranges. This makes exact no-op publication cheap and
/// ensures unrelated table bytes, ignored `pfpb` fields, and all FIB pointers
/// survive a typed edit unchanged.
#[derive(Debug, Clone)]
pub struct Snapshot {
    fib: Arc<FileInformationBlock>,
    table_stream: Arc<[u8]>,
    metadata: DocumentSmartTags,
    topology: Topology,
    links: Arc<[u16]>,
    limits: Limits,
    revision: u64,
}

impl Snapshot {
    /// Parse a Word smart-tag view with the document's LCID-derived ANSI page.
    pub fn parse(fib: &FileInformationBlock, table_stream: &[u8]) -> PackageResult<Self> {
        Self::parse_with_options(fib, table_stream, None, Limits::default())
    }

    /// Parse with an explicit ANSI page and default resource limits.
    pub fn parse_with(
        fib: &FileInformationBlock,
        table_stream: &[u8],
        ansi: Ansi,
    ) -> PackageResult<Self> {
        Self::parse_with_options(fib, table_stream, Some(ansi), Limits::default())
    }

    /// Parse with custom resource limits and LCID-derived ANSI decoding.
    pub fn parse_with_limits(
        fib: &FileInformationBlock,
        table_stream: &[u8],
        limits: Limits,
    ) -> PackageResult<Self> {
        Self::parse_with_options(fib, table_stream, None, limits)
    }

    /// Parse with explicit ANSI decoding and resource limits.
    pub fn parse_with_options(
        fib: &FileInformationBlock,
        table_stream: &[u8],
        ansi: Option<Ansi>,
        limits: Limits,
    ) -> PackageResult<Self> {
        if table_stream.len() > limits.max_bytes {
            return Err(corrupted(
                "smart-tag table stream exceeds the configured size limit",
            ));
        }
        let metadata = ansi
            .map(|ansi| DocumentSmartTags::parse_with(fib, table_stream, ansi))
            .unwrap_or_else(|| DocumentSmartTags::parse(fib, table_stream))?
            .unwrap_or_else(empty_metadata);
        let topology = Topology::capture(fib, table_stream)?;
        let links = codec::bookmark_links(&topology, table_stream)?;
        validation::source(&topology, &metadata, table_stream, &links, limits)?;
        Ok(Self::from_validated(
            Arc::new(fib.clone()),
            Arc::<[u8]>::from(table_stream),
            metadata,
            topology,
            links,
            limits,
        ))
    }

    /// Exact raw FIB bytes captured by this snapshot.
    #[must_use]
    pub fn fib_bytes(&self) -> &[u8] {
        self.fib.raw_data()
    }

    /// Exact raw selected table-stream bytes captured by this snapshot.
    #[must_use]
    pub fn table_stream(&self) -> &[u8] {
        &self.table_stream
    }

    /// Alias for [`Self::table_stream`].
    #[must_use]
    pub fn table_bytes(&self) -> &[u8] {
        self.table_stream()
    }

    /// Typed, inert smart-tag metadata in source order.
    #[must_use]
    pub fn metadata(&self) -> &DocumentSmartTags {
        &self.metadata
    }

    /// Smart-tag bookmark metadata, when the parallel tables are present.
    #[must_use]
    pub fn tags(&self) -> &[DocumentSmartTag] {
        &self.metadata.tags
    }

    /// Recognizer-state ranges in `Plcffactoid` order.
    #[must_use]
    pub fn recognizer_ranges(&self) -> &[super::SmartTagRecognizerRange] {
        &self.metadata.recognizer_ranges
    }

    /// Shared property-bag store, when `FactoidData` is present.
    #[must_use]
    pub fn store(&self) -> Option<&litchi_ole_common::smart_tags::PropertyBagStore> {
        self.metadata.store.as_ref()
    }

    /// The unchanged FIB/table topology captured at parse time.
    #[must_use]
    pub const fn topology(&self) -> &Topology {
        &self.topology
    }

    /// Compact source fingerprint used as a first-stage stale check.
    #[must_use]
    pub const fn fingerprint(&self) -> u64 {
        self.revision
    }

    /// Whether any smart-tag table is present in the source FIB.
    #[must_use]
    pub fn is_present(&self) -> bool {
        self.topology
            .range(super::TableKind::BookmarkInfo)
            .is_some()
            || self
                .topology
                .range(super::TableKind::PropertyBags)
                .is_some()
            || self.topology.range(super::TableKind::Recognizer).is_some()
    }

    /// Return the unchanged FIB and table bytes as owned buffers.
    #[must_use]
    pub fn finish(&self) -> (Vec<u8>, Vec<u8>) {
        (self.fib_bytes().to_vec(), self.table_stream().to_vec())
    }

    /// Start an isolated typed transaction.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        Transaction::new(self.clone())
    }

    /// Alias for [`Self::edit`].
    #[must_use]
    pub fn transaction(&self) -> Transaction {
        self.edit()
    }

    pub(super) fn from_validated(
        fib: Arc<FileInformationBlock>,
        table_stream: Arc<[u8]>,
        metadata: DocumentSmartTags,
        topology: Topology,
        links: Vec<u16>,
        limits: Limits,
    ) -> Self {
        let revision = fingerprint(fib.raw_data(), &table_stream);
        Self {
            fib,
            table_stream,
            metadata,
            topology,
            links: links.into(),
            limits,
            revision,
        }
    }

    fn materialize(&self, metadata: DocumentSmartTags) -> PackageResult<Self> {
        validation::candidate(
            &self.topology,
            &self.metadata,
            &metadata,
            &self.links,
            self.limits,
        )?;
        let table_stream = codec::encode(
            &self.table_stream,
            &self.topology,
            &self.metadata,
            &metadata,
            &self.links,
            self.limits,
        )?;
        if metadata == self.metadata && table_stream == self.table_stream.as_ref() {
            return Ok(self.clone());
        }
        let reparsed = if let Some(store) = &metadata.store {
            DocumentSmartTags::parse_with(&self.fib, &table_stream, store.ansi)?
        } else {
            DocumentSmartTags::parse(&self.fib, &table_stream)?
        }
        .unwrap_or_else(empty_metadata);
        if reparsed != metadata {
            return Err(corrupted(
                "smart-tag candidate failed FIB/table-stream round-trip validation",
            ));
        }
        Ok(Self::from_validated(
            Arc::clone(&self.fib),
            Arc::<[u8]>::from(table_stream),
            metadata,
            self.topology.clone(),
            self.links.to_vec(),
            self.limits,
        ))
    }
}

impl PartialEq for Snapshot {
    fn eq(&self, other: &Self) -> bool {
        self.fib_bytes() == other.fib_bytes() && self.table_stream == other.table_stream
    }
}

impl Eq for Snapshot {}

/// A staged, clone-first edit over one source snapshot.
#[derive(Debug, Clone)]
pub struct Transaction {
    source: Snapshot,
    metadata: DocumentSmartTags,
}

impl Transaction {
    pub(super) fn new(source: Snapshot) -> Self {
        Self {
            metadata: source.metadata.clone(),
            source,
        }
    }

    /// The immutable source used for stale-source checks.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.source
    }

    /// The current typed candidate.
    #[must_use]
    pub const fn metadata(&self) -> &DocumentSmartTags {
        &self.metadata
    }

    /// Whether the typed candidate differs from the source projection.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.metadata != *self.source.metadata()
    }

    /// Replace one complete bookmark while retaining its stable table slot.
    pub fn replace_tag(&mut self, index: usize, value: DocumentSmartTag) -> TxResult<&mut Self> {
        self.update(move |metadata| {
            let tag = metadata
                .tags
                .get_mut(index)
                .ok_or_else(|| invalid("smart-tag bookmark index is outside the table"))?;
            *tag = value;
            Ok(())
        })
    }

    /// Change the stable `FACTOIDINFO` metadata for one bookmark.
    pub fn set_bookmark_info(
        &mut self,
        index: usize,
        value: SmartTagBookmarkInfo,
    ) -> TxResult<&mut Self> {
        self.update(move |metadata| {
            let tag = metadata
                .tags
                .get_mut(index)
                .ok_or_else(|| invalid("smart-tag bookmark index is outside the table"))?;
            tag.info = value;
            Ok(())
        })
    }

    /// Move one bookmark while recomputing all depth fields atomically.
    pub fn set_bookmark_range(
        &mut self,
        index: usize,
        start: u32,
        end: u32,
    ) -> TxResult<&mut Self> {
        self.update(move |metadata| {
            let tag = metadata
                .tags
                .get_mut(index)
                .ok_or_else(|| invalid("smart-tag bookmark index is outside the table"))?;
            tag.start = start;
            tag.end = end;
            validation::recompute_depths(&mut metadata.tags)?;
            Ok(())
        })
    }

    /// Replace one shared property bag without changing its bookmark slot.
    pub fn replace_property_bag(
        &mut self,
        index: usize,
        value: PropertyBag,
    ) -> TxResult<&mut Self> {
        self.update(move |metadata| {
            let tag = metadata
                .tags
                .get_mut(index)
                .ok_or_else(|| invalid("smart-tag bookmark index is outside the table"))?;
            tag.property_bag = value;
            Ok(())
        })
    }

    /// Replace one string-table entry used by property bags or types.
    pub fn set_string(&mut self, index: usize, value: PropertyBagString) -> TxResult<&mut Self> {
        self.update(move |metadata| {
            let store = metadata
                .store
                .as_mut()
                .ok_or_else(|| invalid("FactoidData is absent"))?;
            let entry = store
                .strings
                .get_mut(index)
                .ok_or_else(|| invalid("smart-tag string index is outside the table"))?;
            *entry = value;
            Ok(())
        })
    }

    /// Replace a declared recognizer type while retaining its stable ID.
    pub fn replace_type(&mut self, value: Type) -> TxResult<&mut Self> {
        self.update(move |metadata| {
            let store = metadata
                .store
                .as_mut()
                .ok_or_else(|| invalid("FactoidData is absent"))?;
            let kind = store
                .types
                .iter_mut()
                .find(|kind| kind.id == value.id)
                .ok_or_else(|| invalid("smart-tag type ID is not declared"))?;
            *kind = value;
            Ok(())
        })
    }

    /// Replace one raw property key/value index pair.
    pub fn set_property(
        &mut self,
        bag: usize,
        property: usize,
        value: Property,
    ) -> TxResult<&mut Self> {
        self.update(move |metadata| {
            let tag = metadata
                .tags
                .get_mut(bag)
                .ok_or_else(|| invalid("smart-tag bag index is outside the table"))?;
            let entry = tag
                .property_bag
                .properties
                .get_mut(property)
                .ok_or_else(|| invalid("smart-tag property index is outside the bag"))?;
            *entry = value;
            Ok(())
        })
    }

    /// Replace the key string referenced by one property.
    pub fn set_property_key(
        &mut self,
        bag: usize,
        property: usize,
        value: PropertyBagString,
    ) -> TxResult<&mut Self> {
        self.set_property_string(bag, property, value, true)
    }

    /// Replace the value string referenced by one property.
    pub fn set_property_value(
        &mut self,
        bag: usize,
        property: usize,
        value: PropertyBagString,
    ) -> TxResult<&mut Self> {
        self.set_property_string(bag, property, value, false)
    }

    /// Change the uninterpreted `PropertyBagStore` reserved value.
    pub fn set_reserved_factoid_count(&mut self, value: u32) -> TxResult<&mut Self> {
        self.update(move |metadata| {
            let store = metadata
                .store
                .as_mut()
                .ok_or_else(|| invalid("FactoidData is absent"))?;
            store.reserved_factoid_count = value;
            Ok(())
        })
    }

    /// Change one recognizer-state range while retaining its PLC slot.
    pub fn set_recognizer_range(
        &mut self,
        index: usize,
        value: super::SmartTagRecognizerRange,
    ) -> TxResult<&mut Self> {
        self.update(move |metadata| {
            let range = metadata
                .recognizer_ranges
                .get_mut(index)
                .ok_or_else(|| invalid("smart-tag recognizer range is outside the table"))?;
            *range = value;
            Ok(())
        })
    }

    /// Change only one recognizer state.
    pub fn set_recognizer_state(
        &mut self,
        index: usize,
        state: super::SmartTagRecognizerState,
    ) -> TxResult<&mut Self> {
        self.update(move |metadata| {
            let range = metadata
                .recognizer_ranges
                .get_mut(index)
                .ok_or_else(|| invalid("smart-tag recognizer range is outside the table"))?;
            range.state = state;
            Ok(())
        })
    }

    /// Apply a custom inert metadata edit failure-atomically.
    pub fn update<F>(&mut self, edit: F) -> TxResult<&mut Self>
    where
        F: FnOnce(&mut DocumentSmartTags) -> TxResult<()>,
    {
        let mut candidate = self.metadata.clone();
        edit(&mut candidate)?;
        validation::candidate(
            &self.source.topology,
            &self.source.metadata,
            &candidate,
            &self.source.links,
            self.source.limits,
        )?;
        codec::encode(
            &self.source.table_stream,
            &self.source.topology,
            &self.source.metadata,
            &candidate,
            &self.source.links,
            self.source.limits,
        )?;
        self.metadata = candidate;
        Ok(self)
    }

    /// Materialize the staged candidate as a validated snapshot.
    pub fn snapshot(&self) -> TxResult<Snapshot> {
        self.source
            .materialize(self.metadata.clone())
            .map_err(Error::Invalid)
    }

    /// Discard the candidate and recover the immutable source snapshot.
    #[must_use]
    pub fn rollback(self) -> Snapshot {
        self.source
    }

    /// Commit the candidate as a reversible source-checked patch.
    pub fn commit(self) -> TxResult<Commit> {
        let snapshot = self.snapshot()?;
        let patch = Patch::new(self.source, snapshot.clone());
        Ok(Commit { snapshot, patch })
    }

    fn set_property_string(
        &mut self,
        bag: usize,
        property: usize,
        value: PropertyBagString,
        key: bool,
    ) -> TxResult<&mut Self> {
        self.update(move |metadata| {
            let tag = metadata
                .tags
                .get(bag)
                .ok_or_else(|| invalid("smart-tag bag index is outside the table"))?;
            let index = if key {
                tag.property_bag
                    .properties
                    .get(property)
                    .ok_or_else(|| invalid("smart-tag property index is outside the bag"))?
                    .key_index
            } else {
                tag.property_bag
                    .properties
                    .get(property)
                    .ok_or_else(|| invalid("smart-tag property index is outside the bag"))?
                    .value_index
            };
            let index = usize::try_from(index)
                .map_err(|_| invalid("smart-tag string index overflows usize"))?;
            let store = metadata
                .store
                .as_mut()
                .ok_or_else(|| invalid("FactoidData is absent"))?;
            let entry = store
                .strings
                .get_mut(index)
                .ok_or_else(|| invalid("smart-tag string index is outside the table"))?;
            *entry = value;
            Ok(())
        })
    }
}

/// A committed immutable result and its reversible patch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    /// The post-edit snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// The source-checked reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Whether the serialized FIB/table bytes changed.
    #[must_use]
    pub fn changed(&self) -> bool {
        !self.patch.is_noop()
    }

    /// Split the result into the snapshot and patch.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

/// A reversible replacement of the complete smart-tag snapshot projection.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    fn new(before: Snapshot, after: Snapshot) -> Self {
        Self { before, after }
    }

    /// Source snapshot required by this patch.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Replacement snapshot produced by this patch.
    #[must_use]
    pub const fn after(&self) -> &Snapshot {
        &self.after
    }

    /// Whether the exact FIB and table bytes are unchanged.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after
    }

    /// Apply only to the exact source snapshot used to create this patch.
    pub fn apply(&self, source: &Snapshot) -> TxResult<Snapshot> {
        if source.fingerprint() != self.before.fingerprint() || source != &self.before {
            return Err(Error::Conflict);
        }
        Ok(self.after.clone())
    }

    /// Return the exact inverse replacement.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self::new(self.after.clone(), self.before.clone())
    }
}

/// Errors produced by a staged smart-tag edit.
#[derive(Debug)]
pub enum Error {
    /// The candidate violates a DOC, MS-OSHARED, or fixed-topology invariant.
    Invalid(PackageError),
    /// The patch was applied to a different source snapshot.
    Conflict,
}

impl std::fmt::Display for Error {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Invalid(error) => error.fmt(formatter),
            Self::Conflict => formatter.write_str("smart-tag transaction source conflict"),
        }
    }
}

impl std::error::Error for Error {}

impl From<PackageError> for Error {
    fn from(error: PackageError) -> Self {
        Self::Invalid(error)
    }
}

fn empty_metadata() -> DocumentSmartTags {
    DocumentSmartTags {
        store: None,
        tags: Vec::new(),
        recognizer_ranges: Vec::new(),
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(PackageError::InvalidFormat(message.into()))
}

fn corrupted(message: impl Into<String>) -> PackageError {
    PackageError::Corrupted(message.into())
}

fn fingerprint(fib: &[u8], table_stream: &[u8]) -> u64 {
    let mut value = 0xcbf2_9ce4_8422_2325u64;
    for byte in fib.iter().chain(table_stream) {
        value ^= u64::from(*byte);
        value = value.wrapping_mul(0x0000_0100_0000_01b3);
    }
    value
}
