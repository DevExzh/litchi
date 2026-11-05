//! Shared exact-source working set of the structural table transactions
//! (record 0301, moved from `litchi-numbers` in record 0379).
//!
//! A structural edit stages its changes on a [`Working`] set: every
//! component it touches is decompressed and parsed once from the exact
//! source, canonical framing is checked, and the edit mutates the parsed
//! objects in place. Objects the edit does not touch keep their exact
//! encoded bytes, and members the edit does not touch keep their exact
//! compressed bytes. [`Working::publish`] re-encodes only the touched
//! components, rewrites the package metadata registry when objects were
//! added or removed, drops the stale root previews and reassembles the ZIP
//! once. The leaf reopens the candidate through its ordinary strict ingress
//! ([`TableLeaf::reopen_candidate`]) before any family accepts it.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the working set keeps its helpers beside the phase that uses them"
)]

pub mod budget;
pub mod collect;
pub mod create;
pub mod duplicate;
pub mod formula;
pub mod list_workspace;
pub mod metadata;
pub mod plan;
pub mod references;
/// Exact wire helpers of the structural rewrites (moved from
/// `litchi-numbers`' table topology, record 0412): every rewrite keeps the
/// bytes of each field it does not change.
pub mod wire;

use std::collections::{BTreeMap, BTreeSet, HashMap};
use std::sync::Arc;

use litchi_iwa_archive::SourceCatalog;
use litchi_iwa_archive::package::{EntryEdit, EntryInsertion, SharedBytes};
use litchi_iwa_core::{Archive, ArchiveObject, SnappyStream};

use crate::TableSource;

/// A leaf package a structural transaction publishes: the engines' read
/// view plus the leaf's strict candidate reopen.
pub trait TableLeaf: TableSource + Sized {
    /// Whether a first volatile conditional-highlight rule may complete a
    /// saved document's owner family. Numbers has native evidence (record
    /// 0432); Keynote and Pages keep refusing, as the host does.
    const CREATES_FIRST_VOLATILE_TILE: bool = false;

    /// Reopen exact candidate bytes under this package's ingress profile.
    ///
    /// # Errors
    ///
    /// [`Fault`] when the candidate fails the leaf's strict ingress.  Leaves
    /// retain bounded resource failures (`Limit` and `Allocation`) so the
    /// focused public transaction does not collapse a typed refusal into
    /// generic verification failure.
    fn reopen_candidate(&self, candidate: Arc<[u8]>) -> FaultResult<Self>;
}

/// The root previews a changed structure makes stale.
pub const ROOT_PREVIEWS: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];

/// A content-free failure of a structural transaction, mapped by each
/// family onto its public `Error`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Fault {
    /// The source cannot publish an exact changed artifact (a semantic-only
    /// or opaque source, or a native shape the transaction does not own).
    Unsupported,
    /// The rooted graph or a native payload is malformed or ambiguous.
    Invalid,
    /// A surviving object depends on what the transaction would remove.
    Dependency,
    /// A finite resource ceiling was exceeded.
    Limit {
        /// The figure that exceeded its ceiling.
        kind: FaultLimit,
        /// The observed or requested amount.
        observed: u64,
        /// The ceiling.
        maximum: u64,
    },
    /// A bounded allocation failed.
    Allocation {
        /// The requested amount.
        amount: usize,
    },
    /// The reopened candidate did not reproduce the requested state.
    Verification,
}

/// The finite figures a structural transaction charges.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FaultLimit {
    /// Complete source or candidate package bytes.
    PackageBytes,
    /// Bytes in one decoded component.
    PayloadBytes,
    /// Object identifiers.
    Identifiers,
}

pub type FaultResult<T> = Result<T, Fault>;

impl From<litchi_numbers_wire::table_topology::Error> for Fault {
    fn from(error: litchi_numbers_wire::table_topology::Error) -> Self {
        use litchi_numbers_wire::table_topology::Error as E;
        match error {
            E::Unsupported => Self::Unsupported,
            E::Allocation { amount } => Self::Allocation { amount },
            _ => Self::Invalid,
        }
    }
}

pub fn allocation(amount: usize) -> Fault {
    Fault::Allocation { amount }
}

/// Reserve `amount` more slots in `vector`, or fail with [`Fault::Allocation`].
pub fn reserve<T>(vector: &mut Vec<T>, amount: usize) -> FaultResult<()> {
    vector
        .try_reserve(amount)
        .map_err(|_error| allocation(amount))
}

/// Copy `bytes` into a new owned buffer under a bounded allocation.
pub fn copy_bytes(bytes: &[u8]) -> FaultResult<Vec<u8>> {
    let mut output = Vec::new();
    reserve(&mut output, bytes.len())?;
    output.extend_from_slice(bytes);
    Ok(output)
}

pub fn map_core(error: &litchi_iwa_core::Error) -> Fault {
    match error {
        litchi_iwa_core::Error::Allocation { requested, .. } => allocation(*requested),
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => Fault::Limit {
            kind: match kind {
                litchi_iwa_core::LimitKind::ArchiveBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes
                | litchi_iwa_core::LimitKind::SnappyFrames => FaultLimit::PackageBytes,
                litchi_iwa_core::LimitKind::Objects | litchi_iwa_core::LimitKind::MetadataItems => {
                    FaultLimit::Identifiers
                },
                litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject
                | litchi_iwa_core::LimitKind::ObjectBytes
                | litchi_iwa_core::LimitKind::MessageBytes
                | litchi_iwa_core::LimitKind::HeaderBytes
                | litchi_iwa_core::LimitKind::HeaderFields
                | litchi_iwa_core::LimitKind::HeaderNesting
                | litchi_iwa_core::LimitKind::HeaderMemoryBytes
                | litchi_iwa_core::LimitKind::SnappyChunkBytes
                | litchi_iwa_core::LimitKind::SnappyStreamBytes => FaultLimit::PayloadBytes,
            },
            observed: u64::try_from(*observed).unwrap_or(u64::MAX),
            maximum: u64::try_from(*maximum).unwrap_or(u64::MAX),
        },
        _ => Fault::Invalid,
    }
}

pub fn map_archive(error: &litchi_iwa_archive::Error) -> Fault {
    match error {
        litchi_iwa_archive::Error::Limit {
            observed, maximum, ..
        } => Fault::Limit {
            kind: FaultLimit::PackageBytes,
            observed: *observed,
            maximum: *maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => allocation(*amount),
        _ => Fault::Unsupported,
    }
}

pub fn map_wire(error: &litchi_iwa_common::Error) -> Fault {
    match error {
        litchi_iwa_common::Error::Allocation { amount, .. } => allocation(*amount),
        _ => Fault::Invalid,
    }
}

/// The exact physical source of `package`, or [`Fault::Unsupported`] for a
/// semantic-only or non-exact source.
pub fn exact_source(package: &dyn TableSource) -> FaultResult<&SourceCatalog> {
    let catalog = package.physical().ok_or(Fault::Unsupported)?;
    if !catalog.source_is_exact() {
        return Err(Fault::Unsupported);
    }
    Ok(catalog)
}

/// The shared owner of `package`'s exact source bytes.
pub fn source_owner(package: &dyn TableSource) -> FaultResult<SharedBytes> {
    Ok(package.physical().ok_or(Fault::Unsupported)?.source_owner())
}

/// The root previews present in `catalog` (all three, or none).
pub fn root_previews(catalog: &SourceCatalog) -> FaultResult<Vec<&'static str>> {
    let mut previews = Vec::new();
    reserve(&mut previews, ROOT_PREVIEWS.len())?;
    for name in ROOT_PREVIEWS {
        match catalog
            .package()
            .iter()
            .filter(|entry| entry.name() == name)
            .count()
        {
            0 => {},
            1 => previews.push(name),
            _ => return Err(Fault::Invalid),
        }
    }
    Ok(previews)
}

/// Copy `data` without the fields in `excluded`.
pub fn strip_fields(data: &[u8], excluded: &[u32]) -> FaultResult<Vec<u8>> {
    let view =
        litchi_iwa_common::wire::parse_wire_view(data, litchi_iwa_common::WireLimits::DEFAULT)
            .map_err(|error| map_wire(&error))?;
    let mut output = Vec::new();
    reserve(&mut output, data.len())?;
    for field in view.fields() {
        if !excluded.contains(&field.number()) {
            output.extend_from_slice(field.raw());
        }
    }
    Ok(output)
}

/// Whether `data` holds field `number`.
pub fn has_field(data: &[u8], number: u32) -> FaultResult<bool> {
    let view =
        litchi_iwa_common::wire::parse_wire_view(data, litchi_iwa_common::WireLimits::DEFAULT)
            .map_err(|error| map_wire(&error))?;
    Ok(view.fields().any(|field| field.number() == number))
}

/// Set singular varint field `number` of `data` to `value`.
pub fn put_varint(data: &[u8], number: u32, value: u64) -> FaultResult<Vec<u8>> {
    let present = has_field(data, number)?;
    litchi_iwa_common::wire::patch_varint_field(data, number, present, Some(value))
        .map_err(|error| map_wire(&error))
}

/// Set singular length-delimited field `number` of `data` to `payload`.
pub fn put_bytes(data: &[u8], number: u32, payload: &[u8]) -> FaultResult<Vec<u8>> {
    let present = has_field(data, number)?;
    litchi_iwa_common::wire::patch_length_delimited_field(data, number, present, Some(payload))
        .map_err(|error| map_wire(&error))
}

/// A new single-message object whose header copies `template`'s message
/// versions and merge flag, declaring `references`.
pub fn new_object(
    identifier: u64,
    template: Option<&ArchiveObject>,
    message_type: u32,
    data: Vec<u8>,
    references: Vec<u64>,
    limits: litchi_iwa_core::Limits,
) -> FaultResult<ArchiveObject> {
    let mut object = ArchiveObject::new(
        identifier,
        vec![litchi_iwa_core::RawMessage {
            type_: message_type,
            data,
        }],
        limits,
    )
    .map_err(|error| map_core(&error))?;
    if let Some(template) = template {
        object.archive_info.should_merge = template.archive_info.should_merge;
        if let Some(info) = template.archive_info.message_infos.first() {
            object.archive_info.message_infos[0].versions = info.versions.clone();
        }
    }
    object.archive_info.message_infos[0].object_references = references;
    Ok(object)
}

/// The unique message of `object` whose type is one of `types`.
pub fn unique_message(object: &ArchiveObject, types: &[u32]) -> FaultResult<usize> {
    let mut found = None;
    for (index, message) in object.messages.iter().enumerate() {
        if types.contains(&message.type_) && found.replace(index).is_some() {
            return Err(Fault::Invalid);
        }
    }
    found.ok_or(Fault::Invalid)
}

/// One component the working set has parsed for editing.
struct Edited {
    archive: Archive,
    /// Whether an edit changed the archive.
    dirty: bool,
}

/// The staged state of one structural transaction over an exact source.
pub struct Working<'p, P> {
    source: &'p P,
    catalog: &'p SourceCatalog,
    archive_limits: litchi_iwa_core::Limits,
    edited: BTreeMap<String, Edited>,
    /// Objects added by this transaction and the component holding them.
    inserted: HashMap<u64, String>,
    /// Source objects removed by this transaction.
    removed: BTreeSet<u64>,
    /// The highest identifier this transaction allocated.
    allocated: Option<u64>,
    /// The next identifier to allocate.
    next_identifier: u64,
    /// Whether [`Self::publish`] keeps the root previews (record 0370: the
    /// cell-level families keep them, as the host editor does).
    keep_previews: bool,
    /// The package save token the publication writes, for a transaction
    /// composed of several publications (record 0370g).
    save_token: Option<u64>,
}

impl<'p, P: TableLeaf> Working<'p, P> {
    /// Start a working set over `source`'s exact artifact.
    pub fn new(source: &'p P) -> FaultResult<Self> {
        let catalog = exact_source(source)?;
        let archive_limits = catalog
            .limits()
            .effective_archive_limits()
            .map_err(|error| map_archive(&error))?;
        let mut maximum = 0_u64;
        for object in source
            .catalog()
            .iter()
            .flat_map(|component| component.archive().objects.iter())
        {
            maximum = maximum.max(object.archive_info.identifier.ok_or(Fault::Invalid)?);
        }
        let mut working = Self {
            source,
            catalog,
            archive_limits,
            edited: BTreeMap::new(),
            inserted: HashMap::new(),
            removed: BTreeSet::new(),
            allocated: None,
            next_identifier: 0,
            keep_previews: false,
            save_token: None,
        };
        let registered = metadata::last_object_identifier(&working)?.unwrap_or(0);
        working.next_identifier = maximum.max(registered).checked_add(1).ok_or(Fault::Limit {
            kind: FaultLimit::Identifiers,
            observed: u64::MAX,
            maximum: u64::MAX,
        })?;
        Ok(working)
    }

    /// The source package.
    pub const fn source(&self) -> &'p P {
        self.source
    }

    /// The archive limits every touched component is parsed and encoded
    /// under.
    pub const fn archive_limits(&self) -> litchi_iwa_core::Limits {
        self.archive_limits
    }

    /// The current archive of component `name`.
    pub fn archive(&self, name: &str) -> FaultResult<&Archive> {
        if let Some(edited) = self.edited.get(name) {
            return Ok(&edited.archive);
        }
        self.source
            .catalog()
            .get(name)
            .map(litchi_iwa_archive::Component::archive)
            .ok_or(Fault::Invalid)
    }

    /// The names of every component, in catalog order.
    pub fn component_names(&self) -> Vec<String> {
        self.source
            .catalog()
            .iter()
            .map(|component| component.name().to_owned())
            .collect()
    }

    /// The component holding object `identifier`, if it exists.
    ///
    /// An identifier stored in more than one component (an alias the native
    /// reader coalesces) is refused: a structural edit never chooses one copy.
    pub fn locate(&self, identifier: u64) -> FaultResult<Option<String>> {
        if self.removed.contains(&identifier) {
            return Ok(None);
        }
        if let Some(name) = self.inserted.get(&identifier) {
            return Ok(Some(name.clone()));
        }
        let mut found: Option<&str> = None;
        for component in self.source.catalog().iter() {
            if component.archive().object(identifier).is_some()
                && found.replace(component.name()).is_some()
            {
                return Err(Fault::Unsupported);
            }
        }
        Ok(found.map(str::to_owned))
    }

    /// The current object `identifier`.
    pub fn object(&self, identifier: u64) -> FaultResult<&ArchiveObject> {
        let name = self.locate(identifier)?.ok_or(Fault::Invalid)?;
        self.archive(&name)?
            .object(identifier)
            .ok_or(Fault::Invalid)
    }

    /// Parse component `name` from its exact compressed member for editing.
    fn edit(&mut self, name: &str) -> FaultResult<&mut Edited> {
        if !self.edited.contains_key(name) {
            let entry = self
                .catalog
                .package()
                .iter()
                .find(|entry| entry.name() == name)
                .ok_or(Fault::Invalid)?;
            if entry.is_opaque() {
                return Err(Fault::Unsupported);
            }
            let snappy = self
                .catalog
                .limits()
                .snappy_limits()
                .map_err(|error| map_archive(&error))?;
            let stream =
                SnappyStream::decompress(entry.data(), snappy).map_err(|error| map_core(&error))?;
            let archive = Archive::parse(stream.as_bytes(), self.archive_limits)
                .map_err(|error| map_core(&error))?;
            archive
                .validate_canonical_object_framing(stream.as_bytes(), self.archive_limits)
                .map_err(|error| map_core(&error))?;
            self.edited.insert(
                name.to_owned(),
                Edited {
                    archive,
                    dirty: false,
                },
            );
        }
        self.edited.get_mut(name).ok_or(Fault::Invalid)
    }

    /// The archive of component `name`, parsed for editing and marked
    /// changed.
    pub fn archive_mut(&mut self, name: &str) -> FaultResult<&mut Archive> {
        let edited = self.edit(name)?;
        edited.dirty = true;
        Ok(&mut edited.archive)
    }

    /// The object `identifier`, parsed for editing; its component is marked
    /// changed.
    pub fn object_mut(&mut self, identifier: u64) -> FaultResult<&mut ArchiveObject> {
        let name = self.locate(identifier)?.ok_or(Fault::Invalid)?;
        self.archive_mut(&name)?
            .object_mut(identifier)
            .ok_or(Fault::Invalid)
    }

    /// The identifier the next [`Self::allocate`] returns first (a planner
    /// that names the objects it creates before they are allocated, record
    /// 0376b).
    #[must_use]
    pub const fn next_identifier(&self) -> u64 {
        self.next_identifier
    }

    /// Allocate `count` fresh consecutive identifiers above every identifier
    /// the source holds or its registry has issued.
    pub fn allocate(&mut self, count: usize) -> FaultResult<Vec<u64>> {
        let mut identifiers = Vec::new();
        reserve(&mut identifiers, count)?;
        for _ in 0..count {
            let identifier = self.next_identifier;
            self.next_identifier = identifier.checked_add(1).ok_or(Fault::Limit {
                kind: FaultLimit::Identifiers,
                observed: u64::MAX,
                maximum: u64::MAX,
            })?;
            self.allocated = Some(identifier);
            identifiers.push(identifier);
        }
        Ok(identifiers)
    }

    /// Insert a new object into component `name`.
    pub fn insert(&mut self, name: &str, object: ArchiveObject) -> FaultResult<()> {
        let identifier = object.archive_info.identifier.ok_or(Fault::Invalid)?;
        if self.locate(identifier)?.is_some() {
            return Err(Fault::Invalid);
        }
        let limits = self.archive_limits;
        self.archive_mut(name)?
            .insert_object(object, limits)
            .map_err(|error| map_core(&error))?;
        self.inserted.insert(identifier, name.to_owned());
        Ok(())
    }

    /// Replace object `identifier` in place in component `name`, or insert
    /// it there when the component does not hold it.
    pub fn upsert(&mut self, name: &str, object: ArchiveObject) -> FaultResult<()> {
        let identifier = object.archive_info.identifier.ok_or(Fault::Invalid)?;
        match self.locate(identifier)? {
            Some(current) if current == name => {
                let limits = self.archive_limits;
                self.archive_mut(name)?
                    .upsert_object(object, limits)
                    .map_err(|error| map_core(&error))?;
                Ok(())
            },
            Some(_) => Err(Fault::Invalid),
            None => self.insert(name, object),
        }
    }

    /// Remove object `identifier` from the component holding it.
    pub fn remove(&mut self, identifier: u64) -> FaultResult<ArchiveObject> {
        let name = self.locate(identifier)?.ok_or(Fault::Invalid)?;
        let limits = self.archive_limits;
        let removed = self
            .archive_mut(&name)?
            .remove_object_checked(identifier, limits)
            .map_err(|error| map_core(&error))?;
        if self.inserted.remove(&identifier).is_none() {
            self.removed.insert(identifier);
        }
        Ok(removed)
    }

    /// The largest identifier of an object the working set holds (source
    /// objects not removed, and inserted ones).
    pub fn maximum_identifier(&self) -> u64 {
        let source = self
            .source
            .catalog()
            .iter()
            .flat_map(|component| component.archive().objects.iter())
            .filter_map(|object| object.archive_info.identifier)
            .filter(|identifier| !self.removed.contains(identifier))
            .max()
            .unwrap_or(0);
        self.inserted.keys().copied().max().unwrap_or(0).max(source)
    }

    /// Keep the root previews on [`Self::publish`] instead of deleting them
    /// as stale: a cell-level edit (topology, record 0370) leaves them in
    /// place, as the host editor and the native application's own saves
    /// before it regenerates them do.
    /// Write `token` as the package save token on publication: a
    /// transaction composed of several publications advances its source's
    /// token once (record 0370g).
    pub fn set_save_token(&mut self, token: u64) {
        self.save_token = Some(token);
    }

    /// The source package's save token, when it has package metadata.
    pub fn save_token(&self) -> FaultResult<Option<u64>> {
        metadata::save_token(self)
    }

    pub fn keep_previews(&mut self) {
        self.keep_previews = true;
    }

    /// Re-encode the touched components, rewrite the metadata registry, drop
    /// the stale root previews and reassemble one candidate.
    ///
    /// `registry` describes the registry transition the caller proved: the
    /// UUID and external-reference records of removed objects, and the
    /// external references new objects need. Returns the reopened candidate,
    /// the number of rewritten components and the number of deleted previews.
    pub fn publish(mut self, registry: &metadata::Transition) -> FaultResult<Published<P>> {
        let previews = if self.keep_previews {
            Vec::new()
        } else {
            root_previews(self.catalog)?
        };
        if !previews.is_empty() && previews.len() != ROOT_PREVIEWS.len() {
            return Err(Fault::Unsupported);
        }
        let last = self.allocated;
        let emptied: Vec<String> = self
            .edited
            .iter()
            .filter(|(_, edited)| edited.dirty && edited.archive.objects.is_empty())
            .map(|(name, _)| name.clone())
            .collect();
        metadata::apply(&mut self, registry, last, &emptied)?;
        let snappy_limits = self
            .catalog
            .limits()
            .snappy_limits()
            .map_err(|error| map_archive(&error))?;
        let mut encoded: Vec<(String, Vec<u8>)> = Vec::new();
        for (name, edited) in &self.edited {
            if !edited.dirty || emptied.contains(name) {
                continue;
            }
            let bytes = edited
                .archive
                .to_bytes(self.archive_limits)
                .map_err(|error| map_core(&error))?;
            let compressed =
                SnappyStream::compress(&bytes, snappy_limits).map_err(|error| map_core(&error))?;
            encoded.push((name.clone(), compressed));
        }
        for (name, data) in &registry.entry_writes {
            if encoded.iter().any(|(existing, _)| existing == name) {
                return Err(Fault::Invalid);
            }
            reserve(&mut encoded, 1)?;
            encoded.push((name.clone(), copy_bytes(data)?));
        }
        let mut deleted: Vec<&str> = previews.clone();
        deleted.extend(emptied.iter().map(String::as_str));
        deleted.extend(registry.deleted_entries());
        let insertions: Vec<EntryInsertion<'_>> = registry
            .entry_insertions
            .iter()
            .map(|(name, data)| EntryInsertion::new(name, data))
            .collect();
        let edits: Vec<EntryEdit<'_>> = encoded
            .iter()
            .map(|(name, data)| EntryEdit::new(name, data))
            .collect();
        let output = self
            .catalog
            .package()
            .reassemble_with_changes_to_bytes(&insertions, &edits, &deleted, self.catalog.limits())
            .map_err(|error| map_archive(&error))?;
        let bytes: Arc<[u8]> = output.into();
        let candidate = self.source.reopen_candidate(Arc::clone(&bytes))?;
        let inserted: Vec<&str> = registry
            .entry_insertions
            .iter()
            .map(|(name, _)| name.as_str())
            .collect();
        verify_locality(self.catalog, &candidate, &encoded, &deleted, &inserted)?;
        Ok(Published {
            package: candidate,
            touched_components: encoded.len() + emptied.len(),
            deleted_previews: previews.len(),
        })
    }
}

/// One reassembled, reopened candidate.
pub struct Published<P> {
    pub package: P,
    pub touched_components: usize,
    pub deleted_previews: usize,
}

/// Every member the transaction did not rewrite or delete keeps its exact
/// decoded bytes, in the same order.
fn verify_locality(
    source: &SourceCatalog,
    candidate: &dyn TableSource,
    rewritten: &[(String, Vec<u8>)],
    deleted: &[&str],
    inserted: &[&str],
) -> FaultResult<()> {
    let after = candidate.physical().ok_or(Fault::Verification)?.package();
    let mut retained = source
        .package()
        .iter()
        .filter(|entry| !deleted.contains(&entry.name()));
    let mut found = 0_usize;
    for entry in after.iter() {
        if inserted.contains(&entry.name()) {
            found += 1;
            continue;
        }
        let before = retained.next().ok_or(Fault::Verification)?;
        if before.name() != entry.name() {
            return Err(Fault::Verification);
        }
        let changed = rewritten.iter().any(|(name, _)| name == entry.name());
        if !changed && before.data() != entry.data() {
            return Err(Fault::Verification);
        }
    }
    if retained.next().is_some() || found != inserted.len() {
        return Err(Fault::Verification);
    }
    Ok(())
}
