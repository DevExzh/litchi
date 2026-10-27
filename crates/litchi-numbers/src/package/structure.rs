//! Shared exact-source machinery of the structural Numbers transactions
//! (record 0301): [`crate::sheet::lifecycle`], [`crate::table::lifecycle`]
//! and [`crate::table::order`].
//!
//! A structural edit stages its changes on a [`Working`] set: every
//! component it touches is decompressed and parsed once from the exact
//! source, canonical framing is checked, and the edit mutates the parsed
//! objects in place. Objects the edit does not touch keep their exact
//! encoded bytes, and members the edit does not touch keep their exact
//! compressed bytes. [`Working::publish`] re-encodes only the touched
//! components, rewrites the package metadata registry when objects were
//! added or removed, drops the stale root previews and reassembles the ZIP
//! once. The candidate is reopened through the ordinary strict ingress
//! before any family accepts it, and the family's patch retains both exact
//! artifacts, so its inverse restores the source byte for byte.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the working set keeps its helpers beside the phase that uses them"
)]

pub(super) mod budget;
pub(super) mod collect;
pub(super) mod formula;
pub(super) mod metadata;
pub(super) mod plan;
pub(super) mod references;

use std::collections::{BTreeMap, BTreeSet, HashMap};
use std::sync::Arc;

use litchi_iwa_archive::SourceCatalog;
use litchi_iwa_archive::package::{EntryEdit, OwnedExactArtifacts, SharedBytes};
use litchi_iwa_core::{Archive, ArchiveObject, SnappyStream};

use super::{
    DOCUMENT_MESSAGE_TYPE, FORM_BASED_SHEET_MESSAGE_TYPE, LEGACY_TABLE_INFO_MESSAGE_TYPE, Package,
    SHEET_MESSAGE_TYPE, TABLE_INFO_MESSAGE_TYPE,
};

/// The member holding the root document object.
pub(super) const DOCUMENT_COMPONENT: &str = "Index/Document.iwa";
/// The root previews a changed structure makes stale.
pub(super) const ROOT_PREVIEWS: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];
/// `TSK.TreeNode`, the sidebar order nodes.
pub(super) const TREE_NODE_MESSAGE_TYPE: u32 = 205;
/// `TN.DocumentArchive.sheets`.
pub(super) const DOCUMENT_SHEETS_FIELD: u32 = 1;
/// `TN.DocumentArchive.sidebar_order`.
pub(super) const DOCUMENT_SIDEBAR_FIELD: u32 = 5;
/// `TSK.TreeNode.children`.
pub(super) const TREE_CHILDREN_FIELD: u32 = 2;
/// `TSK.TreeNode.object`.
pub(super) const TREE_OBJECT_FIELD: u32 = 3;
/// `TN.SheetArchive.name`.
pub(super) const SHEET_NAME_FIELD: u32 = 1;
/// `TN.SheetArchive.drawable_infos`.
pub(super) const SHEET_DRAWABLES_FIELD: u32 = 2;
/// `TN.FormBasedSheetArchive.super`.
pub(super) const FORM_SHEET_SUPER_FIELD: u32 = 1;

/// A content-free failure of a structural transaction, mapped by each
/// family onto its public `Error`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum Fault {
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
pub(super) enum FaultLimit {
    /// Complete source or candidate package bytes.
    PackageBytes,
    /// Bytes in one decoded component.
    PayloadBytes,
    /// Object identifiers.
    Identifiers,
}

pub(super) type FaultResult<T> = Result<T, Fault>;

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

pub(super) fn allocation(amount: usize) -> Fault {
    Fault::Allocation { amount }
}

/// Reserve `amount` more slots in `vector`, or fail with [`Fault::Allocation`].
pub(super) fn reserve<T>(vector: &mut Vec<T>, amount: usize) -> FaultResult<()> {
    vector
        .try_reserve(amount)
        .map_err(|_error| allocation(amount))
}

/// Copy `bytes` into a new owned buffer under a bounded allocation.
pub(super) fn copy_bytes(bytes: &[u8]) -> FaultResult<Vec<u8>> {
    let mut output = Vec::new();
    reserve(&mut output, bytes.len())?;
    output.extend_from_slice(bytes);
    Ok(output)
}

pub(super) fn map_core(error: &litchi_iwa_core::Error) -> Fault {
    match error {
        litchi_iwa_core::Error::Allocation { requested, .. } => allocation(*requested),
        _ => Fault::Invalid,
    }
}

pub(super) fn map_archive(error: &litchi_iwa_archive::Error) -> Fault {
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

pub(super) fn map_wire(error: &litchi_iwa_common::Error) -> Fault {
    match error {
        litchi_iwa_common::Error::Allocation { amount, .. } => allocation(*amount),
        _ => Fault::Invalid,
    }
}

/// The exact physical source of `package`, or [`Fault::Unsupported`] for a
/// semantic-only or non-exact source.
pub(super) fn exact_source(package: &Package) -> FaultResult<&SourceCatalog> {
    let catalog = package
        .state
        .components
        .physical()
        .ok_or(Fault::Unsupported)?;
    if !catalog.source_is_exact() {
        return Err(Fault::Unsupported);
    }
    Ok(catalog)
}

/// The shared owner of `package`'s exact source bytes.
pub(super) fn source_owner(package: &Package) -> FaultResult<SharedBytes> {
    Ok(package
        .state
        .components
        .physical()
        .ok_or(Fault::Unsupported)?
        .__source_owner())
}

/// The root previews present in `catalog` (all three, or none).
pub(super) fn root_previews(catalog: &SourceCatalog) -> FaultResult<Vec<&'static str>> {
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
pub(super) fn strip_fields(data: &[u8], excluded: &[u32]) -> FaultResult<Vec<u8>> {
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
pub(super) fn has_field(data: &[u8], number: u32) -> FaultResult<bool> {
    let view =
        litchi_iwa_common::wire::parse_wire_view(data, litchi_iwa_common::WireLimits::DEFAULT)
            .map_err(|error| map_wire(&error))?;
    Ok(view.fields().any(|field| field.number() == number))
}

/// Set singular varint field `number` of `data` to `value`.
pub(super) fn put_varint(data: &[u8], number: u32, value: u64) -> FaultResult<Vec<u8>> {
    let present = has_field(data, number)?;
    litchi_iwa_common::wire::patch_varint_field(data, number, present, Some(value))
        .map_err(|error| map_wire(&error))
}

/// Set singular length-delimited field `number` of `data` to `payload`.
pub(super) fn put_bytes(data: &[u8], number: u32, payload: &[u8]) -> FaultResult<Vec<u8>> {
    let present = has_field(data, number)?;
    litchi_iwa_common::wire::patch_length_delimited_field(data, number, present, Some(payload))
        .map_err(|error| map_wire(&error))
}

/// A new single-message object whose header copies `template`'s message
/// versions and merge flag, declaring `references`.
pub(super) fn new_object(
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
pub(super) fn unique_message(object: &ArchiveObject, types: &[u32]) -> FaultResult<usize> {
    let mut found = None;
    for (index, message) in object.messages.iter().enumerate() {
        if types.contains(&message.type_) && found.replace(index).is_some() {
            return Err(Fault::Invalid);
        }
    }
    found.ok_or(Fault::Invalid)
}

/// The sheet message of a sheet object and the field path of its drawable
/// list (`[2]` for a plain sheet, `[1, 2]` for a form-based sheet).
pub(super) fn sheet_message(object: &ArchiveObject) -> FaultResult<(usize, &'static [u32])> {
    let index = unique_message(object, &[SHEET_MESSAGE_TYPE, FORM_BASED_SHEET_MESSAGE_TYPE])?;
    let path: &'static [u32] = if object.messages[index].type_ == SHEET_MESSAGE_TYPE {
        &[SHEET_DRAWABLES_FIELD]
    } else {
        &[FORM_SHEET_SUPER_FIELD, SHEET_DRAWABLES_FIELD]
    };
    Ok((index, path))
}

/// The table-info message of a table drawable.
pub(super) fn table_info_message(object: &ArchiveObject) -> FaultResult<usize> {
    unique_message(
        object,
        &[TABLE_INFO_MESSAGE_TYPE, LEGACY_TABLE_INFO_MESSAGE_TYPE],
    )
}

/// The document message of the root object.
pub(super) fn document_message(object: &ArchiveObject) -> FaultResult<usize> {
    unique_message(object, &[DOCUMENT_MESSAGE_TYPE])
}

/// One component the working set has parsed for editing.
struct Edited {
    archive: Archive,
    /// Whether an edit changed the archive.
    dirty: bool,
}

/// The staged state of one structural transaction over an exact source.
pub(super) struct Working<'p> {
    source: &'p Package,
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
}

impl<'p> Working<'p> {
    /// Start a working set over `source`'s exact artifact.
    pub(super) fn new(source: &'p Package) -> FaultResult<Self> {
        let catalog = exact_source(source)?;
        let archive_limits = catalog
            .limits()
            .effective_archive_limits()
            .map_err(|error| map_archive(&error))?;
        let mut maximum = 0_u64;
        for object in source.state.components.iter_objects() {
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
    pub(super) const fn source(&self) -> &'p Package {
        self.source
    }

    /// The archive limits every touched component is parsed and encoded
    /// under.
    pub(super) const fn archive_limits(&self) -> litchi_iwa_core::Limits {
        self.archive_limits
    }

    /// The current archive of component `name`.
    pub(super) fn archive(&self, name: &str) -> FaultResult<&Archive> {
        if let Some(edited) = self.edited.get(name) {
            return Ok(&edited.archive);
        }
        self.source
            .state
            .components
            .get_archive(name)
            .ok_or(Fault::Invalid)
    }

    /// The names of every component, in catalog order.
    pub(super) fn component_names(&self) -> Vec<String> {
        self.source
            .state
            .components
            .catalog()
            .iter()
            .map(|component| component.name().to_owned())
            .collect()
    }

    /// The component holding object `identifier`, if it exists.
    ///
    /// An identifier stored in more than one component (an alias the native
    /// reader coalesces) is refused: a structural edit never chooses one copy.
    pub(super) fn locate(&self, identifier: u64) -> FaultResult<Option<String>> {
        if self.removed.contains(&identifier) {
            return Ok(None);
        }
        if let Some(name) = self.inserted.get(&identifier) {
            return Ok(Some(name.clone()));
        }
        let mut found: Option<&str> = None;
        for component in self.source.state.components.catalog().iter() {
            if component.archive().object(identifier).is_some()
                && found.replace(component.name()).is_some()
            {
                return Err(Fault::Unsupported);
            }
        }
        Ok(found.map(str::to_owned))
    }

    /// The current object `identifier`.
    pub(super) fn object(&self, identifier: u64) -> FaultResult<&ArchiveObject> {
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
    pub(super) fn archive_mut(&mut self, name: &str) -> FaultResult<&mut Archive> {
        let edited = self.edit(name)?;
        edited.dirty = true;
        Ok(&mut edited.archive)
    }

    /// The object `identifier`, parsed for editing; its component is marked
    /// changed.
    pub(super) fn object_mut(&mut self, identifier: u64) -> FaultResult<&mut ArchiveObject> {
        let name = self.locate(identifier)?.ok_or(Fault::Invalid)?;
        self.archive_mut(&name)?
            .object_mut(identifier)
            .ok_or(Fault::Invalid)
    }

    /// Allocate `count` fresh consecutive identifiers above every identifier
    /// the source holds or its registry has issued.
    pub(super) fn allocate(&mut self, count: usize) -> FaultResult<Vec<u64>> {
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
    pub(super) fn insert(&mut self, name: &str, object: ArchiveObject) -> FaultResult<()> {
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
    pub(super) fn upsert(&mut self, name: &str, object: ArchiveObject) -> FaultResult<()> {
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
    pub(super) fn remove(&mut self, identifier: u64) -> FaultResult<ArchiveObject> {
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
    pub(super) fn maximum_identifier(&self) -> u64 {
        let source = self
            .source
            .state
            .components
            .iter_objects()
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
    pub(super) fn keep_previews(&mut self) {
        self.keep_previews = true;
    }

    /// Re-encode the touched components, rewrite the metadata registry, drop
    /// the stale root previews and reassemble one candidate.
    ///
    /// `registry` describes the registry transition the caller proved: the
    /// UUID and external-reference records of removed objects, and the
    /// external references new objects need. Returns the reopened candidate,
    /// the number of rewritten components and the number of deleted previews.
    pub(super) fn publish(mut self, registry: &metadata::Transition) -> FaultResult<Published> {
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
        let mut deleted: Vec<&str> = previews.clone();
        deleted.extend(emptied.iter().map(String::as_str));
        deleted.extend(registry.deleted_entries());
        let edits: Vec<EntryEdit<'_>> = encoded
            .iter()
            .map(|(name, data)| EntryEdit::new(name, data))
            .collect();
        let output = self
            .catalog
            .package()
            .reassemble_with_changes_to_bytes(&[], &edits, &deleted, self.catalog.limits())
            .map_err(|error| map_archive(&error))?;
        let bytes: Arc<[u8]> = output.into();
        let candidate = self
            .source
            .reopen_candidate(Arc::clone(&bytes))
            .map_err(|_error| Fault::Verification)?;
        verify_locality(self.catalog, &candidate, &encoded, &deleted)?;
        Ok(Published {
            package: candidate,
            touched_components: encoded.len() + emptied.len(),
            deleted_previews: previews.len(),
        })
    }
}

/// One reassembled, reopened candidate.
pub(super) struct Published {
    pub(super) package: Package,
    pub(super) touched_components: usize,
    pub(super) deleted_previews: usize,
}

/// Every member the transaction did not rewrite or delete keeps its exact
/// decoded bytes, in the same order.
fn verify_locality(
    source: &SourceCatalog,
    candidate: &Package,
    rewritten: &[(String, Vec<u8>)],
    deleted: &[&str],
) -> FaultResult<()> {
    let after = candidate
        .state
        .components
        .physical()
        .ok_or(Fault::Verification)?
        .package();
    let mut retained = source
        .package()
        .iter()
        .filter(|entry| !deleted.contains(&entry.name()));
    for entry in after.iter() {
        let before = retained.next().ok_or(Fault::Verification)?;
        if before.name() != entry.name() {
            return Err(Fault::Verification);
        }
        let changed = rewritten.iter().any(|(name, _)| name == entry.name());
        if !changed && before.data() != entry.data() {
            return Err(Fault::Verification);
        }
    }
    if retained.next().is_some() {
        return Err(Fault::Verification);
    }
    Ok(())
}

/// The exact source/target artifact pair a structural patch retains.
#[derive(Clone, PartialEq, Eq)]
pub(super) struct Artifacts(pub(super) OwnedExactArtifacts);

impl Artifacts {
    /// The pair of `source` and `target`.
    pub(super) fn new(source: &Package, target: &Package) -> FaultResult<Self> {
        Ok(Self(OwnedExactArtifacts::new(
            source_owner(source)?,
            source_owner(target)?,
        )))
    }

    /// The pair of `source` with itself (an exact no-op).
    pub(super) fn unchanged(source: &Package) -> FaultResult<Self> {
        let owner = source_owner(source)?;
        Ok(Self(OwnedExactArtifacts::new(owner.clone(), owner)))
    }

    pub(super) fn is_noop(&self) -> bool {
        self.0.is_byte_noop()
    }

    #[must_use]
    pub(super) fn inverse(&self) -> Self {
        Self(self.0.inverse())
    }

    /// Republish the retained target over `package`, which must be the
    /// retained exact source.
    ///
    /// Returns `None` when `package` is not the patch's source.
    pub(super) fn apply(&self, package: &Package) -> FaultResult<Option<Package>> {
        if !self.0.authorizes_owner(&source_owner(package)?) {
            return Ok(None);
        }
        if self.0.is_byte_noop() {
            return Ok(Some(package.clone()));
        }
        Package::from_source_owner_with_options(self.0.target_owner(), package.state.options)
            .map(Some)
            .map_err(|_error| Fault::Verification)
    }
}

/// The `LimitKind`, `Patch`, `Diagnostics` and `Commit` of one structural
/// family, and the mapping of [`Fault`] onto its `Error` (which must hold
/// the shared variants `UnsupportedSource`, `InvalidSource`,
/// `UnsupportedDependency`, `LimitExceeded`, `Allocation` and
/// `Verification`).
macro_rules! structural_family {
    ($family:literal $(, $wire_variant:ident($wire_kind:ty))?) => {
        /// A finite resource governed by the transaction.
        #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
        #[non_exhaustive]
        pub enum LimitKind {
            /// Complete source or candidate package bytes.
            PackageBytes,
            /// Bytes in one decoded component.
            PayloadBytes,
            /// Native object identifiers.
            Identifiers,
            $(
                /// A bounded wire operation performed by this family.
                $wire_variant($wire_kind),
            )?
        }

        impl ::std::fmt::Display for LimitKind {
            fn fmt(&self, formatter: &mut ::std::fmt::Formatter<'_>) -> ::std::fmt::Result {
                formatter.write_str(match self {
                    Self::PackageBytes => "package bytes",
                    Self::PayloadBytes => "payload bytes",
                    Self::Identifiers => "object identifiers",
                    $(Self::$wire_variant(kind) => return ::std::fmt::Display::fmt(kind, formatter),)?
                })
            }
        }

        impl From<$crate::package::structure::Fault> for Error {
            fn from(fault: $crate::package::structure::Fault) -> Self {
                use $crate::package::structure::{Fault, FaultLimit};
                match fault {
                    Fault::Unsupported => Self::Refused(Refusal::UnsupportedSource),
                    Fault::Invalid => Self::InvalidSource,
                    Fault::Dependency => Self::Refused(Refusal::UnsupportedDependency),
                    Fault::Limit {
                        kind,
                        observed,
                        maximum,
                    } => Self::Limit {
                        kind: match kind {
                            FaultLimit::PackageBytes => LimitKind::PackageBytes,
                            FaultLimit::PayloadBytes => LimitKind::PayloadBytes,
                            FaultLimit::Identifiers => LimitKind::Identifiers,
                        },
                        observed,
                        maximum,
                    },
                    Fault::Allocation { amount } => Self::Allocation { amount },
                    Fault::Verification => Self::Verification,
                }
            }
        }

        #[doc = concat!("A reversible, exact-source ", $family, " patch.")]
        ///
        /// The patch retains the exact source and target artifacts; it
        /// authorizes application only to its source, and its inverse
        /// restores the source byte for byte.
        #[derive(Clone, PartialEq, Eq)]
        pub struct Patch {
            artifacts: $crate::package::structure::Artifacts,
        }

        impl ::std::fmt::Debug for Patch {
            fn fmt(&self, formatter: &mut ::std::fmt::Formatter<'_>) -> ::std::fmt::Result {
                formatter.debug_struct("Patch").finish_non_exhaustive()
            }
        }

        impl Patch {
            /// Whether the patch is an exact artifact no-op.
            #[must_use]
            pub fn is_noop(&self) -> bool {
                self.artifacts.is_noop()
            }

            /// The exact target-to-source inverse patch.
            #[must_use]
            pub fn inverse(&self) -> Self {
                Self {
                    artifacts: self.artifacts.inverse(),
                }
            }
        }

        #[doc = concat!("Content-free diagnostics for one ", $family, " publication.")]
        #[derive(Debug, Clone, Copy, PartialEq, Eq)]
        pub struct Diagnostics {
            changed: bool,
            touched_components: usize,
            deleted_previews: usize,
        }

        impl Diagnostics {
            const fn unchanged() -> Self {
                Self {
                    changed: false,
                    touched_components: 0,
                    deleted_previews: 0,
                }
            }

            const fn published(published: &$crate::package::structure::Published) -> Self {
                Self {
                    changed: true,
                    touched_components: published.touched_components,
                    deleted_previews: published.deleted_previews,
                }
            }

            /// Whether the exact package artifact changed.
            #[must_use]
            pub const fn changed(self) -> bool {
                self.changed
            }

            /// Number of rewritten native components.
            #[must_use]
            pub const fn touched_components(self) -> usize {
                self.touched_components
            }

            /// Number of stale root previews deleted.
            #[must_use]
            pub const fn deleted_previews(self) -> usize {
                self.deleted_previews
            }
        }

        #[doc = concat!("One validated ", $family, " publication.")]
        #[derive(Debug)]
        pub struct Commit {
            package: $crate::Package,
            patch: Patch,
            diagnostics: Diagnostics,
        }

        impl Commit {
            fn new(
                source: &$crate::Package,
                published: $crate::package::structure::Published,
            ) -> Result<Self, Error> {
                Ok(Self {
                    patch: Patch {
                        artifacts: $crate::package::structure::Artifacts::new(
                            source,
                            &published.package,
                        )?,
                    },
                    diagnostics: Diagnostics::published(&published),
                    package: published.package,
                })
            }

            #[allow(dead_code, reason = "only the families with an exact no-op use it")]
            fn unchanged(source: &$crate::Package) -> Result<Self, Error> {
                Ok(Self {
                    patch: Patch {
                        artifacts: $crate::package::structure::Artifacts::unchanged(source)?,
                    },
                    diagnostics: Diagnostics::unchanged(),
                    package: source.clone(),
                })
            }

            fn applied(package: $crate::Package, patch: &Patch) -> Self {
                Self {
                    diagnostics: if patch.is_noop() {
                        Diagnostics::unchanged()
                    } else {
                        Diagnostics {
                            changed: true,
                            touched_components: 0,
                            deleted_previews: 0,
                        }
                    },
                    patch: patch.clone(),
                    package,
                }
            }

            /// Borrow the validated package snapshot.
            #[must_use]
            pub const fn package(&self) -> &$crate::Package {
                &self.package
            }

            /// Consume this publication and return its package snapshot.
            #[must_use]
            pub fn into_package(self) -> $crate::Package {
                self.package
            }

            /// Consume this publication and return its package and patch.
            #[must_use]
            pub fn into_parts(self) -> ($crate::Package, Patch) {
                (self.package, self.patch)
            }

            /// Borrow the reversible exact-source patch.
            #[must_use]
            pub const fn patch(&self) -> &Patch {
                &self.patch
            }

            /// Borrow content-free diagnostics.
            #[must_use]
            pub const fn diagnostics(&self) -> &Diagnostics {
                &self.diagnostics
            }
        }

        fn apply_patch(package: &$crate::Package, patch: &Patch) -> Result<Commit, Error> {
            let target = patch
                .artifacts
                .apply(package)?
                .ok_or(Error::PatchConflict)?;
            Ok(Commit::applied(target, patch))
        }
    };
}

pub(super) use structural_family;
