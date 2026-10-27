//! Shared exact-source machinery of the structural Pages transactions
//! (record 0312): [`crate::section::lifecycle`].
//!
//! The toolkit follows `litchi-numbers`' structural machinery (record 0301):
//! a structural edit stages its changes on a [`Working`] set. Every
//! component it touches is decompressed and parsed once from the exact
//! source, canonical framing is checked, and the edit mutates the parsed
//! objects in place; objects the edit does not touch keep their exact
//! encoded bytes, and members it does not touch keep their exact compressed
//! bytes. [`Working::publish`] re-encodes only the touched components,
//! applies the package metadata registry transition (UUID bindings,
//! external references and the issued-identifier watermark) and reassembles
//! the ZIP once. The candidate is reopened through the ordinary strict
//! ingress, and every member the edit did not rewrite is proven unchanged.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the working set keeps its helpers beside the phase that uses them"
)]

use std::collections::{BTreeMap, BTreeSet, HashMap};
use std::sync::Arc;

use litchi_iwa_archive::package::{EntryEdit, OwnedExactArtifacts, SharedBytes};
use litchi_iwa_common::wire::{
    parse_wire_view, repeated_length_delimited_payloads, repeated_varint_values,
    transform_length_delimited_fields_at_path,
};
use litchi_iwa_core::archive::{FieldObjectReferenceTransition, ObjectReferenceTransition};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::package_metadata_codec::DataReferenceOwnerDescriptor;
use litchi_iwa_protos::package_metadata_codec::{
    Batch, ComponentDescriptor, ComponentSelector, ExternalReferenceAddition,
    ExternalReferenceDescriptor, ExternalReferenceRemoval, ObjectUuidAddition,
    ObjectUuidDescriptor, ObjectUuidRemoval, PackageMetadataVisitor, RemovalBatch,
    RemovalSaveTokenBatch, RewriteError, RewriteLimit, RewriteOptions, SaveTokenBatch, UuidBits,
    inspect_package_metadata_with_visitor, rewrite_package_metadata,
    rewrite_package_metadata_removals_and_save_tokens,
};

use super::Package;

/// The package metadata registry member.
const METADATA_ENTRY: &str = "Index/Metadata.iwa";
/// `TSP.PackageMetadata`.
const METADATA_MESSAGE_TYPE: u32 = 11_006;
/// `TSP.Reference.identifier`.
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;

/// A content-free failure of a structural transaction, mapped by each
/// family onto its public `Error`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum Fault {
    /// The source cannot publish an exact changed artifact, or holds a
    /// native shape the transaction does not own.
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

pub(super) const fn allocation(amount: usize) -> Fault {
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

/// A [`Fault`] carried through the common wire transform callbacks.
pub(super) struct WireFault(pub(super) Fault);

impl From<litchi_iwa_common::Error> for WireFault {
    fn from(error: litchi_iwa_common::Error) -> Self {
        Self(map_wire(&error))
    }
}

// ---------------------------------------------------------------------------
// References
// ---------------------------------------------------------------------------

/// Encode a canonical `TSP.Reference` payload for `identifier`.
pub(super) fn reference_payload(identifier: u64) -> FaultResult<Vec<u8>> {
    let mut output = Vec::new();
    reserve(&mut output, 11)?;
    output.push(0x08);
    let mut value = identifier;
    loop {
        let byte = u8::try_from(value & 0x7f).map_err(|_error| Fault::Invalid)?;
        value >>= 7;
        if value == 0 {
            output.push(byte);
            break;
        }
        output.push(byte | 0x80);
    }
    Ok(output)
}

/// The identifier of a `TSP.Reference` payload. Records carrying anything
/// but their identifier are refused: a structural edit never reinterprets
/// them.
pub(super) fn reference_identifier(payload: &[u8]) -> FaultResult<u64> {
    let view = parse_wire_view(payload, litchi_iwa_common::WireLimits::DEFAULT)
        .map_err(|error| map_wire(&error))?;
    if view
        .fields()
        .any(|field| field.number() != REFERENCE_IDENTIFIER_FIELD || field.wire_type() != 0)
    {
        return Err(Fault::Unsupported);
    }
    let values = repeated_varint_values(payload, REFERENCE_IDENTIFIER_FIELD)
        .map_err(|error| map_wire(&error))?;
    match values.as_slice() {
        [identifier] if *identifier != 0 => Ok(*identifier),
        _ => Err(Fault::Invalid),
    }
}

/// The identifiers of the (repeated) reference field `field` of `data`.
pub(super) fn identifiers(data: &[u8], field: u32) -> FaultResult<Vec<u64>> {
    let records =
        repeated_length_delimited_payloads(data, field).map_err(|error| map_wire(&error))?;
    let mut output = Vec::new();
    reserve(&mut output, records.len())?;
    for record in records {
        output.push(reference_identifier(record)?);
    }
    Ok(output)
}

/// Remap every reference at `path` of `data` through `remap`; references
/// outside the map keep their exact bytes.
pub(super) fn remap_at(
    data: &[u8],
    path: &[u32],
    remap: &HashMap<u64, u64>,
) -> FaultResult<Vec<u8>> {
    transform_length_delimited_fields_at_path(data, path, |reference| {
        let identifier = reference_identifier(reference).map_err(WireFault)?;
        Ok::<_, WireFault>(match remap.get(&identifier) {
            Some(replacement) => reference_payload(*replacement).map_err(WireFault)?,
            None => copy_bytes(reference).map_err(WireFault)?,
        })
    })
    .map_err(|error| error.0)
}

/// Replace message `index` of `object` with `data`, transitioning its
/// header references: `removed` leave the aggregate list and every field
/// list, `appended` join the aggregate list as its suffix.
pub(super) fn replace(
    object: &mut ArchiveObject,
    index: usize,
    data: Vec<u8>,
    removed: &[u64],
    appended: &[u64],
    limits: litchi_iwa_core::Limits,
) -> FaultResult<()> {
    let info = object
        .archive_info
        .message_infos
        .get(index)
        .ok_or(Fault::Invalid)?;
    let message_type = object.messages.get(index).ok_or(Fault::Invalid)?.type_;
    let aggregate_before = info.object_references.clone();
    let mut aggregate_after = Vec::new();
    reserve(
        &mut aggregate_after,
        aggregate_before.len() + appended.len(),
    )?;
    aggregate_after.extend(
        aggregate_before
            .iter()
            .copied()
            .filter(|identifier| !removed.contains(identifier)),
    );
    for identifier in appended {
        if !aggregate_after.contains(identifier) {
            aggregate_after.push(*identifier);
        }
    }
    let mut owned = Vec::new();
    for (field_info_index, field) in info.field_infos.iter().enumerate() {
        if !field
            .object_references
            .iter()
            .any(|identifier| removed.contains(identifier))
        {
            continue;
        }
        let before = field.object_references.clone();
        let after: Vec<u64> = before
            .iter()
            .copied()
            .filter(|identifier| !removed.contains(identifier))
            .collect();
        reserve(&mut owned, 1)?;
        owned.push((
            field_info_index,
            field.path.as_slice().to_vec(),
            before,
            after,
        ));
    }
    let fields: Vec<FieldObjectReferenceTransition<'_>> = owned
        .iter()
        .map(
            |(field_info_index, path, before, after)| FieldObjectReferenceTransition {
                field_info_index: *field_info_index,
                expected_path: path,
                before,
                after,
            },
        )
        .collect();
    object
        .replace_message_transitioning_object_references_preserving_header(
            index,
            RawMessage {
                type_: message_type,
                data,
            },
            ObjectReferenceTransition {
                aggregate_before: &aggregate_before,
                aggregate_after: &aggregate_after,
                fields: &fields,
            },
            limits,
        )
        .map_err(|error| map_core(&error))?;
    Ok(())
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

// ---------------------------------------------------------------------------
// Census
// ---------------------------------------------------------------------------

/// The source's reference census: for each object, the objects it
/// references and the objects that reference it, read from the archive
/// headers' aggregate reference lists (no payload is reinterpreted).
pub(super) struct Census {
    outbound: HashMap<u64, Vec<u64>>,
    inbound: HashMap<u64, Vec<u64>>,
    /// Objects that never join a private closure: the document root and the
    /// registry.
    protected: Vec<u64>,
}

impl Census {
    /// Build the census of every object of `source`.
    pub(super) fn new(source: &Package) -> FaultResult<Self> {
        let mut outbound: HashMap<u64, Vec<u64>> = HashMap::new();
        let mut inbound: HashMap<u64, Vec<u64>> = HashMap::new();
        let mut protected = vec![1_u64];
        for component in source.state.source.components().iter() {
            for object in &component.archive().objects {
                let identifier = object.archive_info.identifier.ok_or(Fault::Invalid)?;
                if component.name() == METADATA_ENTRY {
                    reserve(&mut protected, 1)?;
                    protected.push(identifier);
                }
                let mut targets: Vec<u64> = Vec::new();
                for info in &object.archive_info.message_infos {
                    for target in &info.object_references {
                        if !targets.contains(target) {
                            reserve(&mut targets, 1)?;
                            targets.push(*target);
                        }
                    }
                }
                for target in &targets {
                    let referrers = inbound.entry(*target).or_default();
                    if !referrers.contains(&identifier) {
                        reserve(referrers, 1)?;
                        referrers.push(identifier);
                    }
                }
                outbound.try_reserve(1).map_err(|_error| allocation(1))?;
                outbound.insert(identifier, targets);
            }
        }
        Ok(Self {
            outbound,
            inbound,
            protected,
        })
    }

    /// The objects that reference `identifier`.
    pub(super) fn referrers(&self, identifier: u64) -> &[u64] {
        self.inbound.get(&identifier).map_or(&[], Vec::as_slice)
    }

    /// The objects `identifier` references.
    pub(super) fn targets(&self, identifier: u64) -> &[u64] {
        self.outbound.get(&identifier).map_or(&[], Vec::as_slice)
    }

    /// The private closure of `roots`: the roots and every object all of
    /// whose referrers are in the closure (shared records, such as styles,
    /// stay outside).
    pub(super) fn private_closure(&self, roots: &[u64]) -> FaultResult<BTreeSet<u64>> {
        let mut closure: BTreeSet<u64> = roots.iter().copied().collect();
        loop {
            let mut added = Vec::new();
            for member in &closure {
                for target in self.targets(*member) {
                    if closure.contains(target)
                        || added.contains(target)
                        || self.protected.contains(target)
                        || !self.outbound.contains_key(target)
                    {
                        continue;
                    }
                    if self
                        .referrers(*target)
                        .iter()
                        .all(|referrer| closure.contains(referrer))
                    {
                        reserve(&mut added, 1)?;
                        added.push(*target);
                    }
                }
            }
            if added.is_empty() {
                return Ok(closure);
            }
            closure.extend(added);
        }
    }

    /// Refuse when an object outside `closure` still references a member,
    /// other than through one of the `detached` (referrer, target) edges.
    pub(super) fn require_detached(
        &self,
        closure: &BTreeSet<u64>,
        detached: &[(u64, u64)],
    ) -> FaultResult<()> {
        for member in closure {
            for referrer in self.referrers(*member) {
                if closure.contains(referrer) || detached.contains(&(*referrer, *member)) {
                    continue;
                }
                return Err(Fault::Dependency);
            }
        }
        Ok(())
    }
}

/// Prove that no object of `candidate` is a removed identifier or still
/// references one.
pub(super) fn verify_removed(candidate: &Package, removed: &BTreeSet<u64>) -> FaultResult<()> {
    for component in candidate.state.source.components().iter() {
        for object in &component.archive().objects {
            if object
                .archive_info
                .identifier
                .is_some_and(|identifier| removed.contains(&identifier))
            {
                return Err(Fault::Verification);
            }
            for info in &object.archive_info.message_infos {
                let stale = info
                    .object_references
                    .iter()
                    .chain(
                        info.field_infos
                            .iter()
                            .flat_map(|field| field.object_references.iter()),
                    )
                    .any(|identifier| removed.contains(identifier));
                if stale {
                    return Err(Fault::Verification);
                }
            }
        }
    }
    Ok(())
}

// ---------------------------------------------------------------------------
// Registry
// ---------------------------------------------------------------------------

/// One current component registration.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct Component {
    pub(super) identifier: u64,
    pub(super) locator: String,
}

/// One object UUID binding.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct Uuid {
    pub(super) component: u64,
    pub(super) object: u64,
    pub(super) uuid: UuidBits,
}

/// One cross-component object reference record.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct External {
    pub(super) source: u64,
    pub(super) target: u64,
    pub(super) object: u64,
    pub(super) weak: Option<bool>,
}

/// The registry facts of one metadata payload.
#[derive(Debug, Default)]
pub(super) struct Registry {
    pub(super) components: Vec<Component>,
    pub(super) uuids: Vec<Uuid>,
    pub(super) externals: Vec<External>,
    /// Objects that own a data (media) reference.
    pub(super) data_owners: Vec<u64>,
    /// Component-level external references (no object named): source and
    /// target components.
    pub(super) component_externals: Vec<(u64, u64)>,
    pub(super) last_object_identifier: u64,
    unknown: bool,
}

impl Registry {
    fn component(&self, identifier: u64) -> FaultResult<&Component> {
        self.components
            .iter()
            .find(|component| component.identifier == identifier)
            .ok_or(Fault::Invalid)
    }

    /// The component stored at package member `member`.
    pub(super) fn component_for_member(&self, member: &str) -> FaultResult<&Component> {
        self.components
            .iter()
            .find(|component| {
                member
                    .strip_prefix("Index/")
                    .and_then(|rest| rest.strip_suffix(".iwa"))
                    == Some(component.locator.as_str())
            })
            .ok_or(Fault::Unsupported)
    }

    /// Whether component `source` already records a reference to `object`
    /// of component `target` (by object, or at component level when the
    /// object is the target component's own).
    pub(super) fn covers(&self, source: u64, target: u64, object: u64) -> bool {
        self.externals.iter().any(|external| {
            external.source == source && external.target == target && external.object == object
        }) || (target == object && self.component_externals.contains(&(source, target)))
    }
}

#[derive(Default)]
struct Visitor {
    registry: Registry,
}

fn remember<T>(items: &mut Vec<T>, item: T) -> Result<(), RewriteError> {
    items
        .try_reserve(1)
        .map_err(|_error| RewriteError::allocation(1))?;
    items.push(item);
    Ok(())
}

impl PackageMetadataVisitor for Visitor {
    fn visit_unknown_field(&mut self) -> Result<(), RewriteError> {
        self.registry.unknown = true;
        Ok(())
    }

    fn visit_component(&mut self, component: ComponentDescriptor<'_>) -> Result<(), RewriteError> {
        if !component.is_current() {
            return Ok(());
        }
        remember(
            &mut self.registry.components,
            Component {
                identifier: component.identifier(),
                locator: component.effective_locator().to_owned(),
            },
        )
    }

    fn visit_object_uuid(&mut self, binding: ObjectUuidDescriptor<'_>) -> Result<(), RewriteError> {
        remember(
            &mut self.registry.uuids,
            Uuid {
                component: binding.component().identifier(),
                object: binding.object_identifier(),
                uuid: binding.uuid(),
            },
        )
    }

    fn visit_external_reference(
        &mut self,
        reference: ExternalReferenceDescriptor<'_>,
    ) -> Result<(), RewriteError> {
        if let Some(object) = reference.object_identifier() {
            remember(
                &mut self.registry.externals,
                External {
                    source: reference.source().identifier(),
                    target: reference.target_component_identifier(),
                    object,
                    weak: reference.is_weak(),
                },
            )?;
        } else {
            remember(
                &mut self.registry.component_externals,
                (
                    reference.source().identifier(),
                    reference.target_component_identifier(),
                ),
            )?;
        }
        Ok(())
    }

    fn visit_data_reference_owner(
        &mut self,
        owner: DataReferenceOwnerDescriptor<'_>,
    ) -> Result<(), RewriteError> {
        remember(&mut self.registry.data_owners, owner.object_identifier())
    }
}

fn map_rewrite(error: &RewriteError) -> Fault {
    if let Some(amount) = error.allocation_request() {
        return allocation(amount);
    }
    let figures = |observed: usize, maximum: usize| Fault::Limit {
        kind: FaultLimit::PayloadBytes,
        observed: u64::try_from(observed).unwrap_or(u64::MAX),
        maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
    };
    match error.resource_limit() {
        Some(
            RewriteLimit::InputBytes { observed, maximum }
            | RewriteLimit::OutputBytes { observed, maximum }
            | RewriteLimit::Fields { observed, maximum }
            | RewriteLimit::Work { observed, maximum }
            | RewriteLimit::Components { observed, maximum }
            | RewriteLimit::References { observed, maximum }
            | RewriteLimit::Additions { observed, maximum },
        ) => figures(observed, maximum),
        Some(RewriteLimit::Nesting { observed, maximum }) => {
            figures(observed as usize, maximum as usize)
        },
        Some(_) => figures(0, 0),
        None => Fault::Invalid,
    }
}

/// The registry transition of one transaction.
#[derive(Debug, Default)]
pub(super) struct Transition {
    pub(super) uuid_additions: Vec<Uuid>,
    pub(super) uuid_removals: Vec<Uuid>,
    pub(super) external_removals: Vec<External>,
    /// Cross-component object references the transaction introduces.
    pub(super) external_additions: Vec<External>,
}

impl Transition {
    fn is_empty(&self) -> bool {
        self.uuid_additions.is_empty()
            && self.uuid_removals.is_empty()
            && self.external_removals.is_empty()
            && self.external_additions.is_empty()
    }

    /// Record the registry consequences of removing `objects`: their UUID
    /// bindings and every external reference that addresses them. Media
    /// ownership is refused (its data member would need collecting too).
    pub(super) fn remove_objects(
        &mut self,
        registry: &Registry,
        objects: &[u64],
    ) -> FaultResult<()> {
        for binding in &registry.uuids {
            if objects.contains(&binding.object) {
                reserve(&mut self.uuid_removals, 1)?;
                self.uuid_removals.push(*binding);
            }
        }
        for external in &registry.externals {
            if objects.contains(&external.object) {
                reserve(&mut self.external_removals, 1)?;
                self.external_removals.push(*external);
            }
        }
        if registry
            .data_owners
            .iter()
            .any(|owner| objects.contains(owner))
        {
            return Err(Fault::Unsupported);
        }
        Ok(())
    }

    /// Register a fresh UUID for new object `object` of component
    /// `component`, derived from `seed` and distinct from every registered
    /// UUID (no clock or randomness is consulted).
    pub(super) fn register_uuid(
        &mut self,
        registry: &Registry,
        component: u64,
        object: u64,
        seed: u64,
    ) -> FaultResult<()> {
        let mut state = seed ^ object.rotate_left(17);
        loop {
            let lower = mix(state);
            let upper = mix(lower ^ 0x6c69_7463_6869_0312);
            let uuid = UuidBits::new(lower, upper);
            let taken = registry.uuids.iter().any(|binding| binding.uuid == uuid)
                || self
                    .uuid_additions
                    .iter()
                    .any(|binding| binding.uuid == uuid);
            if !taken {
                reserve(&mut self.uuid_additions, 1)?;
                self.uuid_additions.push(Uuid {
                    component,
                    object,
                    uuid,
                });
                return Ok(());
            }
            state = state.wrapping_add(0x9e37_79b9_7f4a_7c15);
        }
    }
}

/// A 64-bit finalizer (SplitMix64) for deterministic identity derivation.
pub(super) const fn mix(mut value: u64) -> u64 {
    value = value.wrapping_add(0x9e37_79b9_7f4a_7c15);
    value = (value ^ (value >> 30)).wrapping_mul(0xbf58_476d_1ce4_e5b9);
    value = (value ^ (value >> 27)).wrapping_mul(0x94d0_49bb_1331_11eb);
    value ^ (value >> 31)
}

// ---------------------------------------------------------------------------
// Working set
// ---------------------------------------------------------------------------

/// The staged state of one structural transaction over an exact source.
pub(super) struct Working<'p> {
    source: &'p Package,
    archive_limits: litchi_iwa_core::Limits,
    edited: BTreeMap<String, Archive>,
    /// Objects added by this transaction and the component holding them.
    inserted: HashMap<u64, String>,
    /// Source objects removed by this transaction.
    removed: BTreeSet<u64>,
    /// The highest identifier this transaction allocated.
    allocated: Option<u64>,
    /// The next identifier to allocate.
    next_identifier: u64,
    /// The metadata object and message, when the package has a registry.
    route: Option<(u64, usize)>,
}

impl<'p> Working<'p> {
    /// Start a working set over `source`'s exact artifact.
    pub(super) fn new(source: &'p Package) -> FaultResult<Self> {
        let catalog = &source.state.source;
        let archive_limits = catalog
            .limits()
            .effective_archive_limits()
            .map_err(|error| map_archive(&error))?;
        let mut maximum = 0_u64;
        let mut route = None;
        for component in catalog.components().iter() {
            for object in &component.archive().objects {
                let identifier = object.archive_info.identifier.ok_or(Fault::Invalid)?;
                maximum = maximum.max(identifier);
                if component.name() == METADATA_ENTRY {
                    for (index, message) in object.messages.iter().enumerate() {
                        if message.type_ == METADATA_MESSAGE_TYPE
                            && route.replace((identifier, index)).is_some()
                        {
                            return Err(Fault::Invalid);
                        }
                    }
                }
            }
        }
        let mut working = Self {
            source,
            archive_limits,
            edited: BTreeMap::new(),
            inserted: HashMap::new(),
            removed: BTreeSet::new(),
            allocated: None,
            next_identifier: 0,
            route,
        };
        let registered = working
            .registry()?
            .map_or(0, |registry| registry.last_object_identifier);
        working.next_identifier = maximum.max(registered).checked_add(1).ok_or(Fault::Limit {
            kind: FaultLimit::Identifiers,
            observed: u64::MAX,
            maximum: u64::MAX,
        })?;
        Ok(working)
    }

    /// The archive limits every touched component is parsed and encoded
    /// under.
    pub(super) const fn archive_limits(&self) -> litchi_iwa_core::Limits {
        self.archive_limits
    }

    /// The current archive of component `name`.
    fn archive(&self, name: &str) -> FaultResult<&Archive> {
        if let Some(archive) = self.edited.get(name) {
            return Ok(archive);
        }
        self.source
            .state
            .source
            .components()
            .get(name)
            .map(litchi_iwa_archive::Component::archive)
            .ok_or(Fault::Invalid)
    }

    /// The component holding object `identifier`, if it exists. An
    /// identifier stored in more than one component is refused.
    pub(super) fn locate(&self, identifier: u64) -> FaultResult<Option<String>> {
        if self.removed.contains(&identifier) {
            return Ok(None);
        }
        if let Some(name) = self.inserted.get(&identifier) {
            return Ok(Some(name.clone()));
        }
        let mut found: Option<&str> = None;
        for component in self.source.state.source.components().iter() {
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

    /// The archive of component `name`, parsed from its exact compressed
    /// member for editing.
    fn archive_mut(&mut self, name: &str) -> FaultResult<&mut Archive> {
        if !self.edited.contains_key(name) {
            let catalog = &self.source.state.source;
            let entry = catalog
                .package()
                .iter()
                .find(|entry| entry.name() == name)
                .ok_or(Fault::Invalid)?;
            if entry.is_opaque() {
                return Err(Fault::Unsupported);
            }
            let snappy = catalog
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
            self.edited.insert(name.to_owned(), archive);
        }
        self.edited.get_mut(name).ok_or(Fault::Invalid)
    }

    /// The object `identifier`, parsed for editing.
    pub(super) fn object_mut(&mut self, identifier: u64) -> FaultResult<&mut ArchiveObject> {
        let name = self.locate(identifier)?.ok_or(Fault::Invalid)?;
        self.archive_mut(&name)?
            .object_mut(identifier)
            .ok_or(Fault::Invalid)
    }

    /// The identifier the next [`Self::allocate`] returns first.
    pub(super) const fn next_identifier(&self) -> u64 {
        self.next_identifier
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

    fn metadata_payload(&self) -> FaultResult<Option<&[u8]>> {
        let Some((object, message)) = self.route else {
            return Ok(None);
        };
        Ok(Some(
            self.archive(METADATA_ENTRY)?
                .object(object)
                .and_then(|object| object.messages.get(message))
                .map(|message| message.data.as_slice())
                .ok_or(Fault::Invalid)?,
        ))
    }

    fn options(&self, payload: usize, additions: usize) -> RewriteOptions {
        let components = self
            .source
            .state
            .source
            .components()
            .len()
            .max(payload)
            .max(1);
        RewriteOptions::new(
            payload.max(1),
            payload.saturating_mul(2).saturating_add(4_096),
            payload
                .saturating_mul(4)
                .saturating_add(additions.saturating_mul(16))
                .max(64),
            payload
                .saturating_mul(256)
                .saturating_add(
                    components
                        .saturating_mul(additions.max(1))
                        .saturating_mul(64),
                )
                .clamp(1, litchi_iwa_common::WireLimits::MAX_REWRITE_WORK),
            64,
            components.saturating_add(additions),
            payload.max(1),
            additions.max(1),
        )
    }

    /// The registry of the package, or `None` without `Index/Metadata.iwa`.
    pub(super) fn registry(&self) -> FaultResult<Option<Registry>> {
        let Some(payload) = self.metadata_payload()? else {
            return Ok(None);
        };
        let mut visitor = Visitor::default();
        let inspection = inspect_package_metadata_with_visitor(
            payload,
            self.options(payload.len(), 0),
            &mut visitor,
        )
        .map_err(|error| map_rewrite(&error))?;
        visitor.registry.last_object_identifier = inspection.last_object_identifier();
        if visitor.registry.unknown {
            return Err(Fault::Unsupported);
        }
        Ok(Some(visitor.registry))
    }

    /// A deterministic seed for the fresh identities of one transaction: the
    /// registry payload (which carries the document's own identity) and the
    /// first identifier the transaction allocates.
    pub(super) fn identity_seed(&self, first: u64) -> FaultResult<u64> {
        let mut seed = mix(first);
        if let Some(payload) = self.metadata_payload()? {
            for chunk in payload.chunks(8) {
                let mut word = [0_u8; 8];
                word[..chunk.len()].copy_from_slice(chunk);
                seed = mix(seed ^ u64::from_le_bytes(word));
            }
        }
        Ok(seed)
    }

    /// Apply `transition` and raise the watermark to the allocated
    /// identifiers.
    fn apply_registry(&mut self, transition: &Transition) -> FaultResult<()> {
        let Some(route) = self.route else {
            if transition.is_empty() && self.allocated.is_none() {
                return Ok(());
            }
            return Err(Fault::Unsupported);
        };
        let registry = self.registry()?.ok_or(Fault::Invalid)?;
        let current = registry.last_object_identifier;
        let last = self
            .allocated
            .map_or(current, |allocated| allocated.max(current));
        if transition.is_empty() && last == current {
            return Ok(());
        }
        let source = self.metadata_payload()?.ok_or(Fault::Invalid)?;
        let additions = transition
            .uuid_removals
            .len()
            .saturating_add(transition.uuid_additions.len())
            .saturating_add(transition.external_removals.len())
            .saturating_add(transition.external_additions.len());
        let options = self.options(source.len(), additions);
        let selector = |identifier: u64| -> FaultResult<ComponentSelector<'_>> {
            let component = registry.component(identifier)?;
            Ok(ComponentSelector::new(
                component.identifier,
                component.locator.as_str(),
            ))
        };
        let removes =
            !transition.uuid_removals.is_empty() || !transition.external_removals.is_empty();
        let output = if removes {
            if !transition.uuid_additions.is_empty()
                || !transition.external_additions.is_empty()
                || last != current
            {
                // No transaction both adds and removes objects.
                return Err(Fault::Unsupported);
            }
            let mut uuid_removals = Vec::new();
            reserve(&mut uuid_removals, transition.uuid_removals.len())?;
            for binding in &transition.uuid_removals {
                uuid_removals.push(ObjectUuidRemoval::new(
                    selector(binding.component)?,
                    binding.object,
                    binding.uuid,
                ));
            }
            let mut external_removals = Vec::new();
            reserve(&mut external_removals, transition.external_removals.len())?;
            for external in &transition.external_removals {
                external_removals.push(ExternalReferenceRemoval::new(
                    selector(external.source)?,
                    selector(external.target)?,
                    external.object,
                    external.weak,
                ));
            }
            // Every component whose registry record changes advances its
            // save token, as the native app does when it rewrites one.
            let mut tokens: Vec<ComponentSelector<'_>> = Vec::new();
            let touched = transition
                .uuid_removals
                .iter()
                .map(|binding| binding.component)
                .chain(
                    transition
                        .external_removals
                        .iter()
                        .flat_map(|external| [external.source, external.target]),
                );
            for identifier in touched {
                let selector = selector(identifier)?;
                if !tokens.contains(&selector) {
                    reserve(&mut tokens, 1)?;
                    tokens.push(selector);
                }
            }
            rewrite_package_metadata_removals_and_save_tokens(
                source,
                RemovalSaveTokenBatch::new(
                    RemovalBatch::new(current, &uuid_removals, &external_removals, &[]),
                    SaveTokenBatch::new(&tokens),
                ),
                options,
            )
        } else {
            let mut uuid_additions = Vec::new();
            reserve(&mut uuid_additions, transition.uuid_additions.len())?;
            for binding in &transition.uuid_additions {
                uuid_additions.push(ObjectUuidAddition::new(
                    selector(binding.component)?,
                    binding.object,
                    binding.uuid,
                ));
            }
            let mut external_additions = Vec::new();
            reserve(&mut external_additions, transition.external_additions.len())?;
            for external in &transition.external_additions {
                external_additions.push(ExternalReferenceAddition::new(
                    selector(external.source)?,
                    selector(external.target)?,
                    external.object,
                    external.weak,
                ));
            }
            rewrite_package_metadata(
                source,
                Batch::new(current, last, &uuid_additions, &external_additions),
                options,
            )
        }
        .map_err(|error| map_rewrite(&error))?;
        let data = output.into_bytes();
        let limits = self.archive_limits;
        self.archive_mut(METADATA_ENTRY)?
            .object_mut(route.0)
            .ok_or(Fault::Invalid)?
            .replace_message_preserving_header(
                route.1,
                RawMessage {
                    type_: METADATA_MESSAGE_TYPE,
                    data,
                },
                limits,
            )
            .map_err(|error| map_core(&error))?;
        Ok(())
    }

    /// Re-encode the touched components, apply the registry `transition`
    /// and reassemble one candidate; returns the reopened candidate and the
    /// number of rewritten components.
    pub(super) fn publish(mut self, transition: &Transition) -> FaultResult<(Package, usize)> {
        self.apply_registry(transition)?;
        let catalog = &self.source.state.source;
        let snappy_limits = catalog
            .limits()
            .snappy_limits()
            .map_err(|error| map_archive(&error))?;
        let mut encoded: Vec<(String, Vec<u8>)> = Vec::new();
        for (name, archive) in &self.edited {
            let bytes = archive
                .to_bytes(self.archive_limits)
                .map_err(|error| map_core(&error))?;
            let compressed =
                SnappyStream::compress(&bytes, snappy_limits).map_err(|error| map_core(&error))?;
            encoded.push((name.clone(), compressed));
        }
        let edits: Vec<EntryEdit<'_>> = encoded
            .iter()
            .map(|(name, data)| EntryEdit::new(name, data))
            .collect();
        let bytes = catalog
            .package()
            .reassemble_with_changes_to_bytes(&[], &edits, &[], catalog.limits())
            .map_err(|error| map_archive(&error))?;
        let candidate = Package::from_bytes_with_archive_limits(&bytes, catalog.limits())
            .map_err(|_error| Fault::Verification)?;
        // Every member the transaction did not rewrite keeps its bytes.
        let after = candidate.state.source.package();
        if after.iter().count() != catalog.package().iter().count() {
            return Err(Fault::Verification);
        }
        for (before, now) in catalog.package().iter().zip(after.iter()) {
            let rewritten = encoded.iter().any(|(name, _)| name == before.name());
            if before.name() != now.name() || (!rewritten && before.data() != now.data()) {
                return Err(Fault::Verification);
            }
        }
        Ok((candidate, encoded.len()))
    }
}

/// The exact source/target artifact pair a structural patch retains.
#[derive(Clone, PartialEq, Eq)]
pub(super) struct Artifacts(OwnedExactArtifacts);

impl Artifacts {
    /// The pair of `source` and `target`.
    pub(super) fn new(source: &Package, target: &Package) -> Self {
        Self(OwnedExactArtifacts::new(
            SharedBytes::from_shared_slice(source.state.source.shared_source()),
            SharedBytes::from_shared_slice(target.state.source.shared_source()),
        ))
    }

    pub(super) fn is_noop(&self) -> bool {
        self.0.is_byte_noop()
    }

    #[must_use]
    pub(super) fn inverse(&self) -> Self {
        Self(self.0.inverse())
    }

    /// Republish the retained target over `package`, which must be the
    /// retained exact source; `None` when it is not.
    pub(super) fn apply(&self, package: &Package) -> FaultResult<Option<Package>> {
        let owner = SharedBytes::from_shared_slice(package.state.source.shared_source());
        if !self.0.authorizes_owner(&owner) {
            return Ok(None);
        }
        if self.0.is_byte_noop() {
            return Ok(Some(package.clone()));
        }
        let target = Arc::<[u8]>::from(self.0.target_owner().as_ref());
        Package::from_bytes_with_archive_limits(&target, package.state.source.limits())
            .map(Some)
            .map_err(|_error| Fault::Verification)
    }
}
