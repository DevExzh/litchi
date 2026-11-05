//! The package metadata registry side of a structural transaction.
//!
//! `Index/Metadata.iwa` registers every component, the UUIDs of objects
//! other components address by UUID, and every cross-component object
//! reference. A structural edit that adds or removes objects must keep those
//! records exact: removed objects lose their UUID bindings and the external
//! references that pointed at them, new objects that address another
//! component gain an external reference, and the issued-identifier watermark
//! covers every new identifier. The transition is applied through the
//! shared strict codec, which preserves every unselected raw record.

use litchi_iwa_core::RawMessage;
use litchi_iwa_protos::package_metadata_codec::{
    AdditionSaveTokenBatch, Batch, CombinedBatch, CombinedSaveTokenBatch, ComponentDescriptor,
    ComponentRemovalBatch, ComponentSelector, DataReferenceOwnerDescriptor,
    ExternalReferenceAddition, ExternalReferenceDescriptor, ExternalReferenceRemoval,
    ObjectUuidAddition, ObjectUuidDescriptor, ObjectUuidRemoval, PackageMetadataVisitor,
    RemovalBatch, RemovalSaveTokenBatch, RewriteError, RewriteLimit, RewriteOptions,
    SaveTokenBatch, UuidBits, inspect_package_metadata_with_visitor, rewrite_package_metadata,
    rewrite_package_metadata_additions_and_save_tokens,
    rewrite_package_metadata_combined_additions_and_removals_and_save_tokens,
    rewrite_package_metadata_component_removals, rewrite_package_metadata_removals_and_save_tokens,
    rewrite_package_metadata_save_tokens,
};

use litchi_iwa_protos::package_metadata_media_codec::{
    self as media, DataInfoRemoval, DataInfoSnapshot, MediaRewriteBatch,
    PackageMetadataMediaVisitor,
};

use super::{Fault, FaultLimit, FaultResult, TableLeaf, Working, allocation, reserve};
use crate::cell_format::metadata::{ENTRY_NAME, MESSAGE_TYPE, normalized_locator};

/// One current component registration.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Component {
    pub identifier: u64,
    pub locator: String,
}

/// One object UUID binding.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Uuid {
    pub component: u64,
    pub object: u64,
    pub uuid: UuidBits,
}

/// One cross-component object reference record.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct External {
    pub source: u64,
    pub target: u64,
    pub object: u64,
    pub weak: Option<bool>,
}

/// One component data-reference owner record: `object` of `component`
/// owns `count` references to embedded data `data`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DataOwner {
    pub component: u64,
    pub data: u64,
    pub object: u64,
    pub count: u32,
}

/// One new embedded asset: its `DataInfo` record.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DataAddition {
    pub identifier: u64,
    pub digest: Vec<u8>,
    pub preferred: String,
    pub file: String,
    pub length: u64,
}

/// The registry facts of one metadata payload.
#[derive(Debug, Default)]
pub struct Registry {
    pub components: Vec<Component>,
    pub uuids: Vec<Uuid>,
    pub externals: Vec<External>,
    /// Objects that own a data (media) reference.
    pub data_owners: Vec<u64>,
    /// Every component data-reference owner record.
    pub data_references: Vec<DataOwner>,
    pub last_object_identifier: u64,
    unknown: bool,
}

impl Registry {
    /// The current component registered for member `name`.
    pub fn component_for_entry(&self, name: &str) -> FaultResult<&Component> {
        let locator = normalized_locator(name);
        let mut found = None;
        for component in &self.components {
            if component.locator == locator && found.replace(component).is_some() {
                return Err(Fault::Invalid);
            }
        }
        found.ok_or(Fault::Invalid)
    }

    fn component(&self, identifier: u64) -> FaultResult<&Component> {
        self.components
            .iter()
            .find(|component| component.identifier == identifier)
            .ok_or(Fault::Invalid)
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
        let mut locator = String::new();
        locator
            .try_reserve_exact(component.effective_locator().len())
            .map_err(|_error| RewriteError::allocation(component.effective_locator().len()))?;
        locator.push_str(component.effective_locator());
        remember(
            &mut self.registry.components,
            Component {
                identifier: component.identifier(),
                locator,
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
        }
        Ok(())
    }

    fn visit_data_reference_owner(
        &mut self,
        owner: DataReferenceOwnerDescriptor<'_>,
    ) -> Result<(), RewriteError> {
        remember(
            &mut self.registry.data_references,
            DataOwner {
                component: owner.component().identifier(),
                data: owner.data_identifier(),
                object: owner.object_identifier(),
                count: owner.count(),
            },
        )?;
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

/// The metadata object and message in the working set, if the package has
/// a registry.
fn route(working: &Working<'_, impl TableLeaf>) -> FaultResult<Option<(u64, usize)>> {
    let Ok(archive) = working.archive(ENTRY_NAME) else {
        return Ok(None);
    };
    let mut found = None;
    for object in &archive.objects {
        for (index, message) in object.messages.iter().enumerate() {
            if message.type_ == MESSAGE_TYPE {
                let identifier = object.archive_info.identifier.ok_or(Fault::Invalid)?;
                if found.replace((identifier, index)).is_some() {
                    return Err(Fault::Invalid);
                }
            }
        }
    }
    found.map(Some).ok_or(Fault::Invalid)
}

fn payload<'w>(
    working: &'w Working<'_, impl TableLeaf>,
    (object, message): (u64, usize),
) -> FaultResult<&'w [u8]> {
    working
        .archive(ENTRY_NAME)?
        .object(object)
        .and_then(|object| object.messages.get(message))
        .map(|message| message.data.as_slice())
        .ok_or(Fault::Invalid)
}

/// The package save token (`TSP.PackageMetadata.save_token`).
pub fn save_token(working: &Working<'_, impl TableLeaf>) -> FaultResult<Option<u64>> {
    let Some(route) = route(working)? else {
        return Ok(None);
    };
    let values = litchi_iwa_common::wire::repeated_varint_values(payload(working, route)?, 8)
        .map_err(|error| super::map_wire(&error))?;
    Ok(values.last().copied())
}

/// `data` (a `TSP.PackageMetadata`) with the package save token `token`, and
/// every component record that a composed transaction advanced past it
/// brought back to it: the transaction is one save.
fn with_save_token(data: &[u8], token: u64) -> FaultResult<Vec<u8>> {
    use litchi_iwa_common::WireLimits;
    use litchi_iwa_common::wire::{
        append_length_delimited_field, parse_wire_view, patch_varint_field, repeated_varint_values,
    };
    let view =
        parse_wire_view(data, WireLimits::DEFAULT).map_err(|error| super::map_wire(&error))?;
    let mut output = Vec::new();
    reserve(&mut output, data.len())?;
    for field in view.fields() {
        if matches!(field.number(), 3 | 11) {
            let record = field.payload();
            let current =
                repeated_varint_values(record, 12).map_err(|error| super::map_wire(&error))?;
            if current.last().is_some_and(|value| *value > token) {
                let patched = patch_varint_field(record, 12, true, Some(token))
                    .map_err(|error| super::map_wire(&error))?;
                append_length_delimited_field(
                    &mut output,
                    field.number(),
                    &patched,
                    WireLimits::DEFAULT,
                )
                .map_err(|error| super::map_wire(&error))?;
                continue;
            }
        }
        output.extend_from_slice(field.raw());
    }
    patch_varint_field(&output, 8, true, Some(token)).map_err(|error| super::map_wire(&error))
}

fn options(
    working: &Working<'_, impl TableLeaf>,
    payload: usize,
    additions: usize,
) -> RewriteOptions {
    // These are caller/source ceilings, not estimates derived from the current
    // payload.  The prepared codec phase computes the exact candidate size and
    // execution requirements; this profile only prevents an unbounded scan or
    // decode from being admitted while it prepares that transition.
    let profile = working.source().archive_limits().archive_limits();
    let message_bytes = profile
        .max_message_bytes()
        .clamp(1, litchi_iwa_common::WireLimits::MAX_OUTPUT_BYTES);
    let fields = profile
        .max_header_fields()
        .clamp(1, litchi_iwa_common::WireLimits::MAX_FIELDS);
    let work = profile
        .max_header_memory_bytes()
        .clamp(1, litchi_iwa_common::WireLimits::MAX_REWRITE_WORK);
    let nesting = u32::try_from(profile.max_header_nesting()).unwrap_or(u32::MAX);
    // Metadata can retain versioned/component records that are not one-to-one
    // with physical catalog members.  The caller's semantic object ceiling is
    // the finite admission dimension for that registry census.
    let components = working.source().max_objects().max(1);
    RewriteOptions::new(
        payload.max(1).min(message_bytes),
        message_bytes,
        fields,
        work,
        nesting.max(1),
        components,
        working.source().max_references().max(1),
        additions.max(1),
    )
}

/// The registry of the package, or `None` without `Index/Metadata.iwa`.
pub fn registry(working: &Working<'_, impl TableLeaf>) -> FaultResult<Option<Registry>> {
    let Some(route) = route(working)? else {
        return Ok(None);
    };
    let source = payload(working, route)?;
    let mut visitor = Visitor::default();
    let inspection = inspect_package_metadata_with_visitor(
        source,
        options(working, source.len(), 0),
        &mut visitor,
    )
    .map_err(|error| map_rewrite(&error))?;
    visitor.registry.last_object_identifier = inspection.last_object_identifier();
    if visitor.registry.unknown {
        return Err(Fault::Unsupported);
    }
    Ok(Some(visitor.registry))
}

/// The issued-identifier watermark, or `None` without a registry.
pub fn last_object_identifier(working: &Working<'_, impl TableLeaf>) -> FaultResult<Option<u64>> {
    Ok(registry(working)?.map(|registry| registry.last_object_identifier))
}

/// The registry records one transaction adds and removes.
#[derive(Debug, Default)]
pub struct Transition {
    pub uuid_additions: Vec<Uuid>,
    pub uuid_removals: Vec<Uuid>,
    pub external_removals: Vec<External>,
    pub external_additions: Vec<External>,
    /// Removed objects whose identifiers may lower the watermark when they
    /// are its suffix (the host editor's `release_package_identifier_suffix`).
    pub released: Vec<u64>,
    /// Components whose save token advances although their registry
    /// records do not change (the host's `advance_save_tokens_for_entries`
    /// over the components a route rewrites).
    pub stamps: Vec<u64>,
    /// Data-reference owner records of removed objects.
    pub data_owner_removals: Vec<DataOwner>,
    /// Embedded data no surviving object references: its `DataInfo` record
    /// leaves the registry and its package entry is deleted.
    pub data_removals: Vec<(u64, String)>,
    /// Unreferenced data with no package entry (an unmaterialized record):
    /// only its `DataInfo` record leaves the registry (record 0417).
    pub data_record_removals: Vec<u64>,
    /// Data-reference owner records of new objects (a duplicate sharing its
    /// source's embedded data).
    pub data_owner_additions: Vec<DataOwner>,
    /// New embedded assets (their `DataInfo` records).
    pub data_additions: Vec<DataAddition>,
    /// New non-IWA package members (embedded data).
    pub entry_insertions: Vec<(String, Vec<u8>)>,
    /// Non-IWA package members (embedded data) this transaction rewrites.
    pub entry_writes: Vec<(String, Vec<u8>)>,
}

impl Transition {
    /// Record that component `source` no longer references `object` in
    /// component `target` (a no-op without that registry edge); an edge
    /// this transition added is withdrawn.
    pub fn remove_reference(
        &mut self,
        registry: &Registry,
        source: u64,
        target: u64,
        object: u64,
    ) -> FaultResult<()> {
        self.external_additions.retain(|external| {
            !(external.source == source && external.target == target && external.object == object)
        });
        for external in &registry.externals {
            if external.source == source
                && external.target == target
                && external.object == object
                && !self.external_removals.contains(external)
            {
                reserve(&mut self.external_removals, 1)?;
                self.external_removals.push(*external);
            }
        }
        Ok(())
    }

    /// Lower the watermark over `objects` when the last registered
    /// identifier is one of them (record 0370).
    pub fn release_suffix(&mut self, objects: &[u64]) -> FaultResult<()> {
        reserve(&mut self.released, objects.len())?;
        self.released.extend_from_slice(objects);
        Ok(())
    }

    /// Advance the save token of `component` (record 0370d).
    pub fn stamp(&mut self, component: u64) -> FaultResult<()> {
        if !self.stamps.contains(&component) {
            reserve(&mut self.stamps, 1)?;
            self.stamps.push(component);
        }
        Ok(())
    }

    fn is_empty(&self) -> bool {
        self.registry_is_empty()
            && self.data_owner_removals.is_empty()
            && self.data_owner_additions.is_empty()
            && self.data_additions.is_empty()
            && self.data_removals.is_empty()
            && self.data_record_removals.is_empty()
    }

    /// The package entries this transition deletes (unreferenced data).
    pub fn deleted_entries(&self) -> impl Iterator<Item = &str> {
        self.data_removals.iter().map(|(_, entry)| entry.as_str())
    }

    fn registry_is_empty(&self) -> bool {
        self.stamps.is_empty()
            && self.uuid_additions.is_empty()
            && self.uuid_removals.is_empty()
            && self.external_removals.is_empty()
            && self.external_additions.is_empty()
    }

    /// Record the registry consequences of removing `objects`: their UUID
    /// bindings and every external reference that addresses them.
    pub fn remove_objects(&mut self, registry: &Registry, objects: &[u64]) -> FaultResult<()> {
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
        for owner in &registry.data_references {
            if objects.contains(&owner.object) && !self.data_owner_removals.contains(owner) {
                reserve(&mut self.data_owner_removals, 1)?;
                self.data_owner_removals.push(*owner);
            }
        }
        Ok(())
    }

    /// Record the removal of every embedded data whose owners this
    /// transition removed and that nothing in `working` (after its object
    /// removals) or the registry still references: the host's
    /// `remove_unreferenced` over the media assets of a removed drawable.
    /// Data the package's `DataMetadataMap` may describe is refused.
    pub fn release_unreferenced_data(
        &mut self,
        working: &Working<'_, impl TableLeaf>,
        registry: &Registry,
    ) -> FaultResult<()> {
        let mut candidates: Vec<u64> = Vec::new();
        for owner in &self.data_owner_removals {
            if !candidates.contains(&owner.data) {
                reserve(&mut candidates, 1)?;
                candidates.push(owner.data);
            }
        }
        candidates.retain(|data| {
            !registry
                .data_references
                .iter()
                .any(|owner| owner.data == *data && !self.data_owner_removals.contains(owner))
        });
        if candidates.is_empty() {
            return Ok(());
        }
        for name in working.component_names() {
            for object in &working.archive(&name)?.objects {
                for info in &object.archive_info.message_infos {
                    candidates.retain(|data| !info.data_references.contains(data));
                }
            }
        }
        if candidates.is_empty() {
            return Ok(());
        }
        let Some(route) = route(working)? else {
            return Err(Fault::Unsupported);
        };
        let source = payload(working, route)?;
        let mut files = DataFiles::default();
        let report = media::visit_package_metadata_media(
            source,
            media::DecodeOptions::for_source(source),
            &mut files,
        )
        .map_err(|_error| Fault::Invalid)?;
        let map = data_metadata_map(working, report)?;
        if let Some(map) = map {
            let options = media::DecodeOptions::for_source(map.payload());
            let mut kept = Vec::new();
            for data in &candidates {
                if !map
                    .contains_data_identifier(*data, options)
                    .map_err(|_error| Fault::Invalid)?
                {
                    reserve(&mut kept, 1)?;
                    kept.push(*data);
                }
            }
            candidates = kept;
        }
        for data in candidates {
            let (_, file) = files
                .records
                .iter()
                .find(|(identifier, _)| *identifier == data)
                .ok_or(Fault::Invalid)?;
            let entry = data_entry_name(file)?;
            reserve(&mut self.data_removals, 1)?;
            self.data_removals.push((data, entry));
        }
        Ok(())
    }

    /// Register a fresh UUID for new object `object` of component
    /// `component`, derived from `seed` and distinct from every registered
    /// UUID (no clock or randomness is consulted).
    pub fn register_uuid(
        &mut self,
        registry: &Registry,
        component: u64,
        object: u64,
        seed: u64,
    ) -> FaultResult<()> {
        let mut state = seed ^ object.rotate_left(17);
        loop {
            let lower = mix(state);
            let upper = mix(lower ^ 0x6c69_7463_6869_0301);
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

    /// Record that component `source` now references `object` in component
    /// `target`, unless the registry already holds that edge.
    pub fn reference(
        &mut self,
        registry: &Registry,
        source: u64,
        target: u64,
        object: u64,
    ) -> FaultResult<()> {
        if source == target {
            return Ok(());
        }
        let known = |external: &External| {
            external.source == source && external.target == target && external.object == object
        };
        if registry.externals.iter().any(known) || self.external_additions.iter().any(known) {
            return Ok(());
        }
        reserve(&mut self.external_additions, 1)?;
        self.external_additions.push(External {
            source,
            target,
            object,
            weak: None,
        });
        Ok(())
    }
}

/// Apply `transition` and raise the watermark to `allocated`.
pub fn apply(
    working: &mut Working<'_, impl TableLeaf>,
    transition: &Transition,
    allocated: Option<u64>,
    emptied: &[String],
) -> FaultResult<()> {
    let Some(route) = route(working)? else {
        if transition.is_empty() && emptied.is_empty() {
            return Ok(());
        }
        return Err(Fault::Unsupported);
    };
    // Nothing to register: no transition, no released or allocated
    // identifier, no emptied component and no save token (the registry is
    // not decoded).
    if transition.is_empty()
        && transition.released.is_empty()
        && allocated.is_none()
        && emptied.is_empty()
        && working.save_token.is_none()
    {
        return Ok(());
    }
    let registry = registry(working)?.ok_or(Fault::Invalid)?;
    let current = registry.last_object_identifier;
    let mut last = allocated.map_or(current, |allocated| allocated.max(current));
    if allocated.is_none() && transition.released.contains(&current) {
        last = working.maximum_identifier();
    }
    if transition.is_empty()
        && last == current
        && emptied.is_empty()
        && working.save_token.is_none()
    {
        return Ok(());
    }
    // Data owners and unreferenced data leave first: the registry codec
    // refuses to drop the UUID of an object that still owns data.
    let staged;
    let source = if transition.data_owner_removals.is_empty()
        && transition.data_owner_additions.is_empty()
        && transition.data_additions.is_empty()
        && transition.data_removals.is_empty()
        && transition.data_record_removals.is_empty()
    {
        payload(working, route)?
    } else {
        staged = rewrite_media(working, payload(working, route)?, &registry, transition)?;
        staged.as_slice()
    };
    let additions = transition
        .uuid_removals
        .len()
        .saturating_add(transition.uuid_additions.len())
        .saturating_add(transition.external_removals.len())
        .saturating_add(transition.external_additions.len());
    let options = options(working, source.len(), additions);
    let mut restore = None;
    let output = if transition.registry_is_empty() && last == current {
        None
    } else if transition.registry_is_empty() && last < current {
        // Only the watermark goes down (the last issued object was removed
        // and nothing else changes in the registry): a field patch, as the
        // host editor's `release_package_identifier_suffix` writes it.
        restore = Some(last);
        None
    } else if transition.registry_is_empty() {
        Some(
            rewrite_package_metadata(source, Batch::new(current, last, &[], &[]), options)
                .map_err(|error| map_rewrite(&error))?,
        )
    } else {
        let selector = |identifier: u64| -> FaultResult<ComponentSelector<'_>> {
            let component = registry.component(identifier)?;
            Ok(ComponentSelector::new(
                component.identifier,
                component.locator.as_str(),
            ))
        };
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
        // Every component whose registry record changes advances its save
        // token, as the native app does when it rewrites a component.
        let mut tokens: Vec<ComponentSelector<'_>> = Vec::new();
        let touched = transition
            .uuid_removals
            .iter()
            .chain(&transition.uuid_additions)
            .map(|binding| binding.component)
            .chain(
                transition
                    .external_removals
                    .iter()
                    .chain(&transition.external_additions)
                    .flat_map(|external| [external.source, external.target]),
            )
            .chain(transition.stamps.iter().copied());
        for identifier in touched {
            let selector = selector(identifier)?;
            if !tokens.contains(&selector) {
                reserve(&mut tokens, 1)?;
                tokens.push(selector);
            }
        }
        let removes = !uuid_removals.is_empty() || !external_removals.is_empty();
        let adds = !external_additions.is_empty() || !uuid_additions.is_empty();
        // The addition batches require a rising watermark; a transaction
        // that adds registry records without allocating (a reference to an
        // existing style, record 0370b) passes one step up and the
        // watermark is put back to `last` below.
        let unraised = adds && last <= current;
        if unraised {
            restore = Some(last);
        }
        let last = if unraised {
            current.checked_add(1).ok_or(Fault::Limit {
                kind: FaultLimit::Identifiers,
                observed: u64::MAX,
                maximum: u64::MAX,
            })?
        } else {
            last
        };
        let stamped = !transition.stamps.is_empty();
        let output = if !removes && !adds && last == current {
            rewrite_package_metadata_save_tokens(source, SaveTokenBatch::new(&tokens), options)
        } else if !removes && !adds && last > current {
            rewrite_package_metadata_additions_and_save_tokens(
                source,
                AdditionSaveTokenBatch::new(
                    Batch::new(current, last, &[], &[]),
                    SaveTokenBatch::new(&tokens),
                ),
                options,
            )
        } else if !adds && last <= current {
            rewrite_package_metadata_removals_and_save_tokens(
                source,
                RemovalSaveTokenBatch::new(
                    RemovalBatch::new(current, &uuid_removals, &external_removals, &[])
                        .with_new_last_object_identifier(last),
                    SaveTokenBatch::new(&tokens),
                ),
                options,
            )
        } else if adds && !removes && stamped {
            rewrite_package_metadata_additions_and_save_tokens(
                source,
                AdditionSaveTokenBatch::new(
                    Batch::new(current, last, &uuid_additions, &external_additions),
                    SaveTokenBatch::new(&tokens),
                ),
                options,
            )
        } else if adds && !removes {
            rewrite_package_metadata(
                source,
                Batch::new(current, last, &uuid_additions, &external_additions),
                options,
            )
        } else {
            rewrite_package_metadata_combined_additions_and_removals_and_save_tokens(
                source,
                CombinedSaveTokenBatch::new(
                    CombinedBatch::new(
                        current,
                        last,
                        &uuid_additions,
                        &external_additions,
                        &uuid_removals,
                        &external_removals,
                        &[],
                    ),
                    SaveTokenBatch::new(&tokens),
                ),
                options,
            )
        };
        Some(output.map_err(|error| map_rewrite(&error))?)
    };
    let mut data = match output {
        Some(output) => output.into_bytes(),
        None => super::copy_bytes(source)?,
    };
    if let Some(last) = restore {
        data = litchi_iwa_common::wire::patch_varint_field(&data, 1, true, Some(last))
            .map_err(|error| super::map_wire(&error))?;
    }
    if let Some(token) = working.save_token {
        data = with_save_token(&data, token)?;
    }
    if !emptied.is_empty() {
        let mut selectors = Vec::new();
        reserve(&mut selectors, emptied.len())?;
        for name in emptied {
            let component = registry.component_for_entry(name)?;
            selectors.push(ComponentSelector::new(
                component.identifier,
                component.locator.as_str(),
            ));
        }
        let options = self::options(working, data.len(), emptied.len());
        data = rewrite_package_metadata_component_removals(
            &data,
            ComponentRemovalBatch::new(&selectors),
            options,
        )
        .map_err(|error| map_rewrite(&error))?
        .into_bytes();
    }
    let limits = working.archive_limits();
    let object = working
        .archive_mut(ENTRY_NAME)?
        .object_mut(route.0)
        .ok_or(Fault::Invalid)?;
    object
        .replace_message_preserving_header(
            route.1,
            RawMessage {
                type_: MESSAGE_TYPE,
                data,
            },
            limits,
        )
        .map_err(|error| super::map_core(&error))?;
    Ok(())
}

/// The next embedded-data identifier: one above every `DataInfo` record
/// and every asset this transition adds (the host's
/// `insert_unreferenced`).
pub fn next_data_identifier(
    working: &Working<'_, impl TableLeaf>,
    transition: &Transition,
) -> FaultResult<u64> {
    let Some(route) = route(working)? else {
        return Err(Fault::Unsupported);
    };
    let source = payload(working, route)?;
    let mut files = DataFiles::default();
    media::visit_package_metadata_media(
        source,
        media::DecodeOptions::for_source(source),
        &mut files,
    )
    .map_err(|_error| Fault::Invalid)?;
    let maximum = files
        .records
        .iter()
        .map(|(identifier, _)| *identifier)
        .chain(transition.data_additions.iter().map(|data| data.identifier))
        .max()
        .unwrap_or(0);
    maximum.checked_add(1).ok_or(Fault::Invalid)
}

/// The `DataInfo` file names of a metadata payload.
#[derive(Default)]
struct DataFiles {
    records: Vec<(u64, String)>,
}

impl PackageMetadataMediaVisitor for DataFiles {
    fn visit_data_info(
        &mut self,
        data_info: DataInfoSnapshot<'_>,
    ) -> Result<(), media::DecodeError> {
        let name = data_info
            .file_name()
            .unwrap_or_else(|| data_info.preferred_file_name());
        self.records.push((data_info.identifier(), name.to_owned()));
        Ok(())
    }
}

/// The package's `DataMetadataMap`, when the registry names one.
fn data_metadata_map<'w>(
    working: &'w Working<'_, impl TableLeaf>,
    report: media::DecodeReport,
) -> FaultResult<Option<media::DataMetadataMapSource<'w>>> {
    if !report.data_metadata_map_present() {
        return Ok(None);
    }
    let identifier = report
        .data_metadata_map_identifier()
        .ok_or(Fault::Unsupported)?;
    let object = working.object(identifier)?;
    let [message] = object.messages.as_slice() else {
        return Err(Fault::Unsupported);
    };
    let options = media::DecodeOptions::for_source(&message.data);
    media::DataMetadataMapSource::from_source(identifier, &message.data, options)
        .map(Some)
        .map_err(|_error| Fault::Invalid)
}

/// The package entry of a `DataInfo` file name (the host's
/// `data_entry_name`), refusing names that leave `Data/`.
fn data_entry_name(file: &str) -> FaultResult<String> {
    if file.is_empty()
        || file.contains(['\0', '\\'])
        || file.starts_with('/')
        || file
            .split('/')
            .any(|part| part.is_empty() || part == "." || part == "..")
    {
        return Err(Fault::Invalid);
    }
    Ok(format!("Data/{file}"))
}

/// Apply the transition's data-owner and data removals to `source`.
fn rewrite_media(
    working: &Working<'_, impl TableLeaf>,
    source: &[u8],
    registry: &Registry,
    transition: &Transition,
) -> FaultResult<Vec<u8>> {
    let mut owners = Vec::new();
    reserve(&mut owners, transition.data_owner_removals.len())?;
    for owner in &transition.data_owner_removals {
        let component = registry.component(owner.component)?;
        owners.push(media::DataReferenceOwnerRemoval::new(
            media::ComponentSelector::new(component.identifier, component.locator.as_str()),
            owner.data,
            owner.object,
            owner.count,
        ));
    }
    let mut additions = Vec::new();
    reserve(&mut additions, transition.data_owner_additions.len())?;
    for owner in &transition.data_owner_additions {
        let component = registry.component(owner.component)?;
        additions.push(media::DataReferenceOwnerAddition::new(
            media::ComponentSelector::new(component.identifier, component.locator.as_str()),
            owner.data,
            owner.object,
            owner.count,
        ));
    }
    let mut data_additions = Vec::new();
    reserve(&mut data_additions, transition.data_additions.len())?;
    for data in &transition.data_additions {
        data_additions.push(
            media::DataInfoAddition::new(data.identifier, &data.digest, &data.preferred)
                .with_file_name(&data.file)
                .with_materialized_length(data.length),
        );
    }
    let mut removals = Vec::new();
    reserve(
        &mut removals,
        transition
            .data_removals
            .len()
            .saturating_add(transition.data_record_removals.len()),
    )?;
    for (data, _) in &transition.data_removals {
        removals.push(DataInfoRemoval::new(*data));
    }
    for data in &transition.data_record_removals {
        removals.push(DataInfoRemoval::new(*data));
    }
    // The owner and DataInfo planner walks every record of the source per
    // removal; scale the work ceiling to the batch.
    let options = media::DecodeOptions::for_source(source).with_max_work_bytes(
        source.len().max(1).saturating_mul(64).saturating_mul(
            owners
                .len()
                .saturating_add(additions.len())
                .saturating_add(data_additions.len())
                .saturating_add(removals.len())
                .saturating_add(1),
        ),
    );
    let report =
        media::inspect_package_metadata_media(source, options).map_err(|_error| Fault::Invalid)?;
    let mut batch = MediaRewriteBatch::new(&data_additions, &removals, &additions, &owners);
    if !removals.is_empty()
        && let Some(map) = data_metadata_map(working, report)?
    {
        batch = batch.with_data_metadata_map_source(map);
    }
    let output = media::rewrite_package_metadata_media(source, batch, options)
        .map_err(|_error| Fault::Invalid)?;
    Ok(output.into_bytes())
}

/// A 64-bit finalizer (SplitMix64) for deterministic identity derivation.
pub const fn mix(mut value: u64) -> u64 {
    value = value.wrapping_add(0x9e37_79b9_7f4a_7c15);
    value = (value ^ (value >> 30)).wrapping_mul(0xbf58_476d_1ce4_e5b9);
    value = (value ^ (value >> 27)).wrapping_mul(0x94d0_49bb_1331_11eb);
    value ^ (value >> 31)
}

/// A deterministic seed for the fresh identities of one transaction: the
/// registry payload (which carries the document's own identity) and the
/// first identifier the transaction allocates.
pub fn identity_seed(working: &Working<'_, impl TableLeaf>, first: u64) -> FaultResult<u64> {
    let mut seed = mix(first);
    if let Some(route) = route(working)? {
        for chunk in payload(working, route)?.chunks(8) {
            let mut word = [0_u8; 8];
            word[..chunk.len()].copy_from_slice(chunk);
            seed = mix(seed ^ u64::from_le_bytes(word));
        }
    }
    Ok(seed)
}
