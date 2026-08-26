//! Strict PackageMetadata primitives for direct comment-reply transactions.
//!
//! This module deliberately contains no comment-storage or generated-protobuf
//! vocabulary. It owns the metadata side of a reply COW/cull transaction:
//! one exact Metadata.iwa route, an exhaustive current/versioned registry
//! census, collision-safe identifiers, and one prepared atomic transition.

use core::mem::size_of;

use litchi_iwa_core::{ArchiveReferencePolicy, ArchiveReferenceVisitor, Limits as ArchiveLimits};
use litchi_iwa_protos::package_metadata_codec::{
    AdditionSaveTokenBatch, CombinedBatch, CombinedSaveTokenBatch, ComponentDescriptor,
    ComponentSelector, DataReferenceOwnerDescriptor, ExternalReferenceAddition,
    ExternalReferenceDescriptor, ExternalReferenceRemoval, InvalidReason, ObjectUuidAddition,
    ObjectUuidDescriptor, ObjectUuidRemoval, PackageMetadataInspection, PackageMetadataVisitor,
    PreparedPackageMetadataAdditionSaveTokenRewrite,
    PreparedPackageMetadataCombinedSaveTokenRewrite,
    PreparedPackageMetadataRemovalSaveTokenRewrite, PreparedPackageMetadataSaveTokenRewrite,
    RemovalSaveTokenBatch, RewriteError, RewriteLimit, RewriteOptions, RewriteReport,
    SaveTokenBatch, UuidBits, inspect_package_metadata_with_visitor,
    prepare_package_metadata_additions_and_save_tokens,
    prepare_package_metadata_combined_additions_and_removals_and_save_tokens,
    prepare_package_metadata_removals_and_save_tokens, prepare_package_metadata_save_tokens,
};

use super::{Package, metadata};

const METADATA_MESSAGE_TYPE: u32 = 11_006;

/// A content-free classification of a metadata refusal.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum FailureKind {
    MissingRoute,
    AmbiguousRoute,
    InvalidSource,
    Conflict,
    Unsupported,
    VersionedOwnership,
    Limit,
    Allocation,
}

/// Redacted metadata failure returned to the package owner.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct MetadataError {
    pub(super) kind: FailureKind,
    pub(super) observed: usize,
    pub(super) maximum: usize,
    pub(super) allocation: usize,
}

impl MetadataError {
    const fn invalid() -> Self {
        Self {
            kind: FailureKind::InvalidSource,
            observed: 0,
            maximum: 0,
            allocation: 0,
        }
    }

    const fn kind(kind: FailureKind) -> Self {
        Self {
            kind,
            observed: 0,
            maximum: 0,
            allocation: 0,
        }
    }

    const fn limit(observed: usize, maximum: usize) -> Self {
        Self {
            kind: FailureKind::Limit,
            observed,
            maximum,
            allocation: 0,
        }
    }

    const fn allocation(amount: usize) -> Self {
        Self {
            kind: FailureKind::Allocation,
            observed: amount,
            maximum: 0,
            allocation: amount,
        }
    }
}

type Result<T> = core::result::Result<T, MetadataError>;

/// Borrowed unique Metadata.iwa payload and its physical route.
#[derive(Debug, Clone, Copy)]
pub(super) struct MetadataSource<'source> {
    pub(super) route: metadata::MessageRoute,
    pub(super) payload: &'source [u8],
}

impl<'source> MetadataSource<'source> {
    #[must_use]
    pub(super) const fn payload(self) -> &'source [u8] {
        self.payload
    }

    #[must_use]
    pub(super) const fn route(self) -> metadata::MessageRoute {
        self.route
    }
}

/// Codec-owned atomic transition aliases kept private to the package owner.
pub(super) type CombinedTransition<'source> = CombinedBatch<'source>;
pub(super) type PreparedCombinedTransition<'source, 'batch> =
    PreparedPackageMetadataCombinedSaveTokenRewrite<'source, 'batch>;

/// Locate exactly one current Metadata.iwa object carrying type 11006.
pub(super) fn strict_source(source: &Package) -> Result<MetadataSource<'_>> {
    let route = metadata::unique_message_route(source)
        .ok_or_else(|| MetadataError::kind(FailureKind::MissingRoute))?;
    let payload = source
        .state
        .components
        .catalog()
        .get_index(route.component_index)
        .and_then(|component| component.archive().objects.get(route.object_index))
        .and_then(|object| object.messages.get(route.message_index))
        .filter(|message| message.type_ == METADATA_MESSAGE_TYPE)
        .map(|message| message.data.as_slice())
        .ok_or_else(MetadataError::invalid)?;
    Ok(MetadataSource { route, payload })
}

/// One current or versioned component with its source-authoritative locator.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct ComponentFact {
    pub(super) component_index: Option<usize>,
    pub(super) identifier: u64,
    pub(super) preferred_locator: String,
    pub(super) effective_locator: String,
    pub(super) current: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct UuidFact {
    component_index: Option<usize>,
    object_identifier: u64,
    uuid: UuidBits,
    current: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct ExternalFact {
    source_component_index: Option<usize>,
    source_identifier: u64,
    target_component_identifier: u64,
    object_identifier: Option<u64>,
    is_weak: Option<bool>,
    current: bool,
    versioned: bool,
}

/// Exact object-specific external edge used by a cross-component reply
/// graph. Weakness is explicit so a deprecated/ambiguous edge cannot be
/// silently treated as ownership.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct ExternalEdge {
    pub(super) source_component_index: usize,
    pub(super) target_component_index: usize,
    pub(super) object_identifier: u64,
    pub(super) is_weak: Option<bool>,
}

impl ExternalEdge {
    pub(super) const fn new(
        source_component_index: usize,
        target_component_index: usize,
        object_identifier: u64,
        is_weak: Option<bool>,
    ) -> Self {
        Self {
            source_component_index,
            target_component_index,
            object_identifier,
            is_weak,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct DataOwnerFact {
    data_identifier: u64,
    object_identifier: u64,
}

/// Exhaustive current/versioned metadata registry facts.
#[derive(Debug)]
pub(super) struct MetadataRegistry<'source> {
    source: MetadataSource<'source>,
    components: Vec<ComponentFact>,
    uuids: Vec<UuidFact>,
    external_references: Vec<ExternalFact>,
    data_owners: Vec<DataOwnerFact>,
    ambiguous_identifiers: Vec<u64>,
    root_data_map_identifier: Option<u64>,
    last_object_identifier: u64,
    maximum_identifier: u64,
    physical_identifiers: Vec<u64>,
    report: RewriteReport,
}

/// Generic spelling retained for owner code that avoids implementation names.
pub(super) type RegistryFacts<'source> = MetadataRegistry<'source>;

impl<'source> MetadataRegistry<'source> {
    #[must_use]
    pub(super) const fn payload(&self) -> &'source [u8] {
        self.source.payload
    }

    #[must_use]
    pub(super) const fn route(&self) -> metadata::MessageRoute {
        self.source.route
    }

    #[must_use]
    pub(super) const fn report(&self) -> RewriteReport {
        self.report
    }

    #[must_use]
    pub(super) const fn last_object_identifier(&self) -> u64 {
        self.last_object_identifier
    }

    #[must_use]
    pub(super) const fn maximum_identifier(&self) -> u64 {
        self.maximum_identifier
    }

    /// Resolve one exact current component using its effective locator.
    pub(super) fn selector(&self, component_index: usize) -> Result<ComponentSelector<'_>> {
        let mut found = None;
        for component in &self.components {
            if !component.current || component.component_index != Some(component_index) {
                continue;
            }
            if found.replace(component).is_some() {
                return Err(MetadataError::kind(FailureKind::AmbiguousRoute));
            }
        }
        let component = found.ok_or_else(MetadataError::invalid)?;
        Ok(ComponentSelector::new(
            component.identifier,
            component.effective_locator.as_str(),
        ))
    }

    pub(super) fn selectors_for_components(
        &self,
        component_indices: &[usize],
    ) -> Result<Vec<ComponentSelector<'_>>> {
        let mut selectors = Vec::new();
        selectors
            .try_reserve_exact(component_indices.len())
            .map_err(|_| MetadataError::allocation(component_indices.len()))?;
        for &component_index in component_indices {
            let selector = self.selector(component_index)?;
            if selectors.iter().any(|item| *item == selector) {
                continue;
            }
            selectors.push(selector);
        }
        if selectors.is_empty() {
            return Err(MetadataError::invalid());
        }
        Ok(selectors)
    }

    /// Resolve all changed current component selectors in first-seen order,
    /// deduplicating co-located model/storage edits.
    pub(super) fn touched_selectors(
        &self,
        component_indices: &[usize],
    ) -> Result<Vec<ComponentSelector<'_>>> {
        if component_indices.is_empty() {
            return Err(MetadataError::invalid());
        }
        let mut unique = Vec::new();
        unique
            .try_reserve_exact(component_indices.len())
            .map_err(|_| MetadataError::allocation(component_indices.len()))?;
        for &component_index in component_indices {
            if unique.contains(&component_index) {
                continue;
            }
            let _ = self.selector(component_index)?;
            unique.push(component_index);
        }
        self.selectors_for_components(&unique)
    }

    /// Validate one deduplicated current/effective save-token selector set.
    pub(super) fn validate_save_tokens(&self, batch: SaveTokenBatch<'_>) -> Result<()> {
        if batch.selectors().is_empty() {
            return Err(MetadataError::invalid());
        }
        for (index, selector) in batch.selectors().iter().enumerate() {
            let exact = self
                .components
                .iter()
                .filter(|component| {
                    component.current
                        && component.identifier == selector.identifier()
                        && component.effective_locator == selector.locator()
                })
                .count();
            if exact != 1 {
                return Err(MetadataError::kind(FailureKind::Conflict));
            }
            if batch.selectors()[index + 1..]
                .iter()
                .any(|other| other == selector)
            {
                return Err(MetadataError::kind(FailureKind::Conflict));
            }
        }
        Ok(())
    }

    /// Require exactly one current UUID owner and reject every versioned or
    /// alternate namespace claim for the same object.
    pub(super) fn require_current_uuid(
        &self,
        component_index: usize,
        object_identifier: u64,
    ) -> Result<UuidBits> {
        self.reject_non_uuid_owner(object_identifier)?;
        let mut current = None;
        let mut versioned = false;
        for binding in self.uuids.iter().copied() {
            if binding.object_identifier != object_identifier {
                continue;
            }
            if !binding.current {
                versioned = true;
                continue;
            }
            if binding.component_index != Some(component_index)
                || current.replace(binding.uuid).is_some()
            {
                return Err(MetadataError::kind(FailureKind::Conflict));
            }
        }
        if versioned {
            return Err(MetadataError::kind(FailureKind::VersionedOwnership));
        }
        current.ok_or_else(|| MetadataError::kind(FailureKind::MissingRoute))
    }

    /// Inspect an optional current UUID owner. Absence is admitted for
    /// native-compatible objects; hostile alternate ownership is not.
    pub(super) fn current_uuid_if_registered(
        &self,
        component_index: usize,
        object_identifier: u64,
    ) -> Result<Option<UuidBits>> {
        if object_identifier == 0 {
            return Err(MetadataError::invalid());
        }
        self.reject_non_uuid_owner(object_identifier)?;
        let mut current = None;
        let mut versioned = false;
        for binding in self.uuids.iter().copied() {
            if binding.object_identifier != object_identifier {
                continue;
            }
            if !binding.current {
                versioned = true;
            } else if binding.component_index != Some(component_index)
                || current.replace(binding.uuid).is_some()
            {
                return Err(MetadataError::kind(FailureKind::Conflict));
            }
        }
        if versioned {
            return Err(MetadataError::kind(FailureKind::VersionedOwnership));
        }
        Ok(current)
    }

    pub(super) fn current_uuids_if_registered(
        &self,
        owners: &[(usize, u64)],
    ) -> Result<Vec<Option<UuidBits>>> {
        let mut output: Vec<Option<UuidBits>> = Vec::new();
        output
            .try_reserve_exact(owners.len())
            .map_err(|_| MetadataError::allocation(owners.len()))?;
        for (index, &(component_index, object_identifier)) in owners.iter().enumerate() {
            if owners[..index]
                .iter()
                .any(|item| *item == (component_index, object_identifier))
            {
                return Err(MetadataError::kind(FailureKind::Conflict));
            }
            let uuid = self.current_uuid_if_registered(component_index, object_identifier)?;
            if uuid.is_some_and(|value| output.iter().flatten().any(|item| *item == value)) {
                return Err(MetadataError::kind(FailureKind::Conflict));
            }
            output.push(uuid);
        }
        Ok(output)
    }

    /// Prove one current object-specific external edge exactly once.
    pub(super) fn require_external_edge_exact(&self, edge: ExternalEdge) -> Result<()> {
        let source = self.selector(edge.source_component_index)?;
        let target = self.selector(edge.target_component_index)?;
        let mut exact = 0usize;
        for reference in &self.external_references {
            if reference.source_identifier != source.identifier()
                || reference.target_component_identifier != target.identifier()
                || reference.object_identifier != Some(edge.object_identifier)
            {
                continue;
            }
            let weak_matches = match edge.is_weak {
                Some(false) => reference.is_weak != Some(true),
                expected => reference.is_weak == expected,
            };
            if !reference.current || reference.versioned || !weak_matches {
                return Err(MetadataError::kind(FailureKind::Conflict));
            }
            exact = exact
                .checked_add(1)
                .ok_or_else(|| MetadataError::kind(FailureKind::Limit))?;
        }
        if exact != 1 {
            return Err(MetadataError::kind(FailureKind::Conflict));
        }
        Ok(())
    }

    pub(super) fn require_external_edge_absent(&self, edge: ExternalEdge) -> Result<()> {
        let source = self.selector(edge.source_component_index)?;
        let target = self.selector(edge.target_component_index)?;
        if self.external_references.iter().any(|reference| {
            reference.source_identifier == source.identifier()
                && reference.target_component_identifier == target.identifier()
                && (reference.object_identifier.is_none()
                    || reference.object_identifier == Some(edge.object_identifier))
        }) {
            return Err(MetadataError::kind(FailureKind::Conflict));
        }
        Ok(())
    }

    pub(super) fn external_addition(
        &self,
        edge: ExternalEdge,
    ) -> Result<ExternalReferenceAddition<'_>> {
        if edge.object_identifier == 0 {
            return Err(MetadataError::invalid());
        }
        self.require_external_edge_absent(edge)?;
        Ok(ExternalReferenceAddition::new(
            self.selector(edge.source_component_index)?,
            self.selector(edge.target_component_index)?,
            edge.object_identifier,
            edge.is_weak,
        ))
    }

    pub(super) fn external_removal(
        &self,
        edge: ExternalEdge,
    ) -> Result<ExternalReferenceRemoval<'_>> {
        if edge.object_identifier == 0 {
            return Err(MetadataError::invalid());
        }
        self.require_external_edge_exact(edge)?;
        let source = self.selector(edge.source_component_index)?;
        let target = self.selector(edge.target_component_index)?;
        let mut weakness = None;
        for reference in &self.external_references {
            if reference.source_identifier == source.identifier()
                && reference.target_component_identifier == target.identifier()
                && reference.object_identifier == Some(edge.object_identifier)
                && reference.current
                && !reference.versioned
            {
                if weakness.replace(reference.is_weak).is_some() {
                    return Err(MetadataError::kind(FailureKind::Conflict));
                }
            }
        }
        Ok(ExternalReferenceRemoval::new(
            source,
            target,
            edge.object_identifier,
            weakness.ok_or_else(MetadataError::invalid)?,
        ))
    }

    pub(super) fn external_additions(
        &self,
        edges: &[ExternalEdge],
    ) -> Result<Vec<ExternalReferenceAddition<'_>>> {
        let mut output: Vec<ExternalReferenceAddition<'_>> = Vec::new();
        output
            .try_reserve_exact(edges.len())
            .map_err(|_| MetadataError::allocation(edges.len()))?;
        for &edge in edges {
            let addition = self.external_addition(edge)?;
            if output.iter().any(|item| {
                item.source() == addition.source()
                    && item.target() == addition.target()
                    && item.object_identifier() == addition.object_identifier()
            }) {
                return Err(MetadataError::kind(FailureKind::Conflict));
            }
            output.push(addition);
        }
        Ok(output)
    }

    pub(super) fn external_removals(
        &self,
        edges: &[ExternalEdge],
    ) -> Result<Vec<ExternalReferenceRemoval<'_>>> {
        let mut output: Vec<ExternalReferenceRemoval<'_>> = Vec::new();
        output
            .try_reserve_exact(edges.len())
            .map_err(|_| MetadataError::allocation(edges.len()))?;
        for &edge in edges {
            let removal = self.external_removal(edge)?;
            if output.iter().any(|item| {
                item.source() == removal.source()
                    && item.target() == removal.target()
                    && item.object_identifier() == removal.object_identifier()
            }) {
                return Err(MetadataError::kind(FailureKind::Conflict));
            }
            output.push(removal);
        }
        Ok(output)
    }

    /// Reserve an identifier against physical objects and every metadata
    /// namespace: UUID, external, data-owner, root-map, and ambiguous IDs.
    pub(super) fn require_identifier_absent(&self, identifier: u64) -> Result<()> {
        if identifier == 0
            || identifier <= self.last_object_identifier
            || self.physical_identifiers.contains(&identifier)
            || self
                .uuids
                .iter()
                .any(|item| item.object_identifier == identifier)
            || self
                .components
                .iter()
                .any(|item| item.identifier == identifier)
            || self.external_references.iter().any(|item| {
                item.source_identifier == identifier
                    || item.object_identifier == Some(identifier)
                    || item.target_component_identifier == identifier
            })
            || self.data_owners.iter().any(|item| {
                item.object_identifier == identifier || item.data_identifier == identifier
            })
            || self.ambiguous_identifiers.contains(&identifier)
            || self.root_data_map_identifier == Some(identifier)
        {
            return Err(MetadataError::kind(FailureKind::Conflict));
        }
        Ok(())
    }

    pub(super) fn require_uuid_absent(&self, identifier: u64) -> Result<()> {
        self.require_identifier_absent(identifier)
    }

    pub(super) fn uuid_addition(
        &self,
        component_index: usize,
        fresh: FreshIdentifier,
    ) -> Result<ObjectUuidAddition<'_>> {
        self.require_identifier_absent(fresh.identifier)?;
        if self.uuids.iter().any(|item| item.uuid == fresh.uuid) {
            return Err(MetadataError::kind(FailureKind::Conflict));
        }
        Ok(ObjectUuidAddition::new(
            self.selector(component_index)?,
            fresh.identifier,
            fresh.uuid,
        ))
    }

    pub(super) fn uuid_additions(
        &self,
        entries: &[(usize, FreshIdentifier)],
    ) -> Result<Vec<ObjectUuidAddition<'_>>> {
        let mut output: Vec<ObjectUuidAddition<'_>> = Vec::new();
        output
            .try_reserve_exact(entries.len())
            .map_err(|_| MetadataError::allocation(entries.len()))?;
        for &(component_index, fresh) in entries {
            let addition = self.uuid_addition(component_index, fresh)?;
            if output.iter().any(|item| {
                item.object_identifier() == addition.object_identifier()
                    || item.uuid() == addition.uuid()
            }) {
                return Err(MetadataError::kind(FailureKind::Conflict));
            }
            output.push(addition);
        }
        Ok(output)
    }

    pub(super) fn uuid_removal(
        &self,
        component_index: usize,
        object_identifier: u64,
    ) -> Result<ObjectUuidRemoval<'_>> {
        Ok(ObjectUuidRemoval::new(
            self.selector(component_index)?,
            object_identifier,
            self.require_current_uuid(component_index, object_identifier)?,
        ))
    }

    /// Allocate fresh identifiers above the complete physical/metadata
    /// reservation census and deterministic nonzero UUID values.
    pub(super) fn allocate_identifiers(&self, count: usize) -> Result<Vec<FreshIdentifier>> {
        let mut output = Vec::new();
        output
            .try_reserve_exact(count)
            .map_err(|_| MetadataError::allocation(count))?;
        let mut next = self.maximum_identifier;
        let seed = fingerprint(self.payload(), 0xcbf2_9ce4_8422_2325);
        for ordinal in 0..count {
            loop {
                next = next
                    .checked_add(1)
                    .ok_or_else(|| MetadataError::kind(FailureKind::Limit))?;
                if self.require_identifier_absent(next).is_err()
                    || output
                        .iter()
                        .any(|item: &FreshIdentifier| item.identifier == next)
                {
                    continue;
                }
                let ordinal = u64::try_from(ordinal).map_err(|_| MetadataError::invalid())?;
                let uuid = UuidBits::new(
                    next,
                    mix(seed.rotate_left(29) ^ next.rotate_left(17) ^ ordinal),
                );
                if uuid.lower() == 0
                    || uuid.upper() == 0
                    || self.uuids.iter().any(|item| item.uuid == uuid)
                    || output
                        .iter()
                        .any(|item: &FreshIdentifier| item.uuid == uuid)
                {
                    continue;
                }
                output.push(FreshIdentifier {
                    identifier: next,
                    uuid,
                });
                break;
            }
        }
        Ok(output)
    }

    fn reject_non_uuid_owner(&self, identifier: u64) -> Result<()> {
        if self.external_references.iter().any(|item| {
            item.source_identifier == identifier
                || item.object_identifier == Some(identifier)
                || item.target_component_identifier == identifier
        }) || self
            .data_owners
            .iter()
            .any(|item| item.object_identifier == identifier || item.data_identifier == identifier)
            || self
                .components
                .iter()
                .any(|item| item.identifier == identifier)
            || self.ambiguous_identifiers.contains(&identifier)
            || self.root_data_map_identifier == Some(identifier)
        {
            return Err(MetadataError::kind(FailureKind::Conflict));
        }
        Ok(())
    }
}

/// A new native object identifier plus deterministic metadata UUID.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct FreshIdentifier {
    pub(super) identifier: u64,
    pub(super) uuid: UuidBits,
}

#[derive(Debug, Default)]
struct PhysicalFacts {
    maximum_identifier: u64,
    identifiers: Vec<u64>,
    duplicate_identifier: bool,
}

fn physical_facts(source: &Package) -> Result<PhysicalFacts> {
    let mut facts = PhysicalFacts::default();
    for component in source.state.components.catalog().iter() {
        for object in &component.archive().objects {
            let identifier = object
                .archive_info
                .identifier
                .filter(|value| *value != 0)
                .ok_or_else(MetadataError::invalid)?;
            facts
                .identifiers
                .try_reserve(1)
                .map_err(|_| MetadataError::allocation(1))?;
            if facts.identifiers.contains(&identifier) {
                facts.duplicate_identifier = true;
            }
            facts.identifiers.push(identifier);
            facts.maximum_identifier = facts.maximum_identifier.max(identifier);
        }
    }
    Ok(facts)
}

struct RegistryVisitor<'source> {
    source: &'source Package,
    components: Vec<ComponentFact>,
    uuids: Vec<UuidFact>,
    external_references: Vec<ExternalFact>,
    data_owners: Vec<DataOwnerFact>,
    ambiguous_identifiers: Vec<u64>,
    root_data_map_identifier: Option<u64>,
    maximum_identifier: u64,
    invalid: bool,
    allow_unmapped_current: bool,
}

impl<'source> RegistryVisitor<'source> {
    fn new(source: &'source Package, allow_unmapped_current: bool) -> Self {
        Self {
            source,
            components: Vec::new(),
            uuids: Vec::new(),
            external_references: Vec::new(),
            data_owners: Vec::new(),
            ambiguous_identifiers: Vec::new(),
            root_data_map_identifier: None,
            maximum_identifier: 0,
            invalid: false,
            allow_unmapped_current,
        }
    }

    fn push<T>(items: &mut Vec<T>, item: T) -> core::result::Result<(), RewriteError> {
        items
            .try_reserve(1)
            .map_err(|_| RewriteError::allocation(size_of::<T>()))?;
        items.push(item);
        Ok(())
    }

    fn record(&mut self, identifier: u64) {
        self.maximum_identifier = self.maximum_identifier.max(identifier);
        if identifier == 0 {
            self.invalid = true;
        }
    }

    fn finish(
        self,
        source: MetadataSource<'source>,
        inspection: PackageMetadataInspection,
        physical: PhysicalFacts,
    ) -> Result<MetadataRegistry<'source>> {
        if self.invalid || physical.duplicate_identifier {
            return Err(MetadataError::kind(FailureKind::Conflict));
        }
        for (index, binding) in self.uuids.iter().enumerate() {
            if self.uuids[index + 1..]
                .iter()
                .any(|other| other.uuid == binding.uuid)
            {
                return Err(MetadataError::kind(FailureKind::Conflict));
            }
        }
        let mut component_ids = Vec::new();
        let mut locators = Vec::new();
        for component in self.components.iter().filter(|item| item.current) {
            if component.component_index.is_none() && !self.allow_unmapped_current {
                return Err(MetadataError::invalid());
            }
            if component_ids.contains(&component.identifier)
                || locators.contains(&component.effective_locator)
            {
                return Err(MetadataError::kind(FailureKind::AmbiguousRoute));
            }
            component_ids.push(component.identifier);
            locators.push(component.effective_locator.clone());
        }
        let maximum_identifier = physical
            .maximum_identifier
            .max(self.maximum_identifier)
            .max(inspection.last_object_identifier());
        Ok(MetadataRegistry {
            source,
            components: self.components,
            uuids: self.uuids,
            external_references: self.external_references,
            data_owners: self.data_owners,
            ambiguous_identifiers: self.ambiguous_identifiers,
            root_data_map_identifier: self.root_data_map_identifier,
            last_object_identifier: inspection.last_object_identifier(),
            maximum_identifier,
            physical_identifiers: physical.identifiers,
            report: inspection.report(),
        })
    }
}

impl PackageMetadataVisitor for RegistryVisitor<'_> {
    fn visit_unknown_field(&mut self) -> core::result::Result<(), RewriteError> {
        self.invalid = true;
        Ok(())
    }

    fn visit_component(
        &mut self,
        component: ComponentDescriptor<'_>,
    ) -> core::result::Result<(), RewriteError> {
        self.record(component.identifier());
        let component_index = find_physical_descriptor(
            self.source,
            component.preferred_locator(),
            component.effective_locator(),
        );
        if component.is_current() && component_index.is_none() && !self.allow_unmapped_current {
            self.invalid = true;
        }
        Self::push(
            &mut self.components,
            ComponentFact {
                component_index,
                identifier: component.identifier(),
                preferred_locator: component.preferred_locator().to_owned(),
                effective_locator: component.effective_locator().to_owned(),
                current: component.is_current(),
            },
        )
    }

    fn visit_object_uuid(
        &mut self,
        binding: ObjectUuidDescriptor<'_>,
    ) -> core::result::Result<(), RewriteError> {
        self.record(binding.object_identifier());
        let component = binding.component();
        Self::push(
            &mut self.uuids,
            UuidFact {
                component_index: find_physical_descriptor(
                    self.source,
                    component.preferred_locator(),
                    component.effective_locator(),
                ),
                object_identifier: binding.object_identifier(),
                uuid: binding.uuid(),
                current: component.is_current(),
            },
        )
    }

    fn visit_external_reference(
        &mut self,
        reference: ExternalReferenceDescriptor<'_>,
    ) -> core::result::Result<(), RewriteError> {
        self.record(reference.target_component_identifier());
        if let Some(identifier) = reference.object_identifier() {
            self.record(identifier);
        }
        let source = reference.source();
        Self::push(
            &mut self.external_references,
            ExternalFact {
                source_component_index: find_physical_descriptor(
                    self.source,
                    source.preferred_locator(),
                    source.effective_locator(),
                ),
                source_identifier: source.identifier(),
                target_component_identifier: reference.target_component_identifier(),
                object_identifier: reference.object_identifier(),
                is_weak: reference.is_weak(),
                current: source.is_current(),
                versioned: reference.is_versioned(),
            },
        )
    }

    fn visit_data_reference_owner(
        &mut self,
        owner: DataReferenceOwnerDescriptor<'_>,
    ) -> core::result::Result<(), RewriteError> {
        if owner.has_unknown_fields() {
            self.invalid = true;
        }
        self.record(owner.data_identifier());
        self.record(owner.object_identifier());
        Self::push(
            &mut self.data_owners,
            DataOwnerFact {
                data_identifier: owner.data_identifier(),
                object_identifier: owner.object_identifier(),
            },
        )
    }

    fn visit_ambiguous_object_identifier(
        &mut self,
        _component: ComponentDescriptor<'_>,
        identifier: u64,
    ) -> core::result::Result<(), RewriteError> {
        self.record(identifier);
        Self::push(&mut self.ambiguous_identifiers, identifier)
    }

    fn visit_data_metadata_map(
        &mut self,
        object_identifier: u64,
        has_unknown_fields: bool,
    ) -> core::result::Result<(), RewriteError> {
        self.invalid |= has_unknown_fields;
        self.record(object_identifier);
        if self
            .root_data_map_identifier
            .replace(object_identifier)
            .is_some()
        {
            self.invalid = true;
        }
        Ok(())
    }
}

/// Inspect one strict metadata source under caller-supplied aggregate codec
/// limits. Unknown metadata is rejected before facts are returned.
pub(super) fn inspect(source: &Package, options: RewriteOptions) -> Result<MetadataRegistry<'_>> {
    inspect_with_policy(source, options, false)
}

pub(super) fn inspect_cross_component_read(
    source: &Package,
    options: RewriteOptions,
) -> Result<MetadataRegistry<'_>> {
    inspect_with_policy(source, options, true)
}

fn inspect_with_policy(
    source: &Package,
    options: RewriteOptions,
    allow_unmapped_current: bool,
) -> Result<MetadataRegistry<'_>> {
    let metadata = strict_source(source)?;
    let physical = physical_facts(source)?;
    let mut visitor = RegistryVisitor::new(source, allow_unmapped_current);
    let inspection = inspect_package_metadata_with_visitor(metadata.payload, options, &mut visitor)
        .map_err(map_rewrite_error)?;
    visitor.finish(metadata, inspection, physical)
}

/// Strictly inspect every physical ArchiveInfo header under the deletion
/// policy. Call this before culling so opaque inbound owner edges fail closed.
pub(super) fn reject_unknown_archive_metadata(
    source: &Package,
    limits: ArchiveLimits,
) -> Result<()> {
    struct Noop;
    impl ArchiveReferenceVisitor for Noop {
        fn visit_reference(
            &mut self,
            _occurrence: litchi_iwa_core::ArchiveReferenceOccurrence,
        ) -> litchi_iwa_core::Result<()> {
            Ok(())
        }
    }
    for component in source.state.components.catalog().iter() {
        for object in &component.archive().objects {
            object
                .inspect_references_with_policy_and_limits(
                    &mut Noop,
                    ArchiveReferencePolicy::RejectUnknownMetadata,
                    limits,
                )
                .map_err(|_| MetadataError::invalid())?;
        }
    }
    Ok(())
}

fn find_physical_component(source: &Package, locator: &str) -> Option<usize> {
    let mut found = None;
    for (index, component) in source.state.components.catalog().iter().enumerate() {
        if metadata::normalized_locator(component.name()) != locator {
            continue;
        }
        if found.replace(index).is_some() {
            return None;
        }
    }
    found
}

fn find_physical_descriptor(
    source: &Package,
    preferred_locator: &str,
    effective_locator: &str,
) -> Option<usize> {
    let preferred = find_physical_component(source, preferred_locator);
    let effective = find_physical_component(source, effective_locator);
    match (preferred, effective) {
        // An explicit locator cannot silently redirect a descriptor while
        // leaving a different preferred physical member in the package.  A
        // caller must repair that ambiguity before a metadata transition.
        (Some(preferred), Some(effective)) if preferred != effective => None,
        (_, effective) => effective,
    }
}

pub(super) fn prepare_additions<'source, 'batch>(
    facts: &MetadataRegistry<'source>,
    batch: AdditionSaveTokenBatch<'batch>,
    options: RewriteOptions,
) -> Result<PreparedPackageMetadataAdditionSaveTokenRewrite<'source, 'batch>> {
    facts.validate_save_tokens(batch.save_tokens())?;
    prepare_package_metadata_additions_and_save_tokens(facts.payload(), batch, options)
        .map_err(map_rewrite_error)
}

pub(super) fn prepare_removals<'source, 'batch>(
    facts: &MetadataRegistry<'source>,
    batch: RemovalSaveTokenBatch<'batch>,
    options: RewriteOptions,
) -> Result<PreparedPackageMetadataRemovalSaveTokenRewrite<'source, 'batch>> {
    facts.validate_save_tokens(batch.save_tokens())?;
    prepare_package_metadata_removals_and_save_tokens(facts.payload(), batch, options)
        .map_err(map_rewrite_error)
}

pub(super) fn prepare_combined<'source, 'batch>(
    facts: &MetadataRegistry<'source>,
    batch: CombinedSaveTokenBatch<'batch>,
    options: RewriteOptions,
) -> Result<PreparedCombinedTransition<'source, 'batch>> {
    facts.validate_save_tokens(batch.save_tokens())?;
    prepare_package_metadata_combined_additions_and_removals_and_save_tokens(
        facts.payload(),
        batch,
        options,
    )
    .map_err(map_rewrite_error)
}

pub(super) fn prepare_save_tokens<'source, 'batch>(
    facts: &MetadataRegistry<'source>,
    batch: SaveTokenBatch<'batch>,
    options: RewriteOptions,
) -> Result<PreparedPackageMetadataSaveTokenRewrite<'source, 'batch>> {
    facts.validate_save_tokens(batch)?;
    prepare_package_metadata_save_tokens(facts.payload(), batch, options).map_err(map_rewrite_error)
}

fn map_rewrite_error(error: RewriteError) -> MetadataError {
    if let Some(amount) = error.allocation_request() {
        return MetadataError::allocation(amount);
    }
    if let Some(limit) = error.resource_limit() {
        return match limit {
            RewriteLimit::InputBytes { observed, maximum }
            | RewriteLimit::OutputBytes { observed, maximum }
            | RewriteLimit::Fields { observed, maximum }
            | RewriteLimit::Work { observed, maximum }
            | RewriteLimit::Components { observed, maximum }
            | RewriteLimit::References { observed, maximum }
            | RewriteLimit::Additions { observed, maximum } => {
                MetadataError::limit(observed, maximum)
            },
            RewriteLimit::Nesting { observed, maximum } => {
                MetadataError::limit(observed as usize, maximum as usize)
            },
            _ => MetadataError::invalid(),
        };
    }
    match error.invalid_reason() {
        Some(
            InvalidReason::ComponentMismatch
            | InvalidReason::DuplicateSelector
            | InvalidReason::VersionedComponent
            | InvalidReason::SaveTokenMismatch
            | InvalidReason::DuplicateSaveToken
            | InvalidReason::ExistingObjectCollision
            | InvalidReason::ExistingUuidCollision
            | InvalidReason::ExistingReferenceCollision
            | InvalidReason::ConflictingWeakness
            | InvalidReason::RemovalMismatch
            | InvalidReason::VersionedRemoval
            | InvalidReason::CrossComponentRemoval,
        ) => MetadataError::kind(FailureKind::Conflict),
        Some(InvalidReason::RemovalNotFound) => MetadataError::kind(FailureKind::MissingRoute),
        Some(InvalidReason::InvalidIdentifier | InvalidReason::InvalidUuid) => {
            MetadataError::invalid()
        },
        Some(_) | None => MetadataError::invalid(),
    }
}

fn fingerprint(bytes: &[u8], seed: u64) -> u64 {
    bytes.iter().fold(seed, |hash, byte| {
        (hash ^ u64::from(*byte)).wrapping_mul(0x1000_0000_01b3)
    })
}

const fn mix(mut value: u64) -> u64 {
    value ^= value >> 30;
    value = value.wrapping_mul(0xbf58_476d_1ce4_e5b9);
    value ^= value >> 27;
    value = value.wrapping_mul(0x94d0_49bb_1331_11eb);
    value ^ (value >> 31)
}
