//! Strict, bounded PackageMetadata support for the Pop-Up Menu owner.
//!
//! The cell-control transaction has a deliberately small metadata surface:
//! one unique `Index/Metadata.iwa` message, a set of exact current component
//! selectors, and object UUID ownership records for objects that are created
//! or culled by the transaction.  This module keeps that census separate from
//! the native table/list rewrite so the owner can preflight both phases before
//! publishing a ZIP candidate.

use core::mem::size_of;

use litchi_iwa_protos::package_metadata_codec::{
    AdditionSaveTokenBatch, ComponentDescriptor, ComponentSelector, DataReferenceOwnerDescriptor,
    ExternalReferenceDescriptor, InvalidReason, ObjectUuidAddition, ObjectUuidDescriptor,
    ObjectUuidRemoval, PackageMetadataInspection, PackageMetadataVisitor,
    PreparedPackageMetadataAdditionSaveTokenRewrite,
    PreparedPackageMetadataRemovalSaveTokenRewrite, PreparedPackageMetadataSaveTokenRewrite,
    RemovalSaveTokenBatch, RewriteError, RewriteLimit, RewriteOptions, RewriteReport,
    SaveTokenBatch, UuidBits, inspect_package_metadata_with_visitor,
    prepare_package_metadata_additions_and_save_tokens,
    prepare_package_metadata_removals_and_save_tokens, prepare_package_metadata_save_tokens,
};

use super::{Package, metadata};

const METADATA_MESSAGE_TYPE: u32 = 11_006;

/// A content-free classification of a metadata failure.
///
/// The parent transaction maps this vocabulary to its own `Path` and public
/// limit type.  No generated metadata or native object type crosses the
/// Numbers semantic boundary.
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

/// A redacted metadata failure returned to the Pop-Up Menu owner.
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

/// The unique metadata payload and its native archive route.
#[derive(Debug, Clone, Copy)]
pub(super) struct MetadataSource<'source> {
    pub(super) route: metadata::MessageRoute,
    pub(super) payload: &'source [u8],
}

/// Locate exactly one current `Metadata.iwa` object carrying type 11006.
///
/// `metadata::unique_message_route` rejects duplicates and rejects a message
/// with the right type in any member other than `Index/Metadata.iwa`.  The
/// extra payload check here keeps this helper strict if the route helper is
/// ever widened.
pub(super) fn strict_source(source: &Package) -> Result<MetadataSource<'_>> {
    let route = metadata::unique_message_route(source).ok_or_else(MetadataError::kind_missing)?;
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

impl MetadataError {
    const fn kind_missing() -> Self {
        Self::kind(FailureKind::MissingRoute)
    }
}

/// A current or versioned component descriptor, retaining the exact effective
/// locator borrowed from the metadata payload.
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
    component_identifier: u64,
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

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct DataOwnerFact {
    data_identifier: u64,
    object_identifier: u64,
}

/// The strict current/versioned metadata registry census used by popup edits.
///
/// All slices and locators borrow the one inspected metadata payload.  The
/// facts can therefore be retained through a prepared codec transition
/// without reconstructing a normalized locator from a ZIP member name.
#[derive(Debug)]
pub(super) struct RegistryFacts<'source> {
    source: MetadataSource<'source>,
    components: Vec<ComponentFact>,
    uuids: Vec<UuidFact>,
    external_references: Vec<ExternalFact>,
    data_owners: Vec<DataOwnerFact>,
    ambiguous_identifiers: Vec<u64>,
    root_data_map_identifier: Option<u64>,
    last_object_identifier: u64,
    maximum_identifier: u64,
    physical_alias: bool,
    report: RewriteReport,
}

impl<'source> RegistryFacts<'source> {
    #[must_use]
    pub(super) const fn payload(&self) -> &'source [u8] {
        self.source.payload
    }

    #[must_use]
    pub(super) const fn route(&self) -> metadata::MessageRoute {
        self.source.route
    }

    #[must_use]
    pub(super) const fn last_object_identifier(&self) -> u64 {
        self.last_object_identifier
    }

    #[must_use]
    pub(super) const fn report(&self) -> RewriteReport {
        self.report
    }

    /// Whether physical object identifiers contain an exact cross-component
    /// alias.  A popup mutation should reject this unless it rewrites every
    /// alias, which the first owner slice does not do.
    #[must_use]
    pub(super) const fn has_physical_alias(&self) -> bool {
        self.physical_alias
    }

    /// Return the one current metadata selector for a physical component.
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

    /// Build exact current selectors for a deduplicated set of component
    /// indices.  The caller can pass the result directly to
    /// `SaveTokenBatch`; its locators remain borrowed from this census.
    pub(super) fn selectors_for_components(
        &self,
        component_indices: &[usize],
    ) -> Result<Vec<ComponentSelector<'_>>> {
        let mut selectors: Vec<ComponentSelector<'_>> = Vec::new();
        selectors
            .try_reserve_exact(component_indices.len())
            .map_err(|_| MetadataError::allocation(component_indices.len()))?;
        for &component_index in component_indices {
            let selector = self.selector(component_index)?;
            if selectors
                .iter()
                .any(|existing| existing.identifier() == selector.identifier())
            {
                continue;
            }
            selectors.push(selector);
        }
        if selectors.is_empty() {
            return Err(MetadataError::kind(FailureKind::InvalidSource));
        }
        Ok(selectors)
    }

    /// Validate that a codec save-token batch uses the exact effective
    /// current selectors discovered in this census.  The codec repeats this
    /// check during prepare; keeping the cheap preflight here lets the owner
    /// reject a stale or preferred-locator selector before staging a native
    /// candidate.
    pub(super) fn validate_save_tokens(&self, batch: SaveTokenBatch<'_>) -> Result<()> {
        if batch.selectors().is_empty() {
            return Err(MetadataError::invalid());
        }
        for (index, selector) in batch.selectors().iter().copied().enumerate() {
            let mut current = 0usize;
            for component in &self.components {
                if component.identifier != selector.identifier() {
                    continue;
                }
                if component.current && component.effective_locator == selector.locator() {
                    current = current.saturating_add(1);
                }
            }
            // Versioned component records are retained raw and do not
            // participate in current selector matching. The codec still
            // validates their known framing during its source scan.
            if current != 1 {
                return Err(MetadataError::kind(FailureKind::Conflict));
            }
            if batch
                .selectors()
                .iter()
                .skip(index + 1)
                .any(|other| other == &selector)
            {
                return Err(MetadataError::kind(FailureKind::Conflict));
            }
        }
        Ok(())
    }

    /// Require that every component participating in the initial owner slice
    /// is the same current component.
    pub(super) fn require_single_current_component(
        &self,
        component_indices: &[usize],
    ) -> Result<usize> {
        let first = *component_indices
            .first()
            .ok_or_else(MetadataError::invalid)?;
        for &component_index in component_indices.iter().skip(1) {
            if component_index != first {
                return Err(MetadataError::kind(FailureKind::Unsupported));
            }
        }
        let _ = self.selector(first)?;
        Ok(first)
    }

    /// Require one exact current UUID owner and no hostile alternate registry
    /// namespace for the same native object.
    pub(super) fn require_current_uuid(
        &self,
        component_index: usize,
        object_identifier: u64,
    ) -> Result<UuidBits> {
        if self.has_non_uuid_owner(object_identifier) {
            return Err(MetadataError::kind(FailureKind::Conflict));
        }
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
            if binding.component_index != Some(component_index) {
                return Err(MetadataError::kind(FailureKind::Conflict));
            }
            if current.replace(binding.uuid).is_some() {
                return Err(MetadataError::kind(FailureKind::Conflict));
            }
        }
        if versioned {
            return Err(MetadataError::kind(FailureKind::VersionedOwnership));
        }
        current.ok_or_else(|| MetadataError::kind(FailureKind::MissingRoute))
    }

    /// Prove one current, effective-locator-resolved external edge for a
    /// cross-component native graph.  Component UUID records are deliberately
    /// not used as a substitute: sidecar members commonly have no object UUID
    /// bindings, while the owning CalculationEngine component carries the
    /// authoritative external edge. Native producers use both object-specific
    /// records and component-level records (`object_identifier = None`), so
    /// either shape may cover a requested object. A covering edge must be
    /// current, unversioned, and have the exact weak/reference shape requested
    /// by the caller; duplicate or conflicting records fail closed.
    pub(super) fn require_external_edge(
        &self,
        source_component_index: usize,
        target_component_index: usize,
        object_identifier: Option<u64>,
        is_weak: Option<bool>,
    ) -> Result<()> {
        let source = self.selector(source_component_index)?;
        let target_identifier = match self.selector(target_component_index) {
            Ok(target) => target.identifier(),
            Err(error) => {
                // Some native table sidecars are physical IWA members but do
                // not have their own current ComponentInfo. Their one current
                // external edge uses the sidecar object's identifier as the
                // target component identifier. Admit that producer shape only
                // when there is no current or versioned component record to
                // contradict it; ambiguous/versioned selector failures must
                // not be normalized into this fallback.
                let object_identifier = object_identifier.ok_or(error)?;
                if self.components.iter().any(|component| {
                    component.component_index == Some(target_component_index)
                        || component.identifier == object_identifier
                }) {
                    return Err(error);
                }
                object_identifier
            },
        };
        let mut exact = 0usize;
        let mut conflicting = false;
        for reference in &self.external_references {
            if reference.source_component_index != Some(source_component_index)
                || reference.source_identifier != source.identifier()
                || reference.target_component_identifier != target_identifier
                || (reference.object_identifier != object_identifier
                    && reference.object_identifier.is_some())
            {
                continue;
            }
            if !reference.current || reference.versioned || reference.is_weak != is_weak {
                conflicting = true;
                continue;
            }
            exact = exact.saturating_add(1);
        }
        if conflicting || exact != 1 {
            return Err(MetadataError::kind(FailureKind::Conflict));
        }
        Ok(())
    }

    /// Reject any existing ownership record before appending a fresh UUID.
    pub(super) fn require_uuid_absent(&self, object_identifier: u64) -> Result<()> {
        if self
            .uuids
            .iter()
            .any(|binding| binding.object_identifier == object_identifier)
            || self.has_non_uuid_owner(object_identifier)
        {
            return Err(MetadataError::kind(FailureKind::Conflict));
        }
        Ok(())
    }

    /// Prepare the exact UUID addition record for a selected current component.
    pub(super) fn uuid_addition(
        &self,
        component_index: usize,
        identifier: FreshIdentifier,
    ) -> Result<ObjectUuidAddition<'_>> {
        self.require_single_current_component(core::slice::from_ref(&component_index))?;
        self.require_uuid_absent(identifier.identifier)?;
        Ok(ObjectUuidAddition::new(
            self.selector(component_index)?,
            identifier.identifier,
            identifier.uuid,
        ))
    }

    /// Prepare the exact current UUID removal record and prove that no known
    /// metadata namespace still points at the object.
    pub(super) fn uuid_removal(
        &self,
        component_index: usize,
        object_identifier: u64,
    ) -> Result<ObjectUuidRemoval<'_>> {
        let uuid = self.require_current_uuid(component_index, object_identifier)?;
        Ok(ObjectUuidRemoval::new(
            self.selector(component_index)?,
            object_identifier,
            uuid,
        ))
    }

    /// Allocate fresh object identifiers and deterministic nonzero UUID pairs.
    /// The starting point is above physical objects and all metadata-owned
    /// identifiers, not merely the advisory root watermark.
    pub(super) fn allocate_identifiers(&self, count: usize) -> Result<Vec<FreshIdentifier>> {
        let mut identifiers = Vec::new();
        identifiers
            .try_reserve_exact(count)
            .map_err(|_| MetadataError::allocation(count))?;
        let seed = fingerprint(self.payload(), 0xcbf2_9ce4_8422_2325);
        let mut next = self.maximum_identifier;
        for ordinal in 0..count {
            next = next
                .checked_add(1)
                .ok_or_else(|| MetadataError::kind(FailureKind::Limit))?;
            let ordinal = u64::try_from(ordinal).map_err(|_| MetadataError::invalid())?;
            let mut salt = 0u64;
            let uuid = loop {
                let upper = mix(seed.rotate_left(29)
                    ^ next.rotate_left(17)
                    ^ ordinal
                    ^ salt.rotate_left(7));
                let candidate = UuidBits::new(next, upper);
                if candidate.lower() != 0
                    && candidate.upper() != 0
                    && !self.uuids.iter().any(|binding| binding.uuid == candidate)
                    && !identifiers
                        .iter()
                        .any(|item: &FreshIdentifier| item.uuid == candidate)
                {
                    break candidate;
                }
                salt = salt
                    .checked_add(1)
                    .ok_or_else(|| MetadataError::kind(FailureKind::Limit))?;
            };
            identifiers.push(FreshIdentifier {
                identifier: next,
                uuid,
            });
        }
        Ok(identifiers)
    }

    fn has_non_uuid_owner(&self, identifier: u64) -> bool {
        self.external_references.iter().any(|reference| {
            reference.object_identifier == Some(identifier)
                || reference.target_component_identifier == identifier
        }) || self.data_owners.iter().any(|owner| {
            owner.object_identifier == identifier || owner.data_identifier == identifier
        }) || self.ambiguous_identifiers.contains(&identifier)
            || self.root_data_map_identifier == Some(identifier)
    }
}

/// A new native object identifier plus its metadata UUID pair.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct FreshIdentifier {
    pub(super) identifier: u64,
    pub(super) uuid: UuidBits,
}

/// Inspect the unique metadata payload with one caller-supplied aggregate
/// budget.  The returned report is the exact codec inspection report and must
/// be merged into the owner transaction budget once.
pub(super) fn inspect(source: &Package, options: RewriteOptions) -> Result<RegistryFacts<'_>> {
    inspect_with_policy(source, options, false)
}

/// Inspect metadata for the bounded read-only cross-component projection.
/// Native packages may list current file-backed Data components that are not
/// IWA members; retain them in the collision census without treating their
/// absence as mutation authority. Every component selected by an actual edge
/// is still required to resolve through `RegistryFacts::selector`.
pub(super) fn inspect_cross_component_read(
    source: &Package,
    options: RewriteOptions,
) -> Result<RegistryFacts<'_>> {
    inspect_with_policy(source, options, true)
}

fn inspect_with_policy(
    source: &Package,
    options: RewriteOptions,
    allow_unmapped_current: bool,
) -> Result<RegistryFacts<'_>> {
    let metadata = strict_source(source)?;
    let physical = physical_identifiers(source)?;
    let mut visitor = RegistryVisitor::new(source, allow_unmapped_current);
    let inspection = inspect_package_metadata_with_visitor(metadata.payload, options, &mut visitor)
        .map_err(map_rewrite_error)?;
    visitor.finish(metadata, inspection, physical)
}

#[derive(Debug, Default)]
struct PhysicalFacts {
    maximum_identifier: u64,
    duplicate_identifier: bool,
}

fn physical_identifiers(source: &Package) -> Result<PhysicalFacts> {
    let mut seen = Vec::new();
    let mut facts = PhysicalFacts::default();
    for component in source.state.components.catalog().iter() {
        for object in &component.archive().objects {
            let identifier = object
                .archive_info
                .identifier
                .filter(|identifier| *identifier != 0)
                .ok_or_else(MetadataError::invalid)?;
            seen.try_reserve(1)
                .map_err(|_| MetadataError::allocation(1))?;
            if seen.contains(&identifier) {
                facts.duplicate_identifier = true;
            }
            seen.push(identifier);
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

    fn push<T>(items: &mut Vec<T>, item: T) -> Result<()> {
        items
            .try_reserve(1)
            .map_err(|_| MetadataError::allocation(size_of::<T>()))?;
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
    ) -> Result<RegistryFacts<'source>> {
        if self.invalid {
            return Err(MetadataError::invalid());
        }
        // A UUID is an ownership key, not merely an opaque decoration.  Two
        // metadata records carrying the same UUID would make a later native
        // mutation ambiguous even when their numeric object identifiers differ
        // (the codec's identifier collision checks cannot detect this value
        // alias).  Reject the whole census before any candidate is staged.
        for (index, binding) in self.uuids.iter().enumerate() {
            if self.uuids[index + 1..]
                .iter()
                .any(|other| other.uuid == binding.uuid)
            {
                return Err(MetadataError::kind(FailureKind::Conflict));
            }
        }
        let mut current_identifiers = Vec::new();
        let mut current_locators = Vec::new();
        for component in self.components.iter().filter(|component| component.current) {
            if component.component_index.is_none() && !self.allow_unmapped_current {
                return Err(MetadataError::invalid());
            }
            current_identifiers
                .try_reserve(1)
                .map_err(|_| MetadataError::allocation(1))?;
            current_locators
                .try_reserve(1)
                .map_err(|_| MetadataError::allocation(1))?;
            if current_identifiers.contains(&component.identifier)
                || current_locators.contains(&component.effective_locator)
            {
                return Err(MetadataError::kind(FailureKind::AmbiguousRoute));
            }
            current_identifiers.push(component.identifier);
            current_locators.push(component.effective_locator.clone());
        }
        let maximum_identifier = physical
            .maximum_identifier
            .max(self.maximum_identifier)
            .max(inspection.last_object_identifier());
        Ok(RegistryFacts {
            source,
            components: self.components,
            uuids: self.uuids,
            external_references: self.external_references,
            data_owners: self.data_owners,
            ambiguous_identifiers: self.ambiguous_identifiers,
            root_data_map_identifier: self.root_data_map_identifier,
            last_object_identifier: inspection.last_object_identifier(),
            maximum_identifier,
            physical_alias: physical.duplicate_identifier,
            report: inspection.report(),
        })
    }
}

impl PackageMetadataVisitor for RegistryVisitor<'_> {
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
        .map_err(|error| RewriteError::allocation(error.allocation))
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
                component_identifier: component.identifier(),
                object_identifier: binding.object_identifier(),
                uuid: binding.uuid(),
                current: component.is_current(),
            },
        )
        .map_err(|error| RewriteError::allocation(error.allocation))
    }

    fn visit_external_reference(
        &mut self,
        reference: ExternalReferenceDescriptor<'_>,
    ) -> core::result::Result<(), RewriteError> {
        self.record(reference.target_component_identifier());
        if let Some(identifier) = reference.object_identifier() {
            self.record(identifier);
        }
        let component = reference.source();
        Self::push(
            &mut self.external_references,
            ExternalFact {
                source_component_index: find_physical_descriptor(
                    self.source,
                    component.preferred_locator(),
                    component.effective_locator(),
                ),
                source_identifier: component.identifier(),
                target_component_identifier: reference.target_component_identifier(),
                object_identifier: reference.object_identifier(),
                is_weak: reference.is_weak(),
                current: component.is_current(),
                versioned: reference.is_versioned(),
            },
        )
        .map_err(|error| RewriteError::allocation(error.allocation))
    }

    fn visit_data_reference_owner(
        &mut self,
        owner: DataReferenceOwnerDescriptor<'_>,
    ) -> core::result::Result<(), RewriteError> {
        self.record(owner.data_identifier());
        self.record(owner.object_identifier());
        Self::push(
            &mut self.data_owners,
            DataOwnerFact {
                data_identifier: owner.data_identifier(),
                object_identifier: owner.object_identifier(),
            },
        )
        .map_err(|error| RewriteError::allocation(error.allocation))
    }

    fn visit_ambiguous_object_identifier(
        &mut self,
        _component: ComponentDescriptor<'_>,
        identifier: u64,
    ) -> core::result::Result<(), RewriteError> {
        self.record(identifier);
        Self::push(&mut self.ambiguous_identifiers, identifier)
            .map_err(|error| RewriteError::allocation(error.allocation))
    }

    fn visit_data_metadata_map(
        &mut self,
        object_identifier: u64,
        _has_unknown_fields: bool,
    ) -> core::result::Result<(), RewriteError> {
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

fn find_physical_component(source: &Package, locator: &str) -> Option<usize> {
    let mut found = None;
    for (index, component) in source.state.components.catalog().iter().enumerate() {
        if metadata::normalized_locator(component.name()) != locator {
            continue;
        }
        if found.is_some() {
            return None;
        }
        found = Some(index);
    }
    found
}

fn find_physical_descriptor(
    source: &Package,
    preferred_locator: &str,
    effective_locator: &str,
) -> Option<usize> {
    // ComponentInfo.locator (the effective locator) is authoritative. Native
    // producers commonly retain a generic preferred locator while the exact
    // current member is named by an explicit locator; both physical spellings
    // may legitimately exist. When they differ, require the effective member
    // and never fall back to the preferred spelling.
    if preferred_locator != effective_locator {
        find_physical_component(source, effective_locator)
    } else {
        find_physical_component(source, preferred_locator)
    }
}

/// Prepare a strict addition + root/save-token transition.  The returned
/// codec object is output-free until its `execute` method is called with the
/// exact execution requirements.
pub(super) fn prepare_additions<'source, 'batch>(
    facts: &RegistryFacts<'source>,
    batch: AdditionSaveTokenBatch<'batch>,
    options: RewriteOptions,
) -> Result<PreparedPackageMetadataAdditionSaveTokenRewrite<'source, 'batch>> {
    facts.validate_save_tokens(batch.save_tokens())?;
    prepare_package_metadata_additions_and_save_tokens(facts.payload(), batch, options)
        .map_err(map_rewrite_error)
}

/// Prepare a strict UUID-removal + root/save-token transition.
pub(super) fn prepare_removals<'source, 'batch>(
    facts: &RegistryFacts<'source>,
    batch: RemovalSaveTokenBatch<'batch>,
    options: RewriteOptions,
) -> Result<PreparedPackageMetadataRemovalSaveTokenRewrite<'source, 'batch>> {
    facts.validate_save_tokens(batch.save_tokens())?;
    prepare_package_metadata_removals_and_save_tokens(facts.payload(), batch, options)
        .map_err(map_rewrite_error)
}

/// Prepare a root + selected current-component save-token transition with no
/// registry change.
pub(super) fn prepare_save_tokens<'source, 'batch>(
    facts: &RegistryFacts<'source>,
    batch: SaveTokenBatch<'batch>,
    options: RewriteOptions,
) -> Result<PreparedPackageMetadataSaveTokenRewrite<'source, 'batch>> {
    facts.validate_save_tokens(batch)?;
    prepare_package_metadata_save_tokens(facts.payload(), batch, options).map_err(map_rewrite_error)
}

/// Map a prepared codec failure without leaking the codec's generated/error
/// vocabulary through the package owner.
pub(super) fn map_rewrite_error(error: RewriteError) -> MetadataError {
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
