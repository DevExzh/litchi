//! Owner creation and exact-source rewrite internals for Pages body-table hidden axes.
//!
//! This private child owns the candidate-building path. The parent module keeps
//! the public selector/edit/patch API and graph proof, while this module
//! exposes one bounded run boundary for the rewrite transaction.

use std::num::NonZeroU64;
use std::sync::Arc;

use litchi_iwa_archive::{SourceCatalog, package::EntryEdit};
use litchi_iwa_common::WireLimits;
use litchi_iwa_core::archive::ObjectReferenceTransition;
use litchi_iwa_core::{
    Archive, ArchiveObject, ArchiveReferenceOccurrence, ArchiveReferencePolicy,
    ArchiveReferenceVisitor, CanonicalObjectReferenceField, CanonicalObjectReferenceFields,
    RawMessage,
};
use litchi_iwa_protos::package_metadata_codec::{
    AdditionSaveTokenBatch, Batch as MetadataBatch, ComponentSelector, ObjectUuidAddition,
    PackageMetadataVisitor, RewriteError as MetadataRewriteError,
    RewriteOptions as MetadataRewriteOptions, SaveTokenBatch, UuidBits,
    inspect_package_metadata_with_visitor, prepare_package_metadata_additions_and_save_tokens,
};

use super::{
    AxisIndex, BodyTableHiddenAxesError, BodyTableHiddenAxesLimitKind, CURRENT_MESSAGE_VERSIONS,
    FILTER_SET_MESSAGE_TYPE, FORMULA_OWNER_MESSAGE_TYPE, Graph,
    HIDDEN_STATE_FORMULA_OWNER_MESSAGE_TYPE, HiddenAxes, LEGACY_UID_MAP_MESSAGE_TYPE,
    METADATA_ENTRY_NAME, METADATA_MESSAGE_TYPE, MODEL_COLUMN_REFERENCE_PATH,
    MODEL_ROW_REFERENCE_PATH, MetadataRoute, OWNER_CREATION_FIELD_INFOS, OWNER_CREATION_MESSAGES,
    OWNER_CREATION_OBJECTS, ObjectLocation, Package, UID_MAP_MESSAGE_TYPE, UidIndex, charge_codec,
    charge_style_codec, codec, codec_options, formula_owner_for, formula_owner_messages,
    global_objects, map_archive_error, map_codec_error, map_core_error, map_lock_error,
    map_package_error, map_page_layout_error, map_style_codec_error, map_uid_codec_error,
    metadata_route, metadata_scan_work, object_location, page_layout, style_codec,
    style_codec_options, table_lock, uid_codec, validate_aggregate_formula_owner_component,
    validate_aggregate_formula_owner_metadata, validate_aggregate_formula_owner_provenance,
    validate_empty_formula_owner_dependencies, validate_message_metadata,
    validate_no_reference_metadata, validate_reference_shape, validate_uuid,
};

/// Execute the verified hidden-axis rewrite and return its reopened package.
pub(super) fn run(
    source: &Package,
    graph: &Graph,
    axes: &HiddenAxes,
    budget: &mut table_lock::WireBudget,
) -> Result<(Package, usize, Arc<[u64]>), BodyTableHiddenAxesError> {
    let result = rewrite(source, graph, axes, budget)?;
    Ok((
        result.package,
        result.touched_components,
        result.added_object_ids,
    ))
}

#[derive(Clone, Copy)]
struct UidMapDimensions {
    identifier: u64,
    columns: u32,
    rows: u32,
}

#[derive(Clone, Copy)]
struct OwnerCreationIds {
    column_formula: u64,
    row_formula: u64,
    column_filter: u64,
    row_filter: u64,
}

impl OwnerCreationIds {
    const COUNT: usize = OWNER_CREATION_OBJECTS;

    const fn as_array(self) -> [u64; Self::COUNT] {
        [
            self.column_formula,
            self.row_formula,
            self.column_filter,
            self.row_filter,
        ]
    }
}

struct MetadataPlan<'source> {
    route: MetadataRoute,
    payload: &'source [u8],
    selector: ComponentSelector<'source>,
    expected_last_identifier: u64,
    new_last_identifier: u64,
    uuids: [UuidBits; OWNER_CREATION_OBJECTS],
}

struct MetadataProbe<'source> {
    route: MetadataRoute,
    payload: &'source [u8],
    selector: ComponentSelector<'source>,
    expected_last_identifier: u64,
    maximum_identifier: u64,
}

struct OwnerCreationPlan<'source> {
    ids: OwnerCreationIds,
    object_ids: Arc<[u64]>,
    owner: codec::HiddenStatesOwnerSnapshot,
    active_uuid: codec::UuidSnapshot,
    formula_columns_ref: codec::ReferenceSnapshot,
    formula_rows_ref: codec::ReferenceSnapshot,
    metadata: Option<MetadataPlan<'source>>,
}

struct RewriteResult {
    package: Package,
    added_object_ids: Arc<[u64]>,
    touched_components: usize,
}

#[derive(Default)]
struct MaximumReferenceVisitor {
    maximum: u64,
    occurrences: usize,
}

impl ArchiveReferenceVisitor for MaximumReferenceVisitor {
    fn visit_reference(
        &mut self,
        occurrence: ArchiveReferenceOccurrence,
    ) -> litchi_iwa_core::Result<()> {
        self.maximum = self
            .maximum
            .max(occurrence.object_identifier)
            .max(occurrence.referenced_identifier);
        self.occurrences = self.occurrences.checked_add(1).ok_or({
            litchi_iwa_core::Error::InvalidArchive {
                offset: 0,
                reason: "reference census overflow",
            }
        })?;
        Ok(())
    }
}

/// Return the largest identifier occurring in either the physical object set
/// or its complete archive-header reference inventory.  Object identifiers
/// are not sufficient for allocation: a dangling reference in an unrelated
/// object is still a collision with a newly introduced object ID.
fn maximum_physical_identifier(
    package: &Package,
    objects: &[ObjectLocation],
    budget: &mut table_lock::WireBudget,
) -> Result<u64, BodyTableHiddenAxesError> {
    let archive_limits = package
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let mut work = 0usize;
    for location in objects {
        let object = package
            .state
            .source
            .components()
            .get_index(location.component_index)
            .and_then(|component| component.archive().objects.get(location.object_index))
            .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
        for info in &object.archive_info.message_infos {
            work = work.checked_add(metadata_scan_work(info)?).ok_or(
                BodyTableHiddenAxesError::LimitExceeded {
                    kind: BodyTableHiddenAxesLimitKind::TransactionWork,
                    observed: u64::MAX,
                    maximum: u64::try_from(budget.wire_limits().max_rewrite_work())
                        .unwrap_or(u64::MAX),
                },
            )?;
        }
    }
    let mut maximum = objects
        .iter()
        .map(|location| location.identifier)
        .max()
        .unwrap_or(0);
    for location in objects {
        let object = package
            .state
            .source
            .components()
            .get_index(location.component_index)
            .and_then(|component| component.archive().objects.get(location.object_index))
            .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
        for info in &object.archive_info.message_infos {
            for identifier in info
                .object_references
                .iter()
                .chain(info.data_references.iter())
            {
                maximum = maximum.max(*identifier);
            }
            for field in &info.field_infos {
                for identifier in field
                    .object_references
                    .iter()
                    .chain(field.data_references.iter())
                {
                    maximum = maximum.max(*identifier);
                }
            }
        }
        let mut references = MaximumReferenceVisitor::default();
        let occurrence_count = object
            .inspect_references_with_policy_and_limits(
                &mut references,
                ArchiveReferencePolicy::RejectUnknownMetadata,
                archive_limits,
            )
            .map_err(map_core_error)?;
        work =
            work.checked_add(occurrence_count)
                .ok_or(BodyTableHiddenAxesError::LimitExceeded {
                    kind: BodyTableHiddenAxesLimitKind::TransactionWork,
                    observed: u64::MAX,
                    maximum: u64::try_from(budget.wire_limits().max_rewrite_work())
                        .unwrap_or(u64::MAX),
                })?;
        maximum = maximum.max(references.maximum);
    }
    budget.charge_payload_work(work).map_err(map_lock_error)?;
    Ok(maximum)
}

fn normalized_locator(name: &str) -> &str {
    name.strip_prefix("Index/")
        .and_then(|name| name.strip_suffix(".iwa"))
        .unwrap_or(name)
}

struct MetadataCensus<'source> {
    target_locator: &'source str,
    maximum: u64,
    target_component_identifier: Option<u64>,
    target_component_matches: usize,
    target_object_identifiers: Vec<u64>,
    component_identifiers: Vec<u64>,
    object_bindings: Vec<(u64, u64)>,
    uuids: Vec<UuidBits>,
    unknown: bool,
}

impl<'source> MetadataCensus<'source> {
    fn new(
        target_locator: &'source str,
        target_object_capacity: usize,
        metadata_item_capacity: usize,
        component_capacity: usize,
    ) -> Result<Self, BodyTableHiddenAxesError> {
        let mut target_object_identifiers = Vec::new();
        target_object_identifiers
            .try_reserve_exact(target_object_capacity.max(metadata_item_capacity))
            .map_err(|_| BodyTableHiddenAxesError::Allocation {
                amount: target_object_capacity.max(metadata_item_capacity),
            })?;
        let mut component_identifiers = Vec::new();
        component_identifiers
            .try_reserve_exact(component_capacity.max(metadata_item_capacity))
            .map_err(|_| BodyTableHiddenAxesError::Allocation {
                amount: component_capacity.max(metadata_item_capacity),
            })?;
        let mut object_bindings = Vec::new();
        object_bindings
            .try_reserve_exact(metadata_item_capacity)
            .map_err(|_| BodyTableHiddenAxesError::Allocation {
                amount: metadata_item_capacity,
            })?;
        let mut uuids = Vec::new();
        uuids
            .try_reserve_exact(metadata_item_capacity)
            .map_err(|_| BodyTableHiddenAxesError::Allocation {
                amount: metadata_item_capacity,
            })?;
        Ok(Self {
            target_locator,
            maximum: 0,
            target_component_identifier: None,
            target_component_matches: 0,
            target_object_identifiers,
            component_identifiers,
            object_bindings,
            uuids,
            unknown: false,
        })
    }

    fn record(&mut self, identifier: u64) {
        self.maximum = self.maximum.max(identifier);
    }
}

impl PackageMetadataVisitor for MetadataCensus<'_> {
    fn visit_unknown_field(&mut self) -> Result<(), MetadataRewriteError> {
        self.unknown = true;
        Ok(())
    }

    fn visit_component(
        &mut self,
        component: litchi_iwa_protos::package_metadata_codec::ComponentDescriptor<'_>,
    ) -> Result<(), MetadataRewriteError> {
        self.record(component.identifier());
        if component.is_current() {
            self.component_identifiers.push(component.identifier());
        }
        if component.is_current() && component.effective_locator() == self.target_locator {
            self.target_component_matches = self.target_component_matches.saturating_add(1);
            self.target_component_identifier = Some(component.identifier());
        }
        Ok(())
    }

    fn visit_object_uuid(
        &mut self,
        binding: litchi_iwa_protos::package_metadata_codec::ObjectUuidDescriptor<'_>,
    ) -> Result<(), MetadataRewriteError> {
        self.record(binding.component().identifier());
        self.record(binding.object_identifier());
        if binding.component().is_current() {
            self.object_bindings.push((
                binding.component().identifier(),
                binding.object_identifier(),
            ));
            self.uuids.push(binding.uuid());
        }
        if binding.component().is_current()
            && self
                .target_component_identifier
                .is_some_and(|identifier| identifier == binding.component().identifier())
        {
            self.target_object_identifiers
                .push(binding.object_identifier());
        }
        Ok(())
    }

    fn visit_external_reference(
        &mut self,
        reference: litchi_iwa_protos::package_metadata_codec::ExternalReferenceDescriptor<'_>,
    ) -> Result<(), MetadataRewriteError> {
        self.record(reference.source().identifier());
        self.record(reference.target_component_identifier());
        if let Some(identifier) = reference.object_identifier() {
            self.record(identifier);
        }
        Ok(())
    }

    fn visit_data_reference(
        &mut self,
        reference: litchi_iwa_protos::package_metadata_codec::DataReferenceDescriptor<'_>,
    ) -> Result<(), MetadataRewriteError> {
        self.record(reference.component().identifier());
        self.record(reference.data_identifier());
        Ok(())
    }

    fn visit_data_reference_owner(
        &mut self,
        owner: litchi_iwa_protos::package_metadata_codec::DataReferenceOwnerDescriptor<'_>,
    ) -> Result<(), MetadataRewriteError> {
        self.record(owner.component().identifier());
        self.record(owner.data_identifier());
        self.record(owner.object_identifier());
        Ok(())
    }

    fn visit_ambiguous_object_identifier(
        &mut self,
        component: litchi_iwa_protos::package_metadata_codec::ComponentDescriptor<'_>,
        identifier: u64,
    ) -> Result<(), MetadataRewriteError> {
        self.record(component.identifier());
        self.record(identifier);
        Ok(())
    }

    fn visit_data_metadata_map(
        &mut self,
        object_identifier: u64,
        _has_unknown_fields: bool,
    ) -> Result<(), MetadataRewriteError> {
        self.record(object_identifier);
        Ok(())
    }
}

struct MetadataUuidCollisionVisitor {
    targets: [UuidBits; OWNER_CREATION_OBJECTS],
    found: bool,
    unknown: bool,
}

impl MetadataUuidCollisionVisitor {
    fn new(targets: [UuidBits; OWNER_CREATION_OBJECTS]) -> Self {
        Self {
            targets,
            found: false,
            unknown: false,
        }
    }
}

impl PackageMetadataVisitor for MetadataUuidCollisionVisitor {
    fn visit_unknown_field(&mut self) -> Result<(), MetadataRewriteError> {
        self.unknown = true;
        Ok(())
    }

    fn visit_object_uuid(
        &mut self,
        binding: litchi_iwa_protos::package_metadata_codec::ObjectUuidDescriptor<'_>,
    ) -> Result<(), MetadataRewriteError> {
        if self.targets.contains(&binding.uuid()) {
            self.found = true;
        }
        Ok(())
    }
}

fn metadata_options(source: &Package, bytes: usize, additions: usize) -> MetadataRewriteOptions {
    let physical_max = source.state.source.limits().max_iwa_stream_bytes();
    let addition_bytes = additions.saturating_mul(128);
    MetadataRewriteOptions::new(
        bytes.max(1).min(physical_max),
        bytes
            .saturating_add(addition_bytes)
            .saturating_add(256)
            .max(1)
            .min(physical_max),
        bytes.saturating_mul(64).clamp(1, WireLimits::MAX_FIELDS),
        bytes
            .saturating_mul(256)
            .saturating_add(additions.saturating_mul(bytes.max(1)))
            .clamp(1, WireLimits::MAX_REWRITE_WORK),
        64,
        bytes
            .max(source.state.source.components().len())
            .saturating_mul(4)
            .max(1),
        source
            .state
            .source
            .limits()
            .archive_limits()
            .max_metadata_items()
            .max(1),
        additions.max(1),
    )
}

fn charge_metadata_report(
    budget: &mut table_lock::WireBudget,
    report: litchi_iwa_protos::package_metadata_codec::RewriteReport,
) -> Result<(), BodyTableHiddenAxesError> {
    budget
        .charge_codec_report(
            report.fields(),
            report.work_bytes(),
            report.max_depth(),
            report.references_scanned(),
        )
        .and_then(|_| {
            let work = report
                .input_bytes()
                .checked_add(report.output_bytes())
                .and_then(|value| value.checked_add(report.components_scanned()))
                .and_then(|value| value.checked_add(report.additions()))
                .and_then(|value| value.checked_add(report.allocations()))
                .and_then(|value| value.checked_add(report.retained_bytes()))
                .and_then(|value| value.checked_add(report.scratch_bytes()))
                .ok_or(table_lock::BodyTableLockError::LimitExceeded {
                    kind: table_lock::BodyTableLockLimitKind::WireWork,
                    observed: u64::MAX,
                    maximum: u64::try_from(budget.wire_limits().max_rewrite_work())
                        .unwrap_or(u64::MAX),
                })?;
            budget.charge_payload_work(work)
        })
        .map_err(map_lock_error)
}

fn charge_metadata_requirements(
    budget: &mut table_lock::WireBudget,
    requirements: litchi_iwa_protos::package_metadata_codec::RewriteExecutionRequirements,
) -> Result<(), BodyTableHiddenAxesError> {
    budget
        .charge_output_bytes(requirements.output_bytes())
        .and_then(|_| {
            budget.charge_codec_report(
                requirements.fields(),
                requirements.work_bytes(),
                0,
                requirements.references(),
            )
        })
        .and_then(|_| {
            let work = requirements
                .components()
                .checked_add(requirements.allocations())
                .and_then(|value| value.checked_add(requirements.retained_bytes()))
                .and_then(|value| value.checked_add(requirements.scratch_bytes()))
                .ok_or(table_lock::BodyTableLockError::LimitExceeded {
                    kind: table_lock::BodyTableLockLimitKind::WireWork,
                    observed: u64::MAX,
                    maximum: u64::try_from(budget.wire_limits().max_rewrite_work())
                        .unwrap_or(u64::MAX),
                })?;
            budget.charge_payload_work(work)
        })
        .map_err(map_lock_error)
}

fn owner_creation_uuid(identifier: u64, role: u64) -> UuidBits {
    UuidBits::new(
        identifier ^ 0x9e37_79b9_7f4a_7c15 ^ role,
        identifier.rotate_left(29) ^ 0xd1b5_4a32_d192_ed03 ^ role.rotate_left(17),
    )
}

fn prepare_metadata_probe<'source>(
    source: &'source Package,
    graph: &Graph,
    physical_maximum: u64,
    budget: &mut table_lock::WireBudget,
) -> Result<Option<MetadataProbe<'source>>, BodyTableHiddenAxesError> {
    let Some(route) = metadata_route(source, budget)? else {
        return Ok(None);
    };
    let component = source
        .state
        .source
        .components()
        .get_index(route.component_index)
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    let payload = component
        .archive()
        .objects
        .get(route.object_index)
        .and_then(|object| object.messages.get(route.message_index))
        .filter(|message| message.type_ == METADATA_MESSAGE_TYPE)
        .map(|message| message.data.as_slice())
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    let target_component = source
        .state
        .source
        .components()
        .get_index(graph.target.model_component_index)
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    let target_locator = normalized_locator(target_component.name());
    let target_object_count = target_component.archive().objects.len();
    let metadata_item_capacity = payload
        .len()
        .min(
            source
                .state
                .source
                .limits()
                .archive_limits()
                .max_metadata_items(),
        )
        .max(1);
    let component_capacity = source
        .state
        .source
        .components()
        .len()
        .checked_add(1)
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    let target_identity_capacity = target_object_count.max(metadata_item_capacity);
    let component_identity_capacity = component_capacity.max(metadata_item_capacity);
    let metadata_identity_capacity =
        metadata_item_capacity
            .checked_mul(2)
            .ok_or(BodyTableHiddenAxesError::LimitExceeded {
                kind: BodyTableHiddenAxesLimitKind::PayloadItems,
                observed: u64::MAX,
                maximum: u64::try_from(budget.maximum_payload_references()).unwrap_or(u64::MAX),
            })?;
    budget
        .charge_payload_items(target_identity_capacity)
        .and_then(|_| budget.charge_payload_items(component_identity_capacity))
        .and_then(|_| budget.charge_payload_items(metadata_identity_capacity))
        .and_then(|_| budget.charge_payload_work(target_object_count))
        .and_then(|_| budget.charge_payload_work(component_capacity))
        .and_then(|_| budget.charge_payload_work(target_identity_capacity))
        .and_then(|_| budget.charge_payload_work(component_identity_capacity))
        .and_then(|_| budget.charge_payload_work(metadata_identity_capacity))
        .map_err(map_lock_error)?;
    let mut physical_object_identifiers = Vec::new();
    physical_object_identifiers
        .try_reserve_exact(target_object_count)
        .map_err(|_| BodyTableHiddenAxesError::Allocation {
            amount: target_object_count,
        })?;
    for object in &target_component.archive().objects {
        physical_object_identifiers.push(
            object
                .archive_info
                .identifier
                .ok_or(BodyTableHiddenAxesError::InvalidSource)?,
        );
    }
    physical_object_identifiers.sort_unstable();
    if physical_object_identifiers
        .windows(2)
        .any(|pair| pair[0] == pair[1])
    {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    let mut census = MetadataCensus::new(
        target_locator,
        target_object_count,
        metadata_item_capacity,
        component_identity_capacity,
    )?;
    let inspection = inspect_package_metadata_with_visitor(
        payload,
        metadata_options(source, payload.len(), OWNER_CREATION_OBJECTS),
        &mut census,
    )
    .map_err(map_metadata_error)?;
    charge_metadata_report(budget, inspection.report())?;
    if census.unknown
        || census.target_component_matches != 1
        || census
            .target_component_identifier
            .is_none_or(|identifier| identifier == 0)
        || census.target_object_identifiers.is_empty()
    {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    census.component_identifiers.sort_unstable();
    if census
        .component_identifiers
        .windows(2)
        .any(|pair| pair[0] == 0 || pair[0] == pair[1])
    {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    census.object_bindings.sort_unstable();
    if census
        .object_bindings
        .windows(2)
        .any(|pair| pair[0] == pair[1] || pair[0].1 == 0)
    {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    // Component callbacks are currently emitted before object callbacks by
    // the codec, but derive the selected component's physical registry from
    // the binding census after the complete pass.  This keeps the authority
    // proof correct if a future lazy projection changes callback ordering.
    census.target_object_identifiers.clear();
    let target_component_identifier = census
        .target_component_identifier
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    for &(component_identifier, object_identifier) in &census.object_bindings {
        if component_identifier == target_component_identifier {
            if census.target_object_identifiers.len() == census.target_object_identifiers.capacity()
            {
                return Err(BodyTableHiddenAxesError::LimitExceeded {
                    kind: BodyTableHiddenAxesLimitKind::PayloadItems,
                    observed: u64::try_from(
                        census.target_object_identifiers.len().saturating_add(1),
                    )
                    .unwrap_or(u64::MAX),
                    maximum: u64::try_from(census.target_object_identifiers.capacity())
                        .unwrap_or(u64::MAX),
                });
            }
            census.target_object_identifiers.push(object_identifier);
        }
    }
    census
        .uuids
        .sort_unstable_by_key(|uuid| (uuid.lower(), uuid.upper()));
    if census.uuids.windows(2).any(|pair| pair[0] == pair[1]) {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    census.target_object_identifiers.sort_unstable();
    if census
        .target_object_identifiers
        .windows(2)
        .any(|pair| pair[0] == pair[1])
        || census.target_object_identifiers.len() != physical_object_identifiers.len()
        || census.target_object_identifiers != physical_object_identifiers
    {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    let maximum_identifier = physical_maximum
        .max(inspection.last_object_identifier())
        .max(census.maximum);
    let selector = ComponentSelector::new(
        census
            .target_component_identifier
            .ok_or(BodyTableHiddenAxesError::InvalidSource)?,
        target_locator,
    );
    Ok(Some(MetadataProbe {
        route,
        payload,
        selector,
        expected_last_identifier: inspection.last_object_identifier(),
        maximum_identifier,
    }))
}

fn finish_metadata_plan<'source>(
    source: &'source Package,
    probe: MetadataProbe<'source>,
    ids: OwnerCreationIds,
    budget: &mut table_lock::WireBudget,
) -> Result<MetadataPlan<'source>, BodyTableHiddenAxesError> {
    let id_array = ids.as_array();
    let uuids = [
        owner_creation_uuid(id_array[0], 1),
        owner_creation_uuid(id_array[1], 2),
        owner_creation_uuid(id_array[2], 3),
        owner_creation_uuid(id_array[3], 4),
    ];
    if uuids
        .iter()
        .any(|uuid| uuid.lower() == 0 && uuid.upper() == 0)
        || uuids.windows(2).any(|pair| pair[0] == pair[1])
    {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    let mut collision = MetadataUuidCollisionVisitor::new(uuids);
    let inspection = inspect_package_metadata_with_visitor(
        probe.payload,
        metadata_options(source, probe.payload.len(), 0),
        &mut collision,
    )
    .map_err(map_metadata_error)?;
    charge_metadata_report(budget, inspection.report())?;
    if collision.unknown || collision.found {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    let new_last_identifier = probe
        .maximum_identifier
        .checked_add(OWNER_CREATION_OBJECTS as u64)
        .filter(|identifier| *identifier != 0)
        .ok_or(BodyTableHiddenAxesError::LimitExceeded {
            kind: BodyTableHiddenAxesLimitKind::PayloadObjects,
            observed: u64::MAX,
            maximum: u64::MAX - 1,
        })?;
    if new_last_identifier <= probe.expected_last_identifier {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    Ok(MetadataPlan {
        route: probe.route,
        payload: probe.payload,
        selector: probe.selector,
        expected_last_identifier: probe.expected_last_identifier,
        new_last_identifier,
        uuids,
    })
}

fn next_owner_creation_ids(floor: u64) -> Result<OwnerCreationIds, BodyTableHiddenAxesError> {
    let first = floor
        .checked_add(1)
        .filter(|identifier| *identifier != 0)
        .ok_or(BodyTableHiddenAxesError::LimitExceeded {
            kind: BodyTableHiddenAxesLimitKind::PayloadObjects,
            observed: u64::MAX,
            maximum: u64::MAX - 1,
        })?;
    let second = first
        .checked_add(1)
        .ok_or(BodyTableHiddenAxesError::LimitExceeded {
            kind: BodyTableHiddenAxesLimitKind::PayloadObjects,
            observed: u64::MAX,
            maximum: u64::MAX - 1,
        })?;
    let third = second
        .checked_add(1)
        .ok_or(BodyTableHiddenAxesError::LimitExceeded {
            kind: BodyTableHiddenAxesLimitKind::PayloadObjects,
            observed: u64::MAX,
            maximum: u64::MAX - 1,
        })?;
    let fourth = third
        .checked_add(1)
        .ok_or(BodyTableHiddenAxesError::LimitExceeded {
            kind: BodyTableHiddenAxesLimitKind::PayloadObjects,
            observed: u64::MAX,
            maximum: u64::MAX - 1,
        })?;
    Ok(OwnerCreationIds {
        column_formula: first,
        row_formula: second,
        column_filter: third,
        row_filter: fourth,
    })
}

fn prepare_owner_creation<'source>(
    source: &'source Package,
    graph: &Graph,
    axes: &HiddenAxes,
    budget: &mut table_lock::WireBudget,
) -> Result<OwnerCreationPlan<'source>, BodyTableHiddenAxesError> {
    let objects = global_objects(source, budget)?;
    let physical_maximum = maximum_physical_identifier(source, &objects, budget)?;
    let metadata_probe = prepare_metadata_probe(source, graph, physical_maximum, budget)?;
    let floor = metadata_probe
        .as_ref()
        .map_or(physical_maximum, |probe| probe.maximum_identifier);
    let ids = next_owner_creation_ids(floor)?;
    let object_ids = ids.as_array();
    if object_ids.windows(2).any(|pair| pair[0] == pair[1]) {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    budget
        .charge_payload_objects(OWNER_CREATION_OBJECTS)
        .and_then(|_| budget.charge_payload_messages(OWNER_CREATION_MESSAGES))
        .and_then(|_| budget.charge_payload_items(OWNER_CREATION_FIELD_INFOS))
        .map_err(map_lock_error)?;
    let mut object_ids_owned = Vec::new();
    object_ids_owned
        .try_reserve_exact(object_ids.len())
        .map_err(|_| BodyTableHiddenAxesError::Allocation {
            amount: object_ids.len(),
        })?;
    object_ids_owned.extend_from_slice(&object_ids);
    let object_ids = Arc::from(object_ids_owned);

    let (owner, active_uuid) = creation_states(source, graph, axes, ids, budget)?;
    validate_creation_uuid_namespace(source, &objects, graph, active_uuid, budget)?;
    let formula_columns_ref = codec::ReferenceSnapshot::new(
        NonZeroU64::new(ids.column_formula).ok_or(BodyTableHiddenAxesError::InvalidSource)?,
    );
    let formula_rows_ref = codec::ReferenceSnapshot::new(
        NonZeroU64::new(ids.row_formula).ok_or(BodyTableHiddenAxesError::InvalidSource)?,
    );
    let metadata = metadata_probe
        .map(|probe| finish_metadata_plan(source, probe, ids, budget))
        .transpose()?;
    Ok(OwnerCreationPlan {
        ids,
        object_ids,
        owner,
        active_uuid,
        formula_columns_ref,
        formula_rows_ref,
        metadata,
    })
}

fn creation_states(
    package: &Package,
    graph: &Graph,
    axes: &HiddenAxes,
    ids: OwnerCreationIds,
    budget: &mut table_lock::WireBudget,
) -> Result<(codec::HiddenStatesOwnerSnapshot, codec::UuidSnapshot), BodyTableHiddenAxesError> {
    let formula_uid = formula_owner_for_creation(package, graph, budget)?;
    let active_uuid = codec::UuidSnapshot::new(
        formula_uid
            .lower()
            .checked_add(4)
            .ok_or(BodyTableHiddenAxesError::InvalidSource)?,
        formula_uid.upper(),
    );
    validate_uuid(active_uuid)?;
    let column_extent_uid = codec::UuidSnapshot::new(
        active_uuid
            .lower()
            .checked_add(7)
            .ok_or(BodyTableHiddenAxesError::InvalidSource)?,
        active_uuid.upper(),
    );
    validate_uuid(column_extent_uid)?;
    let row_count = axes
        .iter()
        .filter(|axis| matches!(axis, AxisIndex::Row(_)))
        .count();
    let column_count = axes
        .iter()
        .filter(|axis| matches!(axis, AxisIndex::Column(_)))
        .count();
    let state_count =
        row_count
            .checked_add(column_count)
            .ok_or(BodyTableHiddenAxesError::LimitExceeded {
                kind: BodyTableHiddenAxesLimitKind::PayloadItems,
                observed: u64::MAX,
                maximum: u64::MAX,
            })?;
    // Admit all state and owner collection storage before constructing any
    // vectors. The helper objects themselves are charged separately below.
    budget
        .charge_payload_items(
            state_count
                .checked_add(OWNER_CREATION_OBJECTS)
                .and_then(|value| value.checked_add(OWNER_CREATION_FIELD_INFOS))
                .ok_or(BodyTableHiddenAxesError::LimitExceeded {
                    kind: BodyTableHiddenAxesLimitKind::PayloadItems,
                    observed: u64::MAX,
                    maximum: u64::MAX,
                })?,
        )
        .map_err(map_lock_error)?;
    budget
        .charge_payload_work(state_count.saturating_mul(4).saturating_add(8))
        .map_err(map_lock_error)?;
    let mut rows = Vec::new();
    rows.try_reserve_exact(row_count)
        .map_err(|_| BodyTableHiddenAxesError::Allocation { amount: row_count })?;
    let mut columns = Vec::new();
    columns
        .try_reserve_exact(column_count)
        .map_err(|_| BodyTableHiddenAxesError::Allocation {
            amount: column_count,
        })?;
    for axis in axes.iter() {
        match axis {
            AxisIndex::Row(index) => {
                let uid = graph
                    .rows
                    .get(index)
                    .copied()
                    .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
                rows.push(codec::RowOrColumnStateSnapshot::new(uid).with_user_hidden(Some(true)));
            },
            AxisIndex::Column(index) => {
                let uid = graph
                    .columns
                    .get(index)
                    .copied()
                    .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
                columns
                    .push(codec::RowOrColumnStateSnapshot::new(uid).with_user_hidden(Some(true)));
            },
        }
    }
    let column_filter = codec::ReferenceSnapshot::new(
        NonZeroU64::new(ids.column_filter).ok_or(BodyTableHiddenAxesError::InvalidSource)?,
    );
    let row_filter = codec::ReferenceSnapshot::new(
        NonZeroU64::new(ids.row_filter).ok_or(BodyTableHiddenAxesError::InvalidSource)?,
    );
    let column = codec::HiddenStateExtentSnapshot::new(
        column_extent_uid,
        codec::AxisDirection::Column,
        columns,
    )
    .map_err(|_| BodyTableHiddenAxesError::InvalidSource)?
    .with_needs_to_update_filter_set_for_import(Some(false))
    .with_filter_set(Some(column_filter));
    let row = codec::HiddenStateExtentSnapshot::new(active_uuid, codec::AxisDirection::Row, rows)
        .map_err(|_| BodyTableHiddenAxesError::InvalidSource)?
        .with_needs_to_update_filter_set_for_import(Some(false))
        .with_filter_set(Some(row_filter));
    let state = codec::HiddenStatesSnapshot::new(active_uuid, column, row);
    let owner = codec::HiddenStatesOwnerSnapshot::new(active_uuid, [state])
        .map_err(|_| BodyTableHiddenAxesError::InvalidSource)?;
    Ok((owner, active_uuid))
}

/// Prove that the UUIDs derived for a new hidden-state view are outside every
/// logical UUID namespace already present in the package.  Physical object
/// identifiers and metadata UUIDs are deliberately kept out of this census:
/// they have independent allocation authorities and are checked by their
/// respective transaction plans.
fn validate_creation_uuid_namespace(
    package: &Package,
    objects: &[ObjectLocation],
    graph: &Graph,
    active_uuid: codec::UuidSnapshot,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableHiddenAxesError> {
    let column_extent_uuid = codec::UuidSnapshot::new(
        active_uuid
            .lower()
            .checked_add(7)
            .ok_or(BodyTableHiddenAxesError::InvalidSource)?,
        active_uuid.upper(),
    );
    validate_uuid(column_extent_uuid)?;
    if active_uuid == column_extent_uuid {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    let expected_formula_owner_uid = if graph.profile.is_aggregate_only() {
        Some(codec::UuidSnapshot::new(
            active_uuid
                .lower()
                .checked_sub(4)
                .ok_or(BodyTableHiddenAxesError::InvalidSource)?,
            active_uuid.upper(),
        ))
    } else {
        None
    };

    let mut logical_uuids = Vec::new();
    let initial_capacity = graph
        .rows
        .len()
        .checked_add(graph.columns.len())
        .and_then(|value| value.checked_add(objects.len()))
        .ok_or(BodyTableHiddenAxesError::LimitExceeded {
            kind: BodyTableHiddenAxesLimitKind::PayloadItems,
            observed: u64::MAX,
            maximum: u64::try_from(budget.maximum_payload_references()).unwrap_or(u64::MAX),
        })?;
    budget
        .charge_payload_items(initial_capacity)
        .and_then(|_| budget.charge_payload_work(initial_capacity))
        .map_err(map_lock_error)?;
    logical_uuids
        .try_reserve_exact(initial_capacity)
        .map_err(|_| BodyTableHiddenAxesError::Allocation {
            amount: initial_capacity,
        })?;
    for uuid in graph.rows.iter().chain(graph.columns.iter()).copied() {
        push_logical_uuid(&mut logical_uuids, uuid, budget)?;
    }

    let mut map_dimensions = Vec::new();
    budget
        .charge_payload_items(objects.len())
        .and_then(|_| budget.charge_payload_work(objects.len()))
        .map_err(map_lock_error)?;
    map_dimensions
        .try_reserve_exact(objects.len())
        .map_err(|_| BodyTableHiddenAxesError::Allocation {
            amount: objects.len(),
        })?;
    if let Some(map_identifier) = graph.model.map {
        map_dimensions.push(UidMapDimensions {
            identifier: map_identifier.get(),
            columns: graph.model.columns,
            rows: graph.model.rows,
        });
    }

    for location in objects {
        let object = package
            .state
            .source
            .components()
            .get_index(location.component_index)
            .and_then(|component| component.archive().objects.get(location.object_index))
            .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
        for (message_index, message) in object.messages.iter().enumerate() {
            if message.type_ == HIDDEN_STATE_FORMULA_OWNER_MESSAGE_TYPE {
                budget
                    .charge_payload_work(message.data.len())
                    .map_err(map_lock_error)?;
                let info = validate_message_metadata(
                    object,
                    message_index,
                    HIDDEN_STATE_FORMULA_OWNER_MESSAGE_TYPE,
                )?;
                if !info.object_references.is_empty()
                    || !info.data_references.is_empty()
                    || !info.field_infos.is_empty()
                {
                    return Err(BodyTableHiddenAxesError::UnsupportedDependency);
                }
                let (owner, report) = codec::decode_hidden_state_formula_owner_with_report(
                    message.data.as_slice(),
                    codec_options(budget, message.data.len(), message.data.len())?,
                )
                .map_err(map_codec_error)?;
                charge_codec(budget, report)?;
                if let Some(owner_id) = owner.owner_id() {
                    let uuid = owner_id
                        .words_uuid()
                        .ok_or(BodyTableHiddenAxesError::UnsupportedDependency)?;
                    push_logical_uuid(&mut logical_uuids, uuid, budget)?;
                }
            } else if message.type_ == FORMULA_OWNER_MESSAGE_TYPE {
                budget
                    .charge_payload_work(message.data.len())
                    .map_err(map_lock_error)?;
                let _info =
                    validate_message_metadata(object, message_index, FORMULA_OWNER_MESSAGE_TYPE)?;
                let (owner, report) = codec::decode_formula_owner_dependencies_with_report(
                    message.data.as_slice(),
                    codec_options(budget, message.data.len(), message.data.len())?,
                )
                .map_err(map_codec_error)?;
                charge_codec(budget, report)?;
                if owner.has_dependencies() {
                    let Some(expected_formula_owner_uid) = expected_formula_owner_uid else {
                        return Err(BodyTableHiddenAxesError::UnsupportedDependency);
                    };
                    let Some(reference) = owner.formula_owner() else {
                        return Err(BodyTableHiddenAxesError::UnsupportedDependency);
                    };
                    validate_reference_shape(reference)?;
                    if owner.internal_formula_owner_id() == 0
                        || owner.owner_kind() != Some(1)
                        || owner.base_owner_uid().is_some()
                        || owner.formula_owner_uid() != expected_formula_owner_uid
                        || reference.identifier().get() != graph.target.drawable_identifier.get()
                    {
                        return Err(BodyTableHiddenAxesError::UnsupportedDependency);
                    }
                    let drawable = object_location(objects, graph.target.drawable_identifier)?;
                    if location.component_index == drawable.component_index
                        || location.identifier == graph.target.model_identifier.get()
                    {
                        return Err(BodyTableHiddenAxesError::UnsupportedDependency);
                    }
                    validate_aggregate_formula_owner_component(
                        package,
                        location.component_index,
                        budget,
                    )?;
                    validate_empty_formula_owner_dependencies(message.data.as_slice(), budget)?;
                    validate_aggregate_formula_owner_metadata(
                        object,
                        message_index,
                        drawable.identifier,
                        budget,
                    )?;
                    validate_aggregate_formula_owner_provenance(
                        package,
                        location.identifier,
                        &owner,
                        budget,
                    )?;
                    push_logical_uuid(&mut logical_uuids, owner.formula_owner_uid(), budget)?;
                    continue;
                }
                push_logical_uuid(&mut logical_uuids, owner.formula_owner_uid(), budget)?;
                if let Some(base_owner_uid) = owner.base_owner_uid() {
                    push_logical_uuid(&mut logical_uuids, base_owner_uid, budget)?;
                }
            } else if message.type_ == UID_MAP_MESSAGE_TYPE
                || message.type_ == LEGACY_UID_MAP_MESSAGE_TYPE
            {
                // The map is decoded in a second pass after every qualified
                // model has contributed its dimensions.  A UID-map payload
                // is not self-describing, so decoding it without a rooted
                // model shape would silently accept truncated permutations.
            } else if message.type_ == 6_001 {
                budget
                    .charge_payload_work(message.data.len())
                    .map_err(map_lock_error)?;
                let (model, report) = codec::decode_table_model_with_report(
                    message.data.as_slice(),
                    codec_options(budget, message.data.len(), message.data.len())?,
                )
                .map_err(map_codec_error)?;
                charge_codec(budget, report)?;
                if let Some(reference) = model.base_column_row_uids() {
                    push_uid_map_dimensions(
                        &mut map_dimensions,
                        reference.identifier(),
                        model.number_of_columns(),
                        model.number_of_rows(),
                        budget,
                    )?;
                }
                if let Some(owner) = model.hidden_states_owner() {
                    push_hidden_owner_uuids(&mut logical_uuids, owner, budget)?;
                }
            } else if message.type_ == 6_003 {
                budget
                    .charge_payload_work(message.data.len())
                    .map_err(map_lock_error)?;
                let info_result = codec::decode_table_info_with_report(
                    message.data.as_slice(),
                    codec_options(budget, message.data.len(), message.data.len())?,
                );
                match info_result {
                    Ok((info, report)) => {
                        charge_codec(budget, report)?;
                        if let Some(uuid) = info.hidden_states_uuid() {
                            push_logical_uuid(&mut logical_uuids, uuid, budget)?;
                        }
                    },
                    Err(error)
                        if error.resource_limit().is_none()
                            && error.allocation_amount().is_none() =>
                    {
                        // Message type 6003 is also the current
                        // TST.TableStyleArchive role. Qualify that known
                        // collision with the strict appearance codec before
                        // ignoring it in the UUID census; an unknown or
                        // malformed payload remains a hard source error.
                        let (_, report) = style_codec::decode_table_style_with_report(
                            message.data.as_slice(),
                            style_codec_options(budget, message.data.len())?,
                        )
                        .map_err(map_style_codec_error)?;
                        charge_style_codec(budget, report)?;
                    },
                    Err(error) => return Err(map_codec_error(error)),
                }
            } else if message.type_ == 6_000 {
                // Type 6000 is shared by the indexed table-info/model pair.
                // The selected pair was already qualified by resolve_graph;
                // other occurrences must produce one complete projection.
                let selected = location.component_index == graph.target.model_component_index
                    && location.object_index == graph.target.model_object_index
                    && message_index == graph.target.model_message_index
                    || location.component_index == graph.target.model_component_index
                        && location.object_index == graph.target.object_index
                        && message_index == graph.target.info_message_index;
                if !selected {
                    let attempt_work = message.data.len().checked_mul(2).ok_or(
                        BodyTableHiddenAxesError::LimitExceeded {
                            kind: BodyTableHiddenAxesLimitKind::TransactionWork,
                            observed: u64::MAX,
                            maximum: u64::try_from(budget.wire_limits().max_rewrite_work())
                                .unwrap_or(u64::MAX),
                        },
                    )?;
                    budget
                        .charge_payload_work(attempt_work)
                        .map_err(map_lock_error)?;
                    match codec::decode_table_model_with_report(
                        message.data.as_slice(),
                        codec_options(budget, message.data.len(), message.data.len())?,
                    ) {
                        Ok((model, model_report)) => {
                            // The strict projections are wire-disjoint: the
                            // model's required field 6 is a varint row count,
                            // while TableInfoArchive's field 6 is a
                            // length-delimited UID view.  A successful model
                            // decode therefore establishes the role without a
                            // second full bounded scan.
                            charge_codec(budget, model_report)?;
                            if let Some(reference) = model.base_column_row_uids() {
                                push_uid_map_dimensions(
                                    &mut map_dimensions,
                                    reference.identifier(),
                                    model.number_of_columns(),
                                    model.number_of_rows(),
                                    budget,
                                )?;
                            }
                            if let Some(owner) = model.hidden_states_owner() {
                                push_hidden_owner_uuids(&mut logical_uuids, owner, budget)?;
                            }
                        },
                        Err(error)
                            if error.resource_limit().is_some()
                                || error.allocation_amount().is_some() =>
                        {
                            return Err(map_codec_error(error));
                        },
                        Err(_) => {
                            let (info, report) = codec::decode_table_info_with_report(
                                message.data.as_slice(),
                                codec_options(budget, message.data.len(), message.data.len())?,
                            )
                            .map_err(map_codec_error)?;
                            charge_codec(budget, report)?;
                            if let Some(uuid) = info.hidden_states_uuid() {
                                push_logical_uuid(&mut logical_uuids, uuid, budget)?;
                            }
                        },
                    }
                }
            }
        }
    }
    map_dimensions.sort_unstable_by_key(|dimensions| dimensions.identifier);
    for pair in map_dimensions.windows(2) {
        if pair[0].identifier == pair[1].identifier
            && (pair[0].columns != pair[1].columns || pair[0].rows != pair[1].rows)
        {
            return Err(BodyTableHiddenAxesError::UnsupportedDependency);
        }
    }
    map_dimensions.dedup_by_key(|dimensions| dimensions.identifier);
    let mut map_dimension_cursor = 0usize;
    for location in objects {
        let object = package
            .state
            .source
            .components()
            .get_index(location.component_index)
            .and_then(|component| component.archive().objects.get(location.object_index))
            .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
        let map_message_count = object
            .messages
            .iter()
            .filter(|message| {
                message.type_ == UID_MAP_MESSAGE_TYPE
                    || message.type_ == LEGACY_UID_MAP_MESSAGE_TYPE
            })
            .count();
        if map_message_count == 0 {
            if map_dimensions
                .binary_search_by_key(&location.identifier, |dimensions| dimensions.identifier)
                .is_ok()
            {
                // A qualified model claimed this object as its UID map, but
                // the physical object no longer carries exactly one map
                // message.  Do not let a retagged or truncated map escape
                // the reverse census merely because it is not recognized in
                // this pass.
                return Err(BodyTableHiddenAxesError::InvalidSource);
            }
            continue;
        }
        if map_message_count != 1 {
            return Err(BodyTableHiddenAxesError::InvalidSource);
        }
        let map_identifier = object
            .archive_info
            .identifier
            .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
        let dimensions = map_dimensions
            .binary_search_by_key(&map_identifier, |dimensions| dimensions.identifier)
            .ok()
            .map(|index| map_dimensions[index])
            .ok_or(BodyTableHiddenAxesError::UnsupportedDependency)?;
        if map_dimensions
            .get(map_dimension_cursor)
            .is_none_or(|expected| expected.identifier != dimensions.identifier)
        {
            return Err(BodyTableHiddenAxesError::InvalidSource);
        }
        map_dimension_cursor = map_dimension_cursor
            .checked_add(1)
            .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
        for (message_index, message) in object.messages.iter().enumerate() {
            if message.type_ == UID_MAP_MESSAGE_TYPE || message.type_ == LEGACY_UID_MAP_MESSAGE_TYPE
            {
                append_uid_map_uuids(
                    package,
                    *location,
                    message_index,
                    dimensions,
                    &mut logical_uuids,
                    budget,
                )?;
            }
        }
    }
    if map_dimension_cursor != map_dimensions.len() {
        // A model-derived dimension had no physical object in the source.
        // The selected graph may have carried a stale map reference, and an
        // unselected projection may have been truncated; both are unsafe to
        // treat as an empty UUID namespace.
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    logical_uuids.sort_unstable_by_key(|uuid| (uuid.lower(), uuid.upper()));
    logical_uuids.dedup();
    if logical_uuids
        .binary_search_by_key(&(active_uuid.lower(), active_uuid.upper()), |uuid| {
            (uuid.lower(), uuid.upper())
        })
        .is_ok()
        || logical_uuids
            .binary_search_by_key(
                &(column_extent_uuid.lower(), column_extent_uuid.upper()),
                |uuid| (uuid.lower(), uuid.upper()),
            )
            .is_ok()
    {
        return Err(BodyTableHiddenAxesError::UnsupportedDependency);
    }
    let levels = if logical_uuids.len() <= 1 {
        0
    } else {
        (usize::BITS - (logical_uuids.len() - 1).leading_zeros()) as usize
    };
    budget
        .charge_payload_work(logical_uuids.len().checked_mul(levels).ok_or(
            BodyTableHiddenAxesError::LimitExceeded {
                kind: BodyTableHiddenAxesLimitKind::TransactionWork,
                observed: u64::MAX,
                maximum: u64::try_from(budget.wire_limits().max_rewrite_work()).unwrap_or(u64::MAX),
            },
        )?)
        .map_err(map_lock_error)?;
    Ok(())
}

fn push_logical_uuid(
    values: &mut Vec<codec::UuidSnapshot>,
    uuid: codec::UuidSnapshot,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableHiddenAxesError> {
    validate_uuid(uuid)?;
    if values.len() == values.capacity() {
        budget
            .charge_payload_items(1)
            .and_then(|_| budget.charge_payload_work(1))
            .map_err(map_lock_error)?;
        values
            .try_reserve_exact(1)
            .map_err(|_| BodyTableHiddenAxesError::Allocation { amount: 1 })?;
    }
    values.push(uuid);
    Ok(())
}

fn push_hidden_owner_uuids(
    values: &mut Vec<codec::UuidSnapshot>,
    owner: &codec::HiddenStatesOwnerSnapshot,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableHiddenAxesError> {
    push_logical_uuid(values, owner.owner_uid(), budget)?;
    for state in owner.hidden_states() {
        push_logical_uuid(values, state.hidden_states_uid(), budget)?;
        push_logical_uuid(
            values,
            state.column_hidden_state_extent().hidden_state_extent_uid(),
            budget,
        )?;
        push_logical_uuid(
            values,
            state.row_hidden_state_extent().hidden_state_extent_uid(),
            budget,
        )?;
        for record in state
            .column_hidden_state_extent()
            .base_hidden_states()
            .iter()
            .chain(state.row_hidden_state_extent().base_hidden_states().iter())
        {
            push_logical_uuid(values, record.row_or_column_uid(), budget)?;
        }
    }
    Ok(())
}

fn push_uid_map_dimensions(
    dimensions: &mut Vec<UidMapDimensions>,
    identifier: NonZeroU64,
    columns: u32,
    rows: u32,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableHiddenAxesError> {
    if dimensions.len() == dimensions.capacity() {
        budget
            .charge_payload_items(1)
            .and_then(|_| budget.charge_payload_work(1))
            .map_err(map_lock_error)?;
        dimensions
            .try_reserve_exact(1)
            .map_err(|_| BodyTableHiddenAxesError::Allocation { amount: 1 })?;
    }
    dimensions.push(UidMapDimensions {
        identifier: identifier.get(),
        columns,
        rows,
    });
    Ok(())
}

fn append_uid_map_uuids(
    package: &Package,
    location: ObjectLocation,
    message_index: usize,
    dimensions: UidMapDimensions,
    values: &mut Vec<codec::UuidSnapshot>,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableHiddenAxesError> {
    let object = package
        .state
        .source
        .components()
        .get_index(location.component_index)
        .and_then(|component| component.archive().objects.get(location.object_index))
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    if object.archive_info.identifier != Some(dimensions.identifier) {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    let message = object
        .messages
        .get(message_index)
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    let allow_legacy = message.type_ == LEGACY_UID_MAP_MESSAGE_TYPE
        && info.type_ == LEGACY_UID_MAP_MESSAGE_TYPE
        && info.versions.as_slice() == CURRENT_MESSAGE_VERSIONS;
    codec::validate_column_row_uid_map_message_type(message.type_, allow_legacy)
        .map_err(|_| BodyTableHiddenAxesError::InvalidSource)?;
    validate_no_reference_metadata(object, message_index, message.type_, budget)?;
    let columns =
        usize::try_from(dimensions.columns).map_err(|_| BodyTableHiddenAxesError::InvalidSource)?;
    let rows =
        usize::try_from(dimensions.rows).map_err(|_| BodyTableHiddenAxesError::InvalidSource)?;
    let element_count = columns
        .checked_mul(3)
        .and_then(|value| rows.checked_mul(3).and_then(|row| value.checked_add(row)))
        .ok_or(BodyTableHiddenAxesError::LimitExceeded {
            kind: BodyTableHiddenAxesLimitKind::PayloadItems,
            observed: u64::MAX,
            maximum: u64::try_from(budget.maximum_payload_references()).unwrap_or(u64::MAX),
        })?;
    budget
        .charge_payload_work(message.data.len())
        .and_then(|_| budget.charge_payload_items(element_count))
        .map_err(map_lock_error)?;
    let limits = budget.wire_limits();
    let options = uid_codec::DecodeOptions::new(
        message.data.len(),
        budget.remaining_wire_fields(),
        budget.remaining_wire_work(),
        u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX),
        0,
        element_count,
        limits.max_output_bytes(),
        limits.max_output_bytes(),
    );
    let (map, report) =
        uid_codec::decode_column_row_uid_map(message.data.as_slice(), columns, rows, options)
            .map_err(map_uid_codec_error)?;
    budget
        .charge_codec_report(report.fields(), report.work_bytes(), report.max_depth(), 0)
        .map_err(map_lock_error)?;
    budget
        .charge_payload_work(
            report
                .records()
                .checked_add(report.elements())
                .and_then(|value| value.checked_add(element_count))
                .ok_or(BodyTableHiddenAxesError::LimitExceeded {
                    kind: BodyTableHiddenAxesLimitKind::TransactionWork,
                    observed: u64::MAX,
                    maximum: u64::try_from(budget.wire_limits().max_rewrite_work())
                        .unwrap_or(u64::MAX),
                })?,
        )
        .map_err(map_lock_error)?;
    for uid in map
        .sorted_row_uids()
        .iter()
        .chain(map.sorted_column_uids().iter())
    {
        push_logical_uuid(
            values,
            codec::UuidSnapshot::new(uid.lower(), uid.upper()),
            budget,
        )?;
    }
    Ok(())
}

fn uuid_words(uuid: codec::UuidSnapshot) -> [Option<u32>; 4] {
    [
        Some(uuid.lower() as u32),
        Some((uuid.lower() >> 32) as u32),
        Some(uuid.upper() as u32),
        Some((uuid.upper() >> 32) as u32),
    ]
}

fn helper_formula_value(uuid: codec::UuidSnapshot) -> codec::HiddenStateFormulaOwnerSnapshot {
    codec::HiddenStateFormulaOwnerSnapshot::new(
        Some(codec::CfuuidSnapshot::new(None, uuid_words(uuid))),
        Some(false),
    )
}

fn helper_formula_measure(
    uuid: codec::UuidSnapshot,
    budget: &table_lock::WireBudget,
) -> Result<(codec::DecodeReport, codec::DecodeOptions), BodyTableHiddenAxesError> {
    let options = codec_options(budget, 0, budget.wire_limits().max_output_bytes())?;
    let report = codec::measure_hidden_state_formula_owner(&helper_formula_value(uuid), options)
        .map_err(map_codec_error)?;
    Ok((report, options))
}

fn helper_formula_payload(
    uuid: codec::UuidSnapshot,
    options: codec::DecodeOptions,
    expected: codec::DecodeReport,
) -> Result<Vec<u8>, BodyTableHiddenAxesError> {
    let encoded = codec::encode_hidden_state_formula_owner(&helper_formula_value(uuid), options)
        .map_err(map_codec_error)?;
    if encoded.report() != expected || encoded.bytes().len() != expected.output_bytes() {
        return Err(BodyTableHiddenAxesError::Verification);
    }
    Ok(encoded.into_bytes())
}

fn helper_filter_value() -> Result<codec::FilterSetSnapshot, BodyTableHiddenAxesError> {
    codec::FilterSetSnapshot::new(Some(0), Some(false), Some(false), []).map_err(map_codec_error)
}

fn helper_filter_measure(
    budget: &table_lock::WireBudget,
) -> Result<(codec::DecodeReport, codec::DecodeOptions), BodyTableHiddenAxesError> {
    let options = codec_options(budget, 0, budget.wire_limits().max_output_bytes())?;
    let report =
        codec::measure_filter_set(&helper_filter_value()?, options).map_err(map_codec_error)?;
    Ok((report, options))
}

fn helper_filter_payload(
    options: codec::DecodeOptions,
    expected: codec::DecodeReport,
) -> Result<Vec<u8>, BodyTableHiddenAxesError> {
    let encoded =
        codec::encode_filter_set(&helper_filter_value()?, options).map_err(map_codec_error)?;
    if encoded.report() != expected || encoded.bytes().len() != expected.output_bytes() {
        return Err(BodyTableHiddenAxesError::Verification);
    }
    Ok(encoded.into_bytes())
}

fn build_owner_helper_objects(
    plan: &OwnerCreationPlan<'_>,
    archive_limits: litchi_iwa_core::Limits,
    _wire_limits: WireLimits,
    budget: &mut table_lock::WireBudget,
) -> Result<Vec<ArchiveObject>, BodyTableHiddenAxesError> {
    let row_uid = plan.active_uuid;
    let column_uid = codec::UuidSnapshot::new(
        row_uid
            .lower()
            .checked_add(7)
            .ok_or(BodyTableHiddenAxesError::InvalidSource)?,
        row_uid.upper(),
    );
    validate_uuid(column_uid)?;
    let (column_formula_report, column_formula_options) =
        helper_formula_measure(column_uid, budget)?;
    let (row_formula_report, row_formula_options) = helper_formula_measure(row_uid, budget)?;
    let (filter_report, filter_options) = helper_filter_measure(budget)?;
    let payload_bytes = column_formula_report
        .output_bytes()
        .checked_add(row_formula_report.output_bytes())
        .and_then(|length| length.checked_add(filter_report.output_bytes().checked_mul(2)?))
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    budget
        .charge_output_bytes(payload_bytes)
        .and_then(|_| budget.charge_payload_work(payload_bytes))
        .map_err(map_lock_error)?;
    charge_codec(budget, column_formula_report)?;
    charge_codec(budget, row_formula_report)?;
    charge_codec(budget, filter_report)?;
    let formula_payloads = [
        helper_formula_payload(column_uid, column_formula_options, column_formula_report)?,
        helper_formula_payload(row_uid, row_formula_options, row_formula_report)?,
    ];
    let filter_payload = helper_filter_payload(filter_options, filter_report)?;
    if formula_payloads[0].len() != column_formula_report.output_bytes()
        || formula_payloads[1].len() != row_formula_report.output_bytes()
        || filter_payload.len() != filter_report.output_bytes()
    {
        return Err(BodyTableHiddenAxesError::Verification);
    }
    let ids = plan.ids.as_array();
    let payloads = [
        formula_payloads[0].clone(),
        formula_payloads[1].clone(),
        filter_payload.clone(),
        filter_payload,
    ];
    let types = [
        HIDDEN_STATE_FORMULA_OWNER_MESSAGE_TYPE,
        HIDDEN_STATE_FORMULA_OWNER_MESSAGE_TYPE,
        FILTER_SET_MESSAGE_TYPE,
        FILTER_SET_MESSAGE_TYPE,
    ];
    let mut objects = Vec::new();
    objects
        .try_reserve_exact(OWNER_CREATION_OBJECTS)
        .map_err(|_| BodyTableHiddenAxesError::Allocation {
            amount: OWNER_CREATION_OBJECTS,
        })?;
    for ((identifier, type_), data) in ids.into_iter().zip(types).zip(payloads) {
        let mut messages = Vec::new();
        messages
            .try_reserve_exact(1)
            .map_err(|_| BodyTableHiddenAxesError::Allocation { amount: 1 })?;
        messages.push(RawMessage { type_, data });
        objects.push(
            ArchiveObject::new_with_limits(identifier, messages, archive_limits)
                .map_err(map_core_error)?,
        );
    }
    let helper_archive = Archive { objects };
    let helper_bytes = helper_archive
        .encoded_len_with_limits(archive_limits)
        .map_err(map_core_error)?;
    budget
        .charge_payload_bytes(helper_bytes)
        .and_then(|_| budget.charge_total_payload_bytes(helper_bytes))
        .and_then(|_| budget.charge_payload_work(helper_bytes))
        .map_err(map_lock_error)?;
    Ok(helper_archive.objects)
}

fn rewrite_metadata_entry_for_creation(
    source: &Package,
    plan: &MetadataPlan<'_>,
    payload: Vec<u8>,
    archive_limits: litchi_iwa_core::Limits,
    budget: &mut table_lock::WireBudget,
) -> Result<Vec<u8>, BodyTableHiddenAxesError> {
    let component = source
        .state
        .source
        .components()
        .get_index(plan.route.component_index)
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    let object = component
        .archive()
        .objects
        .get(plan.route.object_index)
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    let old_message = object
        .messages
        .get(plan.route.message_index)
        .filter(|message| message.type_ == METADATA_MESSAGE_TYPE)
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    let entry = source
        .state
        .source
        .package()
        .iter()
        .find(|entry| entry.name() == METADATA_ENTRY_NAME)
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    let (archive_bound, compressed_bound, replacement_entry_bound) = metadata_replacement_bounds(
        source,
        component,
        old_message.data.len(),
        payload.len(),
        entry,
    )?;
    budget
        .charge_output_bytes(archive_bound)
        .and_then(|_| budget.charge_output_bytes(compressed_bound))
        .and_then(|_| budget.charge_output_bytes(replacement_entry_bound))
        .and_then(|_| budget.charge_payload_bytes(archive_bound))
        .and_then(|_| budget.charge_total_payload_bytes(archive_bound))
        .and_then(|_| budget.charge_payload_work(archive_bound))
        .and_then(|_| budget.charge_payload_work(compressed_bound))
        .map_err(map_lock_error)?;
    let (mut archive, _) = page_layout::editable_archive(source, METADATA_ENTRY_NAME)
        .map_err(map_page_layout_error)?;
    let object = archive
        .objects
        .get_mut(plan.route.object_index)
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    object
        .replace_message_preserving_header_with_limits(
            plan.route.message_index,
            RawMessage {
                type_: METADATA_MESSAGE_TYPE,
                data: payload,
            },
            archive_limits,
        )
        .map_err(map_core_error)?;
    let compressed =
        page_layout::compress_archive(archive, archive_limits).map_err(map_page_layout_error)?;
    if compressed.len() > compressed_bound {
        return Err(BodyTableHiddenAxesError::Verification);
    }
    Ok(compressed)
}

fn metadata_replacement_bounds(
    source: &Package,
    component: &litchi_iwa_archive::Component,
    old_payload_len: usize,
    new_payload_len: usize,
    entry: &litchi_iwa_archive::package::Entry,
) -> Result<(usize, usize, usize), BodyTableHiddenAxesError> {
    let stream_length = archive_source_length(component.archive())?;
    let archive_bound = stream_length
        .checked_sub(old_payload_len)
        .and_then(|length| length.checked_add(new_payload_len))
        .and_then(|length| length.checked_add(64))
        .ok_or(BodyTableHiddenAxesError::LimitExceeded {
            kind: BodyTableHiddenAxesLimitKind::OutputBytes,
            observed: u64::MAX,
            maximum: source.state.source.limits().max_input_bytes(),
        })?;
    let compressed_bound = table_lock::snappy_compressed_bound(archive_bound).ok_or(
        BodyTableHiddenAxesError::LimitExceeded {
            kind: BodyTableHiddenAxesLimitKind::OutputBytes,
            observed: u64::MAX,
            maximum: source.state.source.limits().max_input_bytes(),
        },
    )?;
    let replacement_entry_bound = match entry.metadata().central().compression_method() {
        0 => compressed_bound,
        8 => table_lock::deflate_compressed_bound(compressed_bound).ok_or(
            BodyTableHiddenAxesError::LimitExceeded {
                kind: BodyTableHiddenAxesLimitKind::OutputBytes,
                observed: u64::MAX,
                maximum: source.state.source.limits().max_input_bytes(),
            },
        )?,
        _ => return Err(BodyTableHiddenAxesError::UnsupportedSource),
    };
    Ok((archive_bound, compressed_bound, replacement_entry_bound))
}

fn formula_owner_for_creation(
    package: &Package,
    graph: &Graph,
    budget: &mut table_lock::WireBudget,
) -> Result<codec::UuidSnapshot, BodyTableHiddenAxesError> {
    let objects = global_objects(package, budget)?;
    let drawable = object_location(&objects, graph.target.drawable_identifier)?;
    let model = object_location(&objects, graph.target.model_identifier)?;
    if drawable.component_index != model.component_index || drawable.identifier == model.identifier
    {
        return Err(BodyTableHiddenAxesError::UnsupportedDependency);
    }
    let messages = formula_owner_messages(package, &objects, budget)?;
    formula_owner_for(package, &messages, drawable, model, graph.profile, budget)?
        .ok_or(BodyTableHiddenAxesError::UnsupportedDependency)
}

fn desired_owner(
    graph: &Graph,
    axes: &HiddenAxes,
    budget: &mut table_lock::WireBudget,
) -> Result<(codec::HiddenStatesOwnerSnapshot, codec::UuidSnapshot), BodyTableHiddenAxesError> {
    let owner = graph
        .model
        .owner
        .as_ref()
        .ok_or(BodyTableHiddenAxesError::UnsupportedDependency)?;
    let active_uuid = graph
        .info
        .hidden_uuid
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    if owner.hidden_states().len() != 1
        || owner
            .hidden_states()
            .first()
            .is_none_or(|state| state.hidden_states_uid() != active_uuid)
    {
        // The strict owner codec can preserve opaque fields, but it cannot
        // prove byte-exact preservation of an inactive view while changing a
        // sibling.  Reject that shape before allocating a candidate.
        return Err(BodyTableHiddenAxesError::UnsupportedDependency);
    }
    let state = owner
        .hidden_states()
        .first()
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    let column = update_extent(
        state.column_hidden_state_extent(),
        &graph.column_indices,
        &graph.columns,
        axes,
        false,
        budget,
    )?;
    let row = update_extent(
        state.row_hidden_state_extent(),
        &graph.row_indices,
        &graph.rows,
        axes,
        true,
        budget,
    )?;
    let additional_states =
        axes.as_slice()
            .len()
            .checked_mul(2)
            .ok_or(BodyTableHiddenAxesError::LimitExceeded {
                kind: BodyTableHiddenAxesLimitKind::TransactionWork,
                observed: u64::MAX,
                maximum: u64::try_from(budget.wire_limits().max_rewrite_work()).unwrap_or(u64::MAX),
            })?;
    charge_owner_storage(owner, additional_states, budget)?;
    let updated_state = codec::HiddenStatesSnapshot::new(active_uuid, column, row);
    let rebuilt = codec::HiddenStatesOwnerSnapshot::new(owner.owner_uid(), [updated_state])
        .map_err(|_| BodyTableHiddenAxesError::InvalidSource)?;
    Ok((rebuilt, active_uuid))
}

fn update_extent(
    extent: &codec::HiddenStateExtentSnapshot,
    physical: &UidIndex,
    physical_order: &[codec::UuidSnapshot],
    axes: &HiddenAxes,
    row: bool,
    budget: &mut table_lock::WireBudget,
) -> Result<codec::HiddenStateExtentSnapshot, BodyTableHiddenAxesError> {
    let loop_work = extent
        .base_hidden_states()
        .len()
        .checked_add(physical_order.len())
        .ok_or(BodyTableHiddenAxesError::LimitExceeded {
            kind: BodyTableHiddenAxesLimitKind::TransactionWork,
            observed: u64::MAX,
            maximum: u64::try_from(budget.wire_limits().max_rewrite_work()).unwrap_or(u64::MAX),
        })?;
    budget
        .charge_payload_work(loop_work)
        .map_err(map_lock_error)?;
    let mut existing = Vec::new();
    existing
        .try_reserve_exact(extent.base_hidden_states().len())
        .map_err(|_| BodyTableHiddenAxesError::Allocation {
            amount: extent.base_hidden_states().len(),
        })?;
    for state in extent.base_hidden_states() {
        validate_uuid(state.row_or_column_uid())?;
        existing.push(state.row_or_column_uid());
    }
    existing.sort_unstable_by_key(|uid| (uid.lower(), uid.upper()));
    if existing.windows(2).any(|pair| pair[0] == pair[1]) {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    let levels = if existing.len() <= 1 {
        0
    } else {
        (usize::BITS - (existing.len() - 1).leading_zeros()) as usize
    };
    budget
        .charge_payload_work(existing.len().checked_mul(levels).ok_or(
            BodyTableHiddenAxesError::LimitExceeded {
                kind: BodyTableHiddenAxesLimitKind::TransactionWork,
                observed: u64::MAX,
                maximum: u64::try_from(budget.wire_limits().max_rewrite_work()).unwrap_or(u64::MAX),
            },
        )?)
        .map_err(map_lock_error)?;
    let mut states = Vec::new();
    states
        .try_reserve_exact(extent.base_hidden_states().len())
        .map_err(|_| BodyTableHiddenAxesError::Allocation {
            amount: extent.base_hidden_states().len(),
        })?;
    for state in extent.base_hidden_states() {
        let index = physical
            .index_of(state.row_or_column_uid())
            .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
        let hidden = axes.contains(if row {
            AxisIndex::row(index)
        } else {
            AxisIndex::column(index)
        });
        states.push(
            codec::RowOrColumnStateSnapshot::new(state.row_or_column_uid())
                .with_user_hidden(if hidden {
                    Some(true)
                } else {
                    state.user_hidden().filter(|value| !*value)
                })
                .with_filtered(state.filtered())
                .with_pivot_hidden(state.pivot_hidden()),
        );
    }
    for (index, uid) in physical_order.iter().copied().enumerate() {
        if axes.contains(if row {
            AxisIndex::row(index)
        } else {
            AxisIndex::column(index)
        }) && existing
            .binary_search_by_key(&(uid.lower(), uid.upper()), |candidate| {
                (candidate.lower(), candidate.upper())
            })
            .is_err()
        {
            states
                .try_reserve(1)
                .map_err(|_| BodyTableHiddenAxesError::Allocation { amount: 1 })?;
            states.push(codec::RowOrColumnStateSnapshot::new(uid).with_user_hidden(Some(true)));
        }
    }
    codec::HiddenStateExtentSnapshot::new(
        extent.hidden_state_extent_uid(),
        extent.direction(),
        states,
    )
    .map(|value| {
        value
            .with_needs_to_update_filter_set_for_import(
                extent.needs_to_update_filter_set_for_import(),
            )
            .with_filter_set(extent.filter_set())
    })
    .map_err(|_| BodyTableHiddenAxesError::InvalidSource)
}

fn archive_source_length(archive: &Archive) -> Result<usize, BodyTableHiddenAxesError> {
    archive.objects.iter().try_fold(0usize, |end, object| {
        let object_end = usize::try_from(object.header_offset)
            .ok()
            .and_then(|offset| offset.checked_add(usize::try_from(object.header_length).ok()?))
            .and_then(|offset| offset.checked_add(usize::try_from(object.data_length).ok()?))
            .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
        Ok(end.max(object_end))
    })
}

fn charge_owner_storage(
    owner: &codec::HiddenStatesOwnerSnapshot,
    additional: usize,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableHiddenAxesError> {
    let mut count = owner.hidden_states().len();
    for state in owner.hidden_states() {
        count = count
            .checked_add(2)
            .and_then(|value| {
                value.checked_add(
                    state
                        .column_hidden_state_extent()
                        .base_hidden_states()
                        .len(),
                )
            })
            .and_then(|value| {
                value.checked_add(state.row_hidden_state_extent().base_hidden_states().len())
            })
            .ok_or(BodyTableHiddenAxesError::LimitExceeded {
                kind: BodyTableHiddenAxesLimitKind::TransactionWork,
                observed: u64::MAX,
                maximum: u64::try_from(budget.wire_limits().max_rewrite_work()).unwrap_or(u64::MAX),
            })?;
    }
    count = count
        .checked_add(additional)
        .ok_or(BodyTableHiddenAxesError::LimitExceeded {
            kind: BodyTableHiddenAxesLimitKind::TransactionWork,
            observed: u64::MAX,
            maximum: u64::try_from(budget.wire_limits().max_rewrite_work()).unwrap_or(u64::MAX),
        })?;
    budget.charge_payload_work(count).map_err(map_lock_error)
}

fn charge_rewrite_requirements(
    budget: &mut table_lock::WireBudget,
    requirements: codec::RewriteExecutionRequirements,
) -> Result<(), BodyTableHiddenAxesError> {
    let retained_work = requirements
        .states()
        .checked_add(requirements.allocations())
        .and_then(|value| value.checked_add(requirements.retained_bytes()))
        .and_then(|value| value.checked_add(requirements.scratch_bytes()))
        .ok_or(BodyTableHiddenAxesError::LimitExceeded {
            kind: BodyTableHiddenAxesLimitKind::TransactionWork,
            observed: u64::MAX,
            maximum: u64::try_from(budget.wire_limits().max_rewrite_work()).unwrap_or(u64::MAX),
        })?;
    budget
        .charge_output_bytes(requirements.output_bytes())
        .and_then(|_| {
            budget.charge_codec_report(
                requirements.fields(),
                requirements.work_bytes(),
                requirements.max_depth(),
                0,
            )
        })
        .and_then(|_| budget.charge_payload_work(retained_work))
        .map_err(map_lock_error)
}

fn desired_model_snapshot(
    graph: &Graph,
    owner: &codec::HiddenStatesOwnerSnapshot,
    active_uuid: codec::UuidSnapshot,
    formula_columns_ref: Option<codec::ReferenceSnapshot>,
    formula_rows_ref: Option<codec::ReferenceSnapshot>,
    budget: &mut table_lock::WireBudget,
) -> Result<codec::TableModelSnapshot, BodyTableHiddenAxesError> {
    let active = owner
        .hidden_states()
        .first()
        .filter(|state| state.hidden_states_uid() == active_uuid)
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    charge_owner_storage(owner, 0, budget)?;
    let count_work = active
        .row_hidden_state_extent()
        .base_hidden_states()
        .len()
        .checked_add(
            active
                .column_hidden_state_extent()
                .base_hidden_states()
                .len(),
        )
        .and_then(|count| count.checked_mul(5))
        .ok_or(BodyTableHiddenAxesError::LimitExceeded {
            kind: BodyTableHiddenAxesLimitKind::TransactionWork,
            observed: u64::MAX,
            maximum: u64::try_from(budget.wire_limits().max_rewrite_work()).unwrap_or(u64::MAX),
        })?;
    budget
        .charge_payload_work(count_work)
        .map_err(map_lock_error)?;
    let total_rows = active
        .row_hidden_state_extent()
        .base_hidden_states()
        .iter()
        .filter(|state| {
            state.user_hidden() == Some(true)
                || state.filtered() == Some(true)
                || state.pivot_hidden() == Some(true)
        })
        .count();
    let total_columns = active
        .column_hidden_state_extent()
        .base_hidden_states()
        .iter()
        .filter(|state| {
            state.user_hidden() == Some(true)
                || state.filtered() == Some(true)
                || state.pivot_hidden() == Some(true)
        })
        .count();
    let user_rows = active
        .row_hidden_state_extent()
        .base_hidden_states()
        .iter()
        .filter(|state| state.user_hidden() == Some(true))
        .count();
    let user_columns = active
        .column_hidden_state_extent()
        .base_hidden_states()
        .iter()
        .filter(|state| state.user_hidden() == Some(true))
        .count();
    let filtered_rows = active
        .row_hidden_state_extent()
        .base_hidden_states()
        .iter()
        .filter(|state| state.filtered() == Some(true))
        .count();
    Ok(
        codec::TableModelSnapshot::new(graph.model.rows, graph.model.columns)
            .with_number_of_hidden_rows(
                graph
                    .model
                    .hidden_rows
                    .map(|_| u32::try_from(total_rows))
                    .transpose()
                    .map_err(|_| BodyTableHiddenAxesError::InvalidSource)?,
            )
            .with_number_of_hidden_columns(
                graph
                    .model
                    .hidden_columns
                    .map(|_| u32::try_from(total_columns))
                    .transpose()
                    .map_err(|_| BodyTableHiddenAxesError::InvalidSource)?,
            )
            .with_number_of_filtered_rows(
                graph
                    .model
                    .filtered_rows
                    .map(|_| u32::try_from(filtered_rows))
                    .transpose()
                    .map_err(|_| BodyTableHiddenAxesError::InvalidSource)?,
            )
            .with_number_of_user_hidden_rows(
                graph
                    .model
                    .user_rows
                    .map(|_| u32::try_from(user_rows))
                    .transpose()
                    .map_err(|_| BodyTableHiddenAxesError::InvalidSource)?,
            )
            .with_number_of_user_hidden_columns(
                graph
                    .model
                    .user_columns
                    .map(|_| u32::try_from(user_columns))
                    .transpose()
                    .map_err(|_| BodyTableHiddenAxesError::InvalidSource)?,
            )
            .with_hidden_state_formula_owner_for_columns(formula_columns_ref)
            .with_hidden_state_formula_owner_for_rows(formula_rows_ref)
            .with_base_column_row_uids(graph.model.map_ref)
            .with_hidden_states_owner(Some(owner.clone())),
    )
}

fn rewrite(
    source: &Package,
    graph: &Graph,
    axes: &HiddenAxes,
    budget: &mut table_lock::WireBudget,
) -> Result<RewriteResult, BodyTableHiddenAxesError> {
    let component = source
        .state
        .source
        .components()
        .get_index(graph.target.model_component_index)
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    if graph.target.component_index != graph.target.model_component_index {
        return Err(BodyTableHiddenAxesError::UnsupportedDependency);
    }
    let component_name = component.name();
    let entry = source
        .state
        .source
        .package()
        .iter()
        .find(|entry| entry.name() == component_name)
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    if entry.is_opaque() {
        return Err(BodyTableHiddenAxesError::UnsupportedSource);
    }
    let archive_limits = source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let creation = graph.model.owner.is_none();
    if graph.profile.is_native() && creation {
        // Native owner creation has no qualified producer envelope. Keep
        // this restriction at the private rewrite boundary as well.
        return Err(BodyTableHiddenAxesError::UnsupportedDependency);
    }
    let creation_plan = if creation {
        Some(prepare_owner_creation(source, graph, axes, budget)?)
    } else {
        None
    };
    let (owner, active_uuid, formula_columns_ref, formula_rows_ref) =
        if let Some(plan) = creation_plan.as_ref() {
            (
                plan.owner.clone(),
                plan.active_uuid,
                Some(plan.formula_columns_ref),
                Some(plan.formula_rows_ref),
            )
        } else {
            let (owner, active_uuid) = desired_owner(graph, axes, budget)?;
            (
                owner,
                active_uuid,
                graph.model.formula_columns_ref,
                graph.model.formula_rows_ref,
            )
        };
    let desired_model = desired_model_snapshot(
        graph,
        &owner,
        active_uuid,
        formula_columns_ref,
        formula_rows_ref,
        budget,
    )?;
    let desired_info = codec::TableInfoSnapshot::new(graph.info.model_ref)
        .with_view_column_row_uids(graph.info.map_ref)
        .with_hidden_states_uuid(Some(active_uuid));
    let source_archive = component.archive();
    let model_message = source_archive
        .objects
        .get(graph.target.model_object_index)
        .and_then(|object| object.messages.get(graph.target.model_message_index))
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    let info_message = source_archive
        .objects
        .get(graph.target.object_index)
        .and_then(|object| object.messages.get(graph.target.info_message_index))
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    if source_archive
        .objects
        .get(graph.target.model_object_index)
        .and_then(|object| object.archive_info.identifier)
        != Some(graph.target.model_identifier.get())
        || source_archive
            .objects
            .get(graph.target.object_index)
            .and_then(|object| object.archive_info.identifier)
            != Some(graph.target.drawable_identifier.get())
    {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    let model_prepared = codec::prepare_table_model_rewrite(
        model_message.data.as_slice(),
        &desired_model,
        codec_options(
            budget,
            model_message.data.len(),
            budget.wire_limits().max_output_bytes(),
        )?,
    )
    .map_err(map_codec_error)?;
    let info_prepared = codec::prepare_table_info_rewrite(
        info_message.data.as_slice(),
        &desired_info,
        codec_options(
            budget,
            info_message.data.len(),
            budget.wire_limits().max_output_bytes(),
        )?,
    )
    .map_err(map_codec_error)?;
    let model_requirements = model_prepared.execution_requirements();
    let info_requirements = info_prepared.execution_requirements();
    charge_rewrite_requirements(budget, model_requirements)?;
    charge_rewrite_requirements(budget, info_requirements)?;

    let metadata_rewritten = if let Some(plan) = creation_plan
        .as_ref()
        .and_then(|creation| creation.metadata.as_ref())
    {
        let object_ids = creation_plan
            .as_ref()
            .ok_or(BodyTableHiddenAxesError::InvalidSource)?
            .ids
            .as_array();
        let additions = [
            ObjectUuidAddition::new(plan.selector, object_ids[0], plan.uuids[0]),
            ObjectUuidAddition::new(plan.selector, object_ids[1], plan.uuids[1]),
            ObjectUuidAddition::new(plan.selector, object_ids[2], plan.uuids[2]),
            ObjectUuidAddition::new(plan.selector, object_ids[3], plan.uuids[3]),
        ];
        let selectors = [plan.selector];
        let batch = AdditionSaveTokenBatch::new(
            MetadataBatch::new(
                plan.expected_last_identifier,
                plan.new_last_identifier,
                &additions,
                &[],
            ),
            SaveTokenBatch::new(&selectors),
        );
        let prepared = prepare_package_metadata_additions_and_save_tokens(
            plan.payload,
            batch,
            metadata_options(source, plan.payload.len(), additions.len()),
        )
        .map_err(map_metadata_error)?;
        charge_metadata_report(budget, prepared.prepare_report())?;
        let requirements = prepared.execution_requirements();
        charge_metadata_requirements(budget, requirements)?;
        let output = prepared
            .execute(requirements.exact_limits())
            .map_err(map_metadata_error)?
            .into_bytes();
        Some(output)
    } else {
        None
    };
    let helper_archive = creation_plan
        .as_ref()
        .map(|plan| build_owner_helper_objects(plan, archive_limits, budget.wire_limits(), budget))
        .transpose()?
        .map(|objects| Archive { objects });
    let helper_archive_length = helper_archive
        .as_ref()
        .map(|archive| archive.encoded_len_with_limits(archive_limits))
        .transpose()
        .map_err(map_core_error)?
        .unwrap_or(0);
    let metadata_replacement_bound = if let Some(payload) = metadata_rewritten.as_ref() {
        let plan = creation_plan
            .as_ref()
            .and_then(|creation| creation.metadata.as_ref())
            .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
        let metadata_component = source
            .state
            .source
            .components()
            .get_index(plan.route.component_index)
            .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
        let metadata_object = metadata_component
            .archive()
            .objects
            .get(plan.route.object_index)
            .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
        let metadata_message = metadata_object
            .messages
            .get(plan.route.message_index)
            .filter(|message| message.type_ == METADATA_MESSAGE_TYPE)
            .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
        let metadata_entry = source
            .state
            .source
            .package()
            .iter()
            .find(|entry| entry.name() == METADATA_ENTRY_NAME)
            .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
        Some(
            metadata_replacement_bounds(
                source,
                metadata_component,
                metadata_message.data.len(),
                payload.len(),
                metadata_entry,
            )?
            .2,
        )
    } else {
        None
    };

    let stream_length = archive_source_length(source_archive)?;
    let old_payload_length = model_message
        .data
        .len()
        .checked_add(info_message.data.len())
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    let new_payload_length = model_requirements
        .output_bytes()
        .checked_add(info_requirements.output_bytes())
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    let rewritten_bound = stream_length
        .checked_sub(old_payload_length)
        .and_then(|value| value.checked_add(new_payload_length))
        .and_then(|value| value.checked_add(helper_archive_length))
        .and_then(|value| value.checked_add(64))
        .ok_or(BodyTableHiddenAxesError::LimitExceeded {
            kind: BodyTableHiddenAxesLimitKind::OutputBytes,
            observed: u64::MAX,
            maximum: source.state.source.limits().max_input_bytes(),
        })?;
    let compressed_bound = table_lock::snappy_compressed_bound(rewritten_bound).ok_or(
        BodyTableHiddenAxesError::LimitExceeded {
            kind: BodyTableHiddenAxesLimitKind::OutputBytes,
            observed: u64::MAX,
            maximum: source.state.source.limits().max_input_bytes(),
        },
    )?;
    let replacement_compressed_bound = match entry.metadata().central().compression_method() {
        0 => compressed_bound,
        8 => table_lock::deflate_compressed_bound(compressed_bound).ok_or(
            BodyTableHiddenAxesError::LimitExceeded {
                kind: BodyTableHiddenAxesLimitKind::OutputBytes,
                observed: u64::MAX,
                maximum: source.state.source.limits().max_input_bytes(),
            },
        )?,
        _ => return Err(BodyTableHiddenAxesError::UnsupportedSource),
    };
    let old_compressed_size =
        usize::try_from(entry.metadata().compressed_size()).map_err(|_| {
            BodyTableHiddenAxesError::LimitExceeded {
                kind: BodyTableHiddenAxesLimitKind::EntryBytes,
                observed: u64::MAX,
                maximum: source.state.source.limits().max_entry_bytes(),
            }
        })?;
    let old_metadata_compressed_size = if metadata_replacement_bound.is_some() {
        let metadata_entry = source
            .state
            .source
            .package()
            .iter()
            .find(|entry| entry.name() == METADATA_ENTRY_NAME)
            .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
        usize::try_from(metadata_entry.metadata().compressed_size()).map_err(|_| {
            BodyTableHiddenAxesError::LimitExceeded {
                kind: BodyTableHiddenAxesLimitKind::EntryBytes,
                observed: u64::MAX,
                maximum: source.state.source.limits().max_entry_bytes(),
            }
        })?
    } else {
        0
    };
    let package_output_bound = source
        .state
        .source
        .source_bytes()
        .len()
        .checked_sub(old_compressed_size)
        .and_then(|value| value.checked_sub(old_metadata_compressed_size))
        .and_then(|value| value.checked_add(replacement_compressed_bound))
        .and_then(|value| value.checked_add(metadata_replacement_bound.unwrap_or(0)))
        .ok_or(BodyTableHiddenAxesError::LimitExceeded {
            kind: BodyTableHiddenAxesLimitKind::OutputBytes,
            observed: u64::MAX,
            maximum: source.state.source.limits().max_input_bytes(),
        })?;
    budget
        .charge_output_bytes(rewritten_bound)
        .and_then(|_| budget.charge_output_bytes(compressed_bound))
        .and_then(|_| budget.charge_output_bytes(replacement_compressed_bound))
        .and_then(|_| budget.charge_output_bytes(package_output_bound))
        .and_then(|_| budget.charge_payload_bytes(rewritten_bound))
        .and_then(|_| budget.charge_total_payload_bytes(rewritten_bound))
        .and_then(|_| budget.charge_payload_work(rewritten_bound))
        .and_then(|_| budget.charge_payload_work(compressed_bound))
        .and_then(|_| budget.charge_payload_work(replacement_compressed_bound))
        .and_then(|_| budget.charge_payload_work(package_output_bound))
        .and_then(|_| budget.charge_payload_work(info_requirements.output_bytes()))
        .map_err(map_lock_error)?;
    budget
        .precharge_candidate_reopen(
            &source.state.source,
            package_output_bound,
            graph.target.model_component_index,
            compressed_bound,
            rewritten_bound,
            graph.target.model_object_index,
            graph.target.model_message_index,
            model_requirements.output_bytes(),
        )
        .map_err(map_lock_error)?;

    let creation_column_reference = creation_plan.as_ref().map(|plan| [plan.ids.column_formula]);
    let creation_row_reference = creation_plan.as_ref().map(|plan| [plan.ids.row_formula]);
    let creation_transition = if let Some(plan) = creation_plan.as_ref() {
        let source_info = source_archive
            .objects
            .get(graph.target.model_object_index)
            .and_then(|object| {
                object
                    .archive_info
                    .message_infos
                    .get(graph.target.model_message_index)
            })
            .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
        let before_len = source_info.object_references.len();
        let after_len = before_len
            .checked_add(OWNER_CREATION_OBJECTS)
            .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
        budget
            .charge_payload_references(OWNER_CREATION_OBJECTS)
            .and_then(|_| budget.charge_payload_items(OWNER_CREATION_FIELD_INFOS))
            .and_then(|_| budget.charge_payload_work(after_len))
            .map_err(map_lock_error)?;
        let mut aggregate_before = Vec::new();
        aggregate_before
            .try_reserve_exact(before_len)
            .map_err(|_| BodyTableHiddenAxesError::Allocation { amount: before_len })?;
        aggregate_before.extend_from_slice(&source_info.object_references);
        let mut aggregate_after = Vec::new();
        aggregate_after
            .try_reserve_exact(after_len)
            .map_err(|_| BodyTableHiddenAxesError::Allocation { amount: after_len })?;
        aggregate_after.extend_from_slice(&aggregate_before);
        aggregate_after.extend_from_slice(&plan.object_ids);
        if aggregate_after.windows(2).any(|pair| pair[0] == pair[1]) {
            return Err(BodyTableHiddenAxesError::InvalidSource);
        }
        let fields = [
            CanonicalObjectReferenceField {
                path: MODEL_COLUMN_REFERENCE_PATH,
                references: creation_column_reference
                    .as_ref()
                    .ok_or(BodyTableHiddenAxesError::InvalidSource)?,
            },
            CanonicalObjectReferenceField {
                path: MODEL_ROW_REFERENCE_PATH,
                references: creation_row_reference
                    .as_ref()
                    .ok_or(BodyTableHiddenAxesError::InvalidSource)?,
            },
        ];
        Some((aggregate_before, aggregate_after, fields))
    } else {
        None
    };

    let mut archive = page_layout::editable_archive(source, component_name)
        .map_err(map_page_layout_error)?
        .0;
    let model_object = archive
        .objects
        .get_mut(graph.target.model_object_index)
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    if model_object.archive_info.identifier != Some(graph.target.model_identifier.get()) {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    let new_model = model_prepared
        .execute(codec::RewriteExecutionLimits::exact(model_requirements))
        .map_err(map_codec_error)?
        .into_bytes();
    let model_message = RawMessage {
        type_: graph.target.model_message_type,
        data: new_model,
    };
    if let Some((aggregate_before, aggregate_after, fields)) = creation_transition.as_ref() {
        model_object
            .replace_message_transitioning_object_references_with_canonical_fields_preserving_header_with_limits(
                graph.target.model_message_index,
                model_message,
                ObjectReferenceTransition {
                    aggregate_before,
                    aggregate_after,
                    fields: &[],
                },
                CanonicalObjectReferenceFields::Insert(fields),
                archive_limits,
            )
            .map_err(map_core_error)?;
    } else {
        model_object
            .replace_message_preserving_header_with_limits(
                graph.target.model_message_index,
                model_message,
                archive_limits,
            )
            .map_err(map_core_error)?;
    }
    let info_object = archive
        .objects
        .get_mut(graph.target.object_index)
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    let new_info = info_prepared
        .execute(codec::RewriteExecutionLimits::exact(info_requirements))
        .map_err(map_codec_error)?
        .into_bytes();
    info_object
        .replace_message_preserving_header_with_limits(
            graph.target.info_message_index,
            RawMessage {
                type_: graph.target.message_type,
                data: new_info,
            },
            archive_limits,
        )
        .map_err(map_core_error)?;
    if let Some(helper_archive) = helper_archive {
        archive
            .append_objects_with_limits(helper_archive.objects, archive_limits)
            .map_err(map_core_error)?;
    }
    let compressed =
        page_layout::compress_archive(archive, archive_limits).map_err(map_page_layout_error)?;
    if compressed.len() > compressed_bound {
        return Err(BodyTableHiddenAxesError::Verification);
    }
    let metadata_compressed = if let Some(payload) = metadata_rewritten {
        let plan = creation_plan
            .as_ref()
            .and_then(|creation| creation.metadata.as_ref())
            .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
        Some(rewrite_metadata_entry_for_creation(
            source,
            plan,
            payload,
            archive_limits,
            budget,
        )?)
    } else {
        None
    };
    let mut previews = Vec::new();
    previews
        .try_reserve_exact(page_layout::PREVIEW_ENTRY_NAMES.len())
        .map_err(|_| BodyTableHiddenAxesError::Allocation {
            amount: page_layout::PREVIEW_ENTRY_NAMES.len(),
        })?;
    for name in page_layout::PREVIEW_ENTRY_NAMES.iter().copied() {
        if source
            .state
            .source
            .package()
            .iter()
            .any(|entry| entry.name() == name)
        {
            previews.push(name);
        }
    }
    let edit_count = 1usize.saturating_add(usize::from(metadata_compressed.is_some()));
    let mut edits = Vec::new();
    edits
        .try_reserve_exact(edit_count)
        .map_err(|_| BodyTableHiddenAxesError::Allocation { amount: edit_count })?;
    edits.push(EntryEdit::new(component_name, &compressed));
    if let Some(metadata_compressed) = metadata_compressed.as_ref() {
        edits.push(EntryEdit::new(METADATA_ENTRY_NAME, metadata_compressed));
    }
    let prepared = source
        .state
        .source
        .package()
        .prepare_reassembly_with_deletions(&edits, &previews, source.state.source.limits())
        .map_err(map_archive_error)?;
    let requirements = prepared.execution_requirements();
    budget
        .charge_output_bytes(prepared.output_bytes())
        .and_then(|_| budget.charge_payload_work(requirements.scratch_bytes()))
        .and_then(|_| budget.charge_payload_work(requirements.allocations()))
        .map_err(map_lock_error)?;
    let output = prepared
        .execute(requirements.exact_limits())
        .map_err(map_archive_error)?;
    if output.len() > package_output_bound {
        return Err(BodyTableHiddenAxesError::Verification);
    }
    let candidate_source =
        SourceCatalog::from_shared_bytes_with_limits(output.into(), source.state.source.limits())
            .map_err(map_archive_error)?;
    let added_object_ids = creation_plan.as_ref().map_or_else(
        || Arc::from(Vec::<u64>::new()),
        |plan| Arc::clone(&plan.object_ids),
    );
    Ok(RewriteResult {
        package: Package::from_source_catalog(candidate_source).map_err(map_package_error)?,
        added_object_ids,
        touched_components: 1usize.saturating_add(usize::from(metadata_compressed.is_some())),
    })
}

fn map_metadata_error(error: MetadataRewriteError) -> BodyTableHiddenAxesError {
    if let Some(amount) = error.allocation_request() {
        return BodyTableHiddenAxesError::Allocation { amount };
    }
    let Some(limit) = error.resource_limit() else {
        return BodyTableHiddenAxesError::InvalidSource;
    };
    let (kind, observed, maximum) = match limit {
        litchi_iwa_protos::package_metadata_codec::RewriteLimit::InputBytes {
            observed,
            maximum,
        } => (BodyTableHiddenAxesLimitKind::WireBytes, observed, maximum),
        litchi_iwa_protos::package_metadata_codec::RewriteLimit::OutputBytes {
            observed,
            maximum,
        } => (
            BodyTableHiddenAxesLimitKind::WireOutputBytes,
            observed,
            maximum,
        ),
        litchi_iwa_protos::package_metadata_codec::RewriteLimit::Fields { observed, maximum } => {
            (BodyTableHiddenAxesLimitKind::WireFields, observed, maximum)
        },
        litchi_iwa_protos::package_metadata_codec::RewriteLimit::Work { observed, maximum } => {
            (BodyTableHiddenAxesLimitKind::WireWork, observed, maximum)
        },
        litchi_iwa_protos::package_metadata_codec::RewriteLimit::Nesting { observed, maximum } => {
            return BodyTableHiddenAxesError::LimitExceeded {
                kind: BodyTableHiddenAxesLimitKind::WireNesting,
                observed: u64::from(observed),
                maximum: u64::from(maximum),
            };
        },
        litchi_iwa_protos::package_metadata_codec::RewriteLimit::Components {
            observed,
            maximum,
        } => (
            BodyTableHiddenAxesLimitKind::PayloadItems,
            observed,
            maximum,
        ),
        litchi_iwa_protos::package_metadata_codec::RewriteLimit::References {
            observed,
            maximum,
        } => (
            BodyTableHiddenAxesLimitKind::PayloadReferences,
            observed,
            maximum,
        ),
        litchi_iwa_protos::package_metadata_codec::RewriteLimit::Additions {
            observed,
            maximum,
        } => (
            BodyTableHiddenAxesLimitKind::PayloadItems,
            observed,
            maximum,
        ),
        _ => return BodyTableHiddenAxesError::InvalidSource,
    };
    BodyTableHiddenAxesError::LimitExceeded {
        kind,
        observed: u64::try_from(observed).unwrap_or(u64::MAX),
        maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
    }
}
