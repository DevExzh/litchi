//! Bounded dependency proof for changed Pages body-table names.
//!
//! A table name is a semantic input to formula calculation.  The table-model
//! payload alone therefore cannot establish that a changed name is safe: a
//! rooted calculation engine may retain volatile name cells or a pivot owner
//! whose selection is name-sensitive.  This module follows only the rooted
//! Pages calculation-engine route, using the shared Buffa dependency codec for
//! the typed envelopes and the table-lock wire budget for all raw framing.
//!
//! A missing calculation-engine route is retained as a valid source profile.
//! Once the route is present, every reachable owner is resolved and bounded;
//! orphan owner objects do not affect a rename.

use std::collections::HashSet;

use litchi_iwa_common::wire::{RawWireFields, RawWireLimits};
use litchi_iwa_core::{ArchiveObject, MessageInfo, RawMessage};
use litchi_iwa_protos::numbers_table_cell_dependency_codec as dependency_codec;

use super::{BodyTableNameError, Package, TABLE_MODEL_MESSAGE_TYPE};
use crate::package::{page_layout, table_lock};

const DOCUMENT_COMPONENT: &str = "Index/Document.iwa";
const ROOT_OBJECT_IDENTIFIER: u64 = 1;
const ROOT_MESSAGE_TYPE: u32 = 10_000;
const DOCUMENT_SUPER_FIELD: u32 = 15;
const CALCULATION_ENGINE_FIELD: u32 = 4;
const CALCULATION_ENGINE_MESSAGE_TYPE: u32 = 4_000;
const DEPENDENCY_TRACKER_FIELD: u32 = 2;
const FORMULA_OWNER_DEPENDENCY_FIELD: u32 = 6;
const FORMULA_OWNER_MESSAGE_TYPE: u32 = 4_008;
const FORMULA_OWNER_REFERENCE_FIELD: u32 = 11;
const MODEL_PIVOT_OWNER_FIELD: u32 = 85;
const VOLATILE_SHEET_TABLE_NAME_FIELD: u32 = 4;
const REFERENCE_PATH_DOCUMENT_ENGINE: &[u32] = &[DOCUMENT_SUPER_FIELD, CALCULATION_ENGINE_FIELD];
const REFERENCE_PATH_ENGINE_OWNER: &[u32] =
    &[DEPENDENCY_TRACKER_FIELD, FORMULA_OWNER_DEPENDENCY_FIELD];
const REFERENCE_PATH_FORMULA_OWNER: &[u32] = &[FORMULA_OWNER_REFERENCE_FIELD];

#[derive(Clone, Copy)]
struct ObjectLocation<'a> {
    component_index: usize,
    object: &'a ArchiveObject,
}

#[derive(Clone, Copy)]
struct MessageLocation<'a> {
    message_index: usize,
    message: &'a RawMessage,
    info: &'a MessageInfo,
}

/// Prove the rooted dependency closure needed before publishing a changed
/// body-table name.  This function is called only for changed operations;
/// read and semantic no-op paths intentionally do not run the guard.
pub(super) fn validate(
    package: &Package,
    target: &table_lock::BodyTableTarget,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableNameError> {
    reject_selected_model_pivot(package, target, budget)?;
    let Some(engine) = rooted_calculation_engine(package, budget)? else {
        return Ok(());
    };

    let options = dependency_options(budget, engine.message.data.len())?;
    budget
        .charge_payload_work(engine.message.data.len())
        .map_err(map_lock_error)?;
    let (engine_snapshot, report) = dependency_codec::decode_calculation_engine_with_report(
        engine.message.data.as_slice(),
        options,
    )
    .map_err(map_dependency_error)?;
    charge_dependency_report(budget, report)?;

    let tracker_source = engine_snapshot.dependency_tracker();
    let tracker_options = dependency_options(budget, tracker_source.len())?;
    budget
        .charge_payload_work(tracker_source.len())
        .map_err(map_lock_error)?;
    let (_tracker_snapshot, tracker_report) =
        dependency_codec::decode_dependency_tracker_with_report(tracker_source, tracker_options)
            .map_err(map_dependency_error)?;
    charge_dependency_report(budget, tracker_report)?;

    let owner_payloads =
        repeated_length_fields(tracker_source, FORMULA_OWNER_DEPENDENCY_FIELD, 1, budget)?;
    let mut owner_ids = HashSet::new();
    owner_ids
        .try_reserve(owner_payloads.len())
        .map_err(|_| BodyTableNameError::Allocation {
            amount: owner_payloads.len(),
        })?;

    for owner_payload in &owner_payloads {
        let owner_identifier = parse_local_reference(owner_payload, budget)?;
        // The tracker snapshot is intentionally retained as a typed Buffa
        // parity proof.  Raw field traversal below establishes the repeated
        // field cardinality and archive-header declaration path.
        if !owner_ids.insert(owner_identifier) {
            return Err(BodyTableNameError::InvalidSource);
        }
        validate_declared_reference(
            engine.info,
            owner_identifier,
            REFERENCE_PATH_ENGINE_OWNER,
            true,
            budget,
        )?;
        let owner = find_object(package, owner_identifier, budget)?;
        let owner_message = unique_message(owner, FORMULA_OWNER_MESSAGE_TYPE, budget)?
            .ok_or(BodyTableNameError::InvalidSource)?;
        validate_formula_owner_metadata(owner_message.info, budget)?;
        validate_formula_owner_payload(package, owner_message, budget)?;
    }

    // A tracker may legally contain no owner records.  The Buffa snapshot is
    // still consumed above so a malformed or truncated tracker cannot be
    // mistaken for an empty dependency graph; ordinary formula IDs do not
    // make a table-name edit unsafe.
    Ok(())
}

fn rooted_calculation_engine<'a>(
    package: &'a Package,
    budget: &mut table_lock::WireBudget,
) -> Result<Option<MessageLocation<'a>>, BodyTableNameError> {
    let mut root_component = None;
    for (component_index, component) in package.state.source.components().iter().enumerate() {
        budget.charge_payload_work(1).map_err(map_lock_error)?;
        if component.name() != DOCUMENT_COMPONENT {
            continue;
        }
        if root_component
            .replace((component_index, component))
            .is_some()
        {
            return Err(BodyTableNameError::InvalidSource);
        }
    }
    let Some((root_component_index, component)) = root_component else {
        return Err(BodyTableNameError::InvalidSource);
    };
    let mut root = None;
    for (object_index, object) in component.archive().objects.iter().enumerate() {
        budget.charge_payload_work(1).map_err(map_lock_error)?;
        if object.archive_info.identifier != Some(ROOT_OBJECT_IDENTIFIER) {
            continue;
        }
        if root.replace((object_index, object)).is_some() {
            return Err(BodyTableNameError::InvalidSource);
        }
    }
    let Some((_root_object_index, root_object)) = root else {
        return Err(BodyTableNameError::InvalidSource);
    };
    let root_location = ObjectLocation {
        component_index: root_component_index,
        object: root_object,
    };
    let root_message = unique_message(root_location, ROOT_MESSAGE_TYPE, budget)?;
    let Some(root_message) = root_message else {
        return Err(BodyTableNameError::InvalidSource);
    };
    let root_info = root_message.info;
    validate_message_metadata(root_object, root_message.message_index, ROOT_MESSAGE_TYPE)?;
    budget
        .charge_payload_work(root_message.message.data.len())
        .map_err(map_lock_error)?;

    let super_payload =
        optional_length_field(&root_message.message.data, DOCUMENT_SUPER_FIELD, 1, budget)?;
    let Some(super_payload) = super_payload else {
        reject_stale_engine_declarations(package, root_info, budget)?;
        return Ok(None);
    };
    let engine_payload = optional_length_field(super_payload, CALCULATION_ENGINE_FIELD, 2, budget)?;
    let Some(engine_payload) = engine_payload else {
        reject_stale_engine_declarations(package, root_info, budget)?;
        return Ok(None);
    };
    let engine_identifier = parse_local_reference(engine_payload, budget)?;
    validate_declared_reference(
        root_info,
        engine_identifier,
        REFERENCE_PATH_DOCUMENT_ENGINE,
        true,
        budget,
    )?;
    let engine_object = find_object(package, engine_identifier, budget)?;
    let engine_component = package
        .state
        .source
        .components()
        .get_index(engine_object.component_index)
        .ok_or(BodyTableNameError::InvalidSource)?;
    if !is_calculation_engine_component(engine_component.name()) {
        return Err(BodyTableNameError::UnsupportedSource);
    }
    let engine_message = unique_message(engine_object, CALCULATION_ENGINE_MESSAGE_TYPE, budget)?
        .ok_or(BodyTableNameError::InvalidSource)?;
    validate_message_metadata(
        engine_object.object,
        engine_message.message_index,
        CALCULATION_ENGINE_MESSAGE_TYPE,
    )?;
    Ok(Some(MessageLocation {
        message_index: engine_message.message_index,
        message: engine_message.message,
        info: engine_message.info,
    }))
}

fn reject_stale_engine_declarations(
    package: &Package,
    info: &MessageInfo,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableNameError> {
    for field in &info.field_infos {
        budget
            .charge_payload_work(field.path.path.len())
            .and_then(|_| budget.charge_payload_references(field.object_references.len()))
            .and_then(|_| budget.charge_payload_references(field.data_references.len()))
            .map_err(map_lock_error)?;
        if field.path.as_slice() == REFERENCE_PATH_DOCUMENT_ENGINE
            && (!field.object_references.is_empty() || !field.data_references.is_empty())
        {
            return Err(BodyTableNameError::InvalidSource);
        }
    }
    for identifier in &info.object_references {
        let Some(location) = find_object_optional(package, *identifier, budget)? else {
            return Err(BodyTableNameError::InvalidSource);
        };
        if unique_message(location, CALCULATION_ENGINE_MESSAGE_TYPE, budget)?.is_some() {
            return Err(BodyTableNameError::InvalidSource);
        }
    }
    Ok(())
}

fn reject_selected_model_pivot(
    package: &Package,
    target: &table_lock::BodyTableTarget,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableNameError> {
    let component = package
        .state
        .source
        .components()
        .get_index(target.model_component_index)
        .ok_or(BodyTableNameError::InvalidSource)?;
    let object = component
        .archive()
        .objects
        .get(target.model_object_index)
        .ok_or(BodyTableNameError::InvalidSource)?;
    if object.archive_info.identifier != Some(target.model_identifier.get()) {
        return Err(BodyTableNameError::InvalidSource);
    }
    let message = object
        .messages
        .get(target.model_message_index)
        .filter(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
        .ok_or(BodyTableNameError::InvalidSource)?;
    validate_message_metadata(object, target.model_message_index, message.type_)?;
    let pivot = pivot_owner_identifier(message.data.as_slice(), budget)?;
    if let Some(identifier) = pivot {
        if find_object_optional(package, identifier, budget)?.is_none() {
            return Err(BodyTableNameError::InvalidSource);
        }
    }
    if pivot.is_some() {
        return Err(BodyTableNameError::UnsupportedSource);
    }
    Ok(())
}

fn pivot_owner_identifier(
    source: &[u8],
    budget: &mut table_lock::WireBudget,
) -> Result<Option<u64>, BodyTableNameError> {
    let remaining_fields = budget.remaining_wire_fields();
    let remaining_work = budget.remaining_wire_work();
    if !source.is_empty() && (remaining_fields == 0 || remaining_work == 0) {
        return Err(BodyTableNameError::LimitExceeded {
            kind: if remaining_fields == 0 {
                super::BodyTableNameLimitKind::WireFields
            } else {
                super::BodyTableNameLimitKind::WireWork
            },
            observed: 1,
            maximum: 0,
        });
    }
    let limits = RawWireLimits::new(
        budget
            .wire_limits()
            .max_input_bytes()
            .min(RawWireLimits::MAX_INPUT_BYTES),
        remaining_fields.clamp(1, RawWireLimits::MAX_FIELDS),
        budget
            .wire_limits()
            .max_nesting()
            .min(RawWireLimits::MAX_DEPTH),
        remaining_work.clamp(1, RawWireLimits::MAX_WORK),
    )
    .map_err(map_raw_wire_error)?;
    budget
        .charge_payload_work(source.len())
        .map_err(map_lock_error)?;
    let mut fields = RawWireFields::with_limits(source, limits);
    let mut pivot = None;
    let mut max_depth = 0usize;
    while let Some(field) = fields.next().map_err(map_raw_wire_error)? {
        max_depth = max_depth.max(
            field
                .depth()
                .checked_add(field.group_depth())
                .ok_or(BodyTableNameError::InvalidSource)?,
        );
        if field.number() != MODEL_PIVOT_OWNER_FIELD {
            continue;
        }
        if field.wire_type() != 2 || pivot.is_some() {
            return Err(BodyTableNameError::InvalidSource);
        }
        if !field.key_is_canonical() || !field.length_is_canonical() {
            return Err(BodyTableNameError::InvalidSource);
        }
        pivot = Some(parse_local_reference(field.payload(), budget)?);
    }
    let max_depth = u32::try_from(max_depth).map_err(|_| BodyTableNameError::LimitExceeded {
        kind: super::BodyTableNameLimitKind::WireNesting,
        observed: u64::MAX,
        maximum: usize_to_u64(budget.wire_limits().max_nesting()),
    })?;
    budget
        .charge_codec_report(fields.fields(), fields.work(), max_depth, 0)
        .map_err(map_lock_error)?;
    Ok(pivot)
}

fn validate_formula_owner_payload(
    package: &Package,
    message: MessageLocation<'_>,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableNameError> {
    let options = dependency_options(budget, message.message.data.len())?;
    budget
        .charge_payload_work(message.message.data.len())
        .map_err(map_lock_error)?;
    let (owner, report) = dependency_codec::decode_formula_owner_dependencies_with_report(
        message.message.data.as_slice(),
        options,
    )
    .map_err(map_dependency_error)?;
    charge_dependency_report(budget, report)?;
    if let Some(reference) = owner.formula_owner() {
        if reference.deprecated_type().is_some() || reference.deprecated_is_external().is_some() {
            return Err(BodyTableNameError::InvalidSource);
        }
        let raw_reference = optional_length_field(
            &message.message.data,
            FORMULA_OWNER_REFERENCE_FIELD,
            1,
            budget,
        )?
        .ok_or(BodyTableNameError::InvalidSource)?;
        let identifier = parse_local_reference(raw_reference, budget)?;
        if identifier != reference.identifier()
            || find_object_optional(package, identifier, budget)?.is_none()
        {
            return Err(BodyTableNameError::InvalidSource);
        }
        validate_declared_reference(
            message.info,
            identifier,
            REFERENCE_PATH_FORMULA_OWNER,
            false,
            budget,
        )?;
    } else {
        // A stale archive-header edge cannot create a dependency merely by
        // omitting the payload field that would explain it.
        validate_absent_reference_path(message.info, REFERENCE_PATH_FORMULA_OWNER, budget)?;
    }
    if let Some(volatile) = owner.volatile_dependencies() {
        if volatile_name_cells_present(volatile, budget)? {
            return Err(BodyTableNameError::UnsupportedSource);
        }
    }
    Ok(())
}

fn validate_formula_owner_metadata(
    info: &MessageInfo,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableNameError> {
    budget
        .charge_payload_work(info.field_infos.len())
        .and_then(|_| budget.charge_payload_references(info.object_references.len()))
        .and_then(|_| budget.charge_payload_references(info.data_references.len()))
        .map_err(map_lock_error)?;
    if !info.data_references.is_empty()
        || info
            .field_infos
            .iter()
            .any(|field| !field.data_references.is_empty())
    {
        return Err(BodyTableNameError::InvalidSource);
    }
    Ok(())
}

fn validate_declared_reference(
    info: &MessageInfo,
    identifier: u64,
    accepted_path: &[u32],
    require_aggregate: bool,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableNameError> {
    budget
        .charge_payload_items(info.field_infos.len())
        .and_then(|_| budget.charge_payload_references(info.object_references.len()))
        .and_then(|_| budget.charge_payload_references(info.data_references.len()))
        .and_then(|_| budget.charge_payload_work(info.field_infos.len()))
        .map_err(map_lock_error)?;
    let aggregate = info
        .object_references
        .iter()
        .filter(|candidate| **candidate == identifier)
        .count();
    if (require_aggregate && aggregate != 1) || (!require_aggregate && aggregate > 1) {
        return Err(BodyTableNameError::InvalidSource);
    }
    if info.data_references.contains(&identifier) {
        return Err(BodyTableNameError::InvalidSource);
    }
    let mut field_occurrences = 0usize;
    for field in &info.field_infos {
        budget
            .charge_payload_work(field.path.path.len())
            .and_then(|_| budget.charge_payload_references(field.object_references.len()))
            .and_then(|_| budget.charge_payload_references(field.data_references.len()))
            .map_err(map_lock_error)?;
        if field.data_references.contains(&identifier) {
            return Err(BodyTableNameError::InvalidSource);
        }
        let occurrences = field
            .object_references
            .iter()
            .filter(|candidate| **candidate == identifier)
            .count();
        if occurrences == 0 {
            continue;
        }
        if occurrences != 1 || field.path.as_slice() != accepted_path {
            return Err(BodyTableNameError::InvalidSource);
        }
        field_occurrences = field_occurrences
            .checked_add(1)
            .ok_or(BodyTableNameError::InvalidSource)?;
        if !field.data_references.is_empty() {
            return Err(BodyTableNameError::InvalidSource);
        }
    }
    if field_occurrences > 1 {
        return Err(BodyTableNameError::InvalidSource);
    }
    Ok(())
}

fn validate_absent_reference_path(
    info: &MessageInfo,
    path: &[u32],
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableNameError> {
    for field in &info.field_infos {
        budget
            .charge_payload_work(field.path.path.len())
            .and_then(|_| budget.charge_payload_references(field.object_references.len()))
            .and_then(|_| budget.charge_payload_references(field.data_references.len()))
            .map_err(map_lock_error)?;
        if field.path.as_slice() == path
            && (!field.object_references.is_empty() || !field.data_references.is_empty())
        {
            return Err(BodyTableNameError::InvalidSource);
        }
    }
    Ok(())
}

fn volatile_name_cells_present(
    source: &[u8],
    budget: &mut table_lock::WireBudget,
) -> Result<bool, BodyTableNameError> {
    let view = budget.parse(source, 2).map_err(map_lock_error)?;
    let mut seen = [false; 8];
    let mut present = false;
    for field in view.fields() {
        field
            .validate_canonical_framing()
            .map_err(|_| BodyTableNameError::InvalidSource)?;
        let index = match field.number() {
            1 => 0,
            2 => 1,
            3 => 2,
            4 => 3,
            5 => 4,
            7 => 6,
            _ => return Err(BodyTableNameError::InvalidSource),
        };
        if field.wire_type() != 2 || seen[index] {
            return Err(BodyTableNameError::InvalidSource);
        }
        seen[index] = true;
        if field.number() != VOLATILE_SHEET_TABLE_NAME_FIELD {
            continue;
        }
        let nested = budget.parse(field.payload(), 3).map_err(map_lock_error)?;
        for column in nested.fields() {
            column
                .validate_canonical_framing()
                .map_err(|_| BodyTableNameError::InvalidSource)?;
            if column.number() != 1 || column.wire_type() != 2 {
                return Err(BodyTableNameError::InvalidSource);
            }
            let entry = budget.parse(column.payload(), 4).map_err(map_lock_error)?;
            let mut column_seen = false;
            let mut row_set_seen = false;
            for entry_field in entry.fields() {
                entry_field
                    .validate_canonical_framing()
                    .map_err(|_| BodyTableNameError::InvalidSource)?;
                match entry_field.number() {
                    1 if entry_field.wire_type() == 0 && !column_seen => {
                        column_seen = true;
                    },
                    2 if entry_field.wire_type() == 2 && !row_set_seen => {
                        row_set_seen = true;
                    },
                    _ => return Err(BodyTableNameError::InvalidSource),
                }
            }
            if !column_seen || !row_set_seen {
                return Err(BodyTableNameError::InvalidSource);
            }
            present = true;
        }
    }
    Ok(present)
}

fn repeated_length_fields<'a>(
    source: &'a [u8],
    number: u32,
    depth: usize,
    budget: &mut table_lock::WireBudget,
) -> Result<Vec<&'a [u8]>, BodyTableNameError> {
    let view = budget.parse(source, depth).map_err(map_lock_error)?;
    let count = view
        .fields()
        .filter(|field| field.number() == number)
        .count();
    budget
        .charge_payload_items(count)
        .and_then(|_| budget.charge_payload_work(count))
        .map_err(map_lock_error)?;
    let mut values = Vec::new();
    values
        .try_reserve_exact(count)
        .map_err(|_| BodyTableNameError::Allocation { amount: count })?;
    for field in view.fields().filter(|field| field.number() == number) {
        field
            .validate_canonical_framing()
            .map_err(|_| BodyTableNameError::InvalidSource)?;
        if field.wire_type() != 2 {
            return Err(BodyTableNameError::InvalidSource);
        }
        values.push(field.payload());
    }
    Ok(values)
}

fn optional_length_field<'a>(
    source: &'a [u8],
    number: u32,
    depth: usize,
    budget: &mut table_lock::WireBudget,
) -> Result<Option<&'a [u8]>, BodyTableNameError> {
    let values = repeated_length_fields(source, number, depth, budget)?;
    match values.as_slice() {
        [] => Ok(None),
        [value] => Ok(Some(*value)),
        _ => Err(BodyTableNameError::InvalidSource),
    }
}

fn parse_local_reference(
    source: &[u8],
    budget: &mut table_lock::WireBudget,
) -> Result<u64, BodyTableNameError> {
    table_lock::parse_local_reference(budget, source, 2)
        .map(|identifier| identifier.get())
        .map_err(map_lock_error)
}

fn find_object<'a>(
    package: &'a Package,
    identifier: u64,
    budget: &mut table_lock::WireBudget,
) -> Result<ObjectLocation<'a>, BodyTableNameError> {
    find_object_optional(package, identifier, budget)?.ok_or(BodyTableNameError::InvalidSource)
}

fn find_object_optional<'a>(
    package: &'a Package,
    identifier: u64,
    budget: &mut table_lock::WireBudget,
) -> Result<Option<ObjectLocation<'a>>, BodyTableNameError> {
    let mut found = None;
    for (component_index, component) in package.state.source.components().iter().enumerate() {
        budget.charge_payload_work(1).map_err(map_lock_error)?;
        for object in &component.archive().objects {
            budget.charge_payload_work(1).map_err(map_lock_error)?;
            if object.archive_info.identifier != Some(identifier) {
                continue;
            }
            if found.is_some() {
                return Err(BodyTableNameError::InvalidSource);
            }
            found = Some(ObjectLocation {
                component_index,
                object,
            });
        }
    }
    Ok(found)
}

fn unique_message<'a>(
    object: ObjectLocation<'a>,
    message_type: u32,
    budget: &mut table_lock::WireBudget,
) -> Result<Option<MessageLocation<'a>>, BodyTableNameError> {
    let mut found = None;
    for (message_index, message) in object.object.messages.iter().enumerate() {
        budget.charge_payload_work(1).map_err(map_lock_error)?;
        if message.type_ != message_type {
            continue;
        }
        let info = object
            .object
            .archive_info
            .message_infos
            .get(message_index)
            .ok_or(BodyTableNameError::InvalidSource)?;
        if found.is_some() {
            return Err(BodyTableNameError::InvalidSource);
        }
        found = Some(MessageLocation {
            message_index,
            message,
            info,
        });
    }
    Ok(found)
}

fn validate_message_metadata(
    object: &ArchiveObject,
    message_index: usize,
    message_type: u32,
) -> Result<(), BodyTableNameError> {
    let message = object
        .messages
        .get(message_index)
        .ok_or(BodyTableNameError::InvalidSource)?;
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(BodyTableNameError::InvalidSource)?;
    if message.type_ != message_type
        || info.type_ != message.type_
        || usize::try_from(info.length).ok() != Some(message.data.len())
        || object.header_length == 0
    {
        return Err(BodyTableNameError::InvalidSource);
    }
    page_layout::validate_selected_metadata(object, message_index)
        .map_err(super::map_page_layout_error)
}

fn is_calculation_engine_component(name: &str) -> bool {
    let Some(name) = name.strip_prefix("Index/") else {
        return false;
    };
    let Some(name) = name.strip_suffix(".iwa") else {
        return false;
    };
    name == "CalculationEngine"
        || name
            .strip_prefix("CalculationEngine-")
            .is_some_and(|suffix| !suffix.is_empty())
}

fn dependency_options(
    budget: &table_lock::WireBudget,
    source_len: usize,
) -> Result<dependency_codec::DecodeOptions, BodyTableNameError> {
    let limits = budget.wire_limits();
    if source_len > limits.max_input_bytes() {
        return Err(BodyTableNameError::LimitExceeded {
            kind: super::BodyTableNameLimitKind::WireBytes,
            observed: usize_to_u64(source_len),
            maximum: usize_to_u64(limits.max_input_bytes()),
        });
    }
    let source_budget = source_len.max(1);
    let remaining_fields = budget.remaining_wire_fields();
    let remaining_work = budget.remaining_wire_work();
    let remaining_references = budget.remaining_payload_references();
    if source_len != 0 && remaining_fields == 0 {
        return Err(BodyTableNameError::LimitExceeded {
            kind: super::BodyTableNameLimitKind::WireFields,
            observed: 1,
            maximum: 0,
        });
    }
    if source_len != 0 && remaining_work == 0 {
        return Err(BodyTableNameError::LimitExceeded {
            kind: super::BodyTableNameLimitKind::WireWork,
            observed: 1,
            maximum: 0,
        });
    }
    if source_len != 0 && remaining_references == 0 {
        return Err(BodyTableNameError::LimitExceeded {
            kind: super::BodyTableNameLimitKind::PayloadReferences,
            observed: 1,
            maximum: 0,
        });
    }
    let field_budget = source_len
        .saturating_mul(8)
        .clamp(1, limits.max_fields())
        .min(remaining_fields.max(1));
    let work_budget = source_len
        .saturating_mul(128)
        .clamp(1, limits.max_rewrite_work())
        .min(remaining_work.max(1));
    let reference_budget = source_len.max(1).min(remaining_references.max(1));
    Ok(dependency_codec::DecodeOptions::new(
        source_budget,
        field_budget,
        work_budget,
        u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX),
        reference_budget,
        source_budget,
    ))
}

fn charge_dependency_report(
    budget: &mut table_lock::WireBudget,
    report: dependency_codec::DecodeReport,
) -> Result<(), BodyTableNameError> {
    budget
        .charge_codec_report(
            report.fields(),
            report.work_bytes(),
            report.max_depth(),
            report.references(),
        )
        .map_err(map_lock_error)?;
    let retained = report
        .reference_bytes()
        .checked_add(report.text_bytes())
        .ok_or(BodyTableNameError::InvalidSource)?;
    budget.charge_payload_work(retained).map_err(map_lock_error)
}

fn map_dependency_error(error: dependency_codec::DecodeError) -> BodyTableNameError {
    if let Some(amount) = error.allocation_requested() {
        return BodyTableNameError::Allocation { amount };
    }
    let Some(limit) = error.resource_limit() else {
        return BodyTableNameError::InvalidSource;
    };
    let (kind, observed, maximum) = match limit {
        dependency_codec::DecodeLimit::Bytes { observed, maximum } => {
            (super::BodyTableNameLimitKind::WireBytes, observed, maximum)
        },
        dependency_codec::DecodeLimit::References { observed, maximum } => (
            super::BodyTableNameLimitKind::PayloadReferences,
            observed,
            maximum,
        ),
        dependency_codec::DecodeLimit::Text { observed, maximum } => (
            super::BodyTableNameLimitKind::PayloadBytes,
            observed,
            maximum,
        ),
        dependency_codec::DecodeLimit::Fields { observed, maximum } => {
            (super::BodyTableNameLimitKind::WireFields, observed, maximum)
        },
        dependency_codec::DecodeLimit::Work { observed, maximum } => {
            (super::BodyTableNameLimitKind::WireWork, observed, maximum)
        },
        dependency_codec::DecodeLimit::Nesting { observed, maximum } => (
            super::BodyTableNameLimitKind::WireNesting,
            observed as usize,
            maximum as usize,
        ),
        dependency_codec::DecodeLimit::Retained { observed, maximum } => (
            super::BodyTableNameLimitKind::PayloadBytes,
            observed,
            maximum,
        ),
        dependency_codec::DecodeLimit::Allocation { requested } => {
            return BodyTableNameError::Allocation { amount: requested };
        },
        _ => return BodyTableNameError::InvalidSource,
    };
    BodyTableNameError::LimitExceeded {
        kind,
        observed: usize_to_u64(observed),
        maximum: usize_to_u64(maximum),
    }
}

fn map_raw_wire_error(error: litchi_iwa_common::Error) -> BodyTableNameError {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => BodyTableNameError::LimitExceeded {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes => {
                    super::BodyTableNameLimitKind::WireBytes
                },
                litchi_iwa_common::LimitKind::OutputBytes => {
                    super::BodyTableNameLimitKind::WireOutputBytes
                },
                litchi_iwa_common::LimitKind::Fields => super::BodyTableNameLimitKind::WireFields,
                litchi_iwa_common::LimitKind::Nesting => super::BodyTableNameLimitKind::WireNesting,
                litchi_iwa_common::LimitKind::RewriteWork => {
                    super::BodyTableNameLimitKind::WireWork
                },
                litchi_iwa_common::LimitKind::TableRows
                | litchi_iwa_common::LimitKind::TableColumns
                | litchi_iwa_common::LimitKind::TableCells
                | litchi_iwa_common::LimitKind::MaterializedCells => {
                    super::BodyTableNameLimitKind::PayloadItems
                },
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => {
            BodyTableNameError::Allocation { amount }
        },
        litchi_iwa_common::Error::InvalidFormat(_)
        | litchi_iwa_common::Error::InvalidLimit { .. } => BodyTableNameError::InvalidSource,
    }
}

fn map_lock_error(error: table_lock::BodyTableLockError) -> BodyTableNameError {
    super::map_lock_error(error)
}

fn usize_to_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

#[cfg(test)]
mod tests {
    use litchi_iwa_archive::Limits;

    use super::{
        is_calculation_engine_component, optional_length_field, pivot_owner_identifier, table_lock,
        volatile_name_cells_present,
    };

    fn budget() -> table_lock::WireBudget {
        table_lock::WireBudget::new(Limits::default()).expect("default wire budget")
    }

    #[test]
    fn calculation_engine_component_names_require_index_and_nonempty_suffix() {
        assert!(is_calculation_engine_component(
            "Index/CalculationEngine.iwa"
        ));
        assert!(is_calculation_engine_component(
            "Index/CalculationEngine-176.iwa"
        ));
        assert!(!is_calculation_engine_component(
            "CalculationEngine-176.iwa"
        ));
        assert!(!is_calculation_engine_component(
            "Index/CalculationEngine-.iwa"
        ));
        assert!(!is_calculation_engine_component(
            "Index/CalculationEngine.txt"
        ));
    }

    #[test]
    fn volatile_name_cells_distinguish_empty_and_active_sets() {
        let empty = [0x22, 0x00];
        assert_eq!(
            volatile_name_cells_present(&empty, &mut budget()).expect("empty volatile set"),
            false
        );

        // field 4 -> CellCoordSetArchive.field 1 -> ColumnEntry with its
        // required column and row-set fields.
        let active = [0x22, 0x06, 0x0a, 0x04, 0x08, 0x01, 0x12, 0x00];
        assert_eq!(
            volatile_name_cells_present(&active, &mut budget()).expect("active volatile set"),
            true
        );

        let unknown = [0x08, 0x00];
        assert!(matches!(
            volatile_name_cells_present(&unknown, &mut budget()),
            Err(super::BodyTableNameError::InvalidSource)
        ));
    }

    #[test]
    fn pivot_owner_requires_one_canonical_local_reference() {
        let absent = [];
        assert_eq!(
            pivot_owner_identifier(&absent, &mut budget()).unwrap(),
            None
        );

        let present = [0xaa, 0x05, 0x02, 0x08, 0x01];
        assert_eq!(
            pivot_owner_identifier(&present, &mut budget()).unwrap(),
            Some(1)
        );

        // An unrelated start/end group is preserved by the name rewrite and
        // must not make the owned field-85 probe reject the source.
        let unknown_group = [0x0b, 0x10, 0x01, 0x0c, 0xaa, 0x05, 0x02, 0x08, 0x01];
        assert_eq!(
            pivot_owner_identifier(&unknown_group, &mut budget()).unwrap(),
            Some(1)
        );

        let duplicate = [0xaa, 0x05, 0x02, 0x08, 0x01, 0xaa, 0x05, 0x02, 0x08, 0x01];
        assert!(matches!(
            pivot_owner_identifier(&duplicate, &mut budget()),
            Err(super::BodyTableNameError::InvalidSource)
        ));
    }

    #[test]
    fn rooted_engine_route_rejects_duplicate_optional_edges() {
        // TP.DocumentArchive.super is field 15 and TSA.DocumentArchive's
        // calculation_engine is field 4.  Each is singular in its envelope.
        let duplicate_outer = [0x7a, 0x00, 0x7a, 0x00];
        assert!(matches!(
            optional_length_field(&duplicate_outer, 15, 1, &mut budget()),
            Err(super::BodyTableNameError::InvalidSource)
        ));

        let duplicate_engine = [0x22, 0x00, 0x22, 0x00];
        assert!(matches!(
            optional_length_field(&duplicate_engine, 4, 2, &mut budget()),
            Err(super::BodyTableNameError::InvalidSource)
        ));
    }
}
