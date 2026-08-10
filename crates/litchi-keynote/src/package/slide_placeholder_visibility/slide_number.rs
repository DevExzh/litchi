//! Strict slide-number ownership, storage, and node closure validation.

use std::collections::{HashMap, HashSet};

use litchi_iwa_common::{WireLimits, wire::WireView};
use litchi_iwa_core::ArchiveObject;
use litchi_iwa_protos::keynote_slide_number_codec;

use super::{
    Error, LimitKind, Package, SLIDE_MESSAGE_TYPE, SLIDE_NODE_MESSAGE_TYPE, TransactionBudget,
    canonical_varint, charge_reference_metadata_scan, map_wire_error, nested_unique_field,
    physical_catalog, reject_groups, selected_message, strict_reference,
    strict_reference_with_zero, validate_merge_metadata, validate_reference_metadata,
};

pub(super) const PLACEHOLDER_FIELD: u32 = 20;
const STORAGE_MESSAGE_TYPE: u32 = 2_001;
const ATTACHMENT_MESSAGE_TYPE: u32 = 2_043;

#[derive(Clone, Copy)]
enum StorageDependencyPath {
    StyleSheet,
    AttributeTable(u32),
}

struct StorageDependencyDeclaration {
    path: StorageDependencyPath,
    aggregate_occurrences: usize,
    field_attributed: bool,
}

impl StorageDependencyPath {
    fn matches(self, path: &[u32]) -> bool {
        match self {
            Self::StyleSheet => path == [2],
            Self::AttributeTable(field) => path == [field, 1, 2],
        }
    }
}

impl TransactionBudget {
    const fn remaining_work(&self) -> usize {
        self.maximum_work.saturating_sub(self.work)
    }

    const fn remaining_fields(&self) -> usize {
        self.maximum_fields.saturating_sub(self.fields)
    }

    fn charge_decode_report(
        &mut self,
        report: keynote_slide_number_codec::DecodeReport,
    ) -> Result<(), Error> {
        self.charge_fields(report.fields())?;
        self.charge_work(report.work_bytes())
    }

    fn map_decode_error(&self, error: keynote_slide_number_codec::DecodeError) -> Error {
        match error.resource_limit() {
            Some(keynote_slide_number_codec::DecodeLimit::Bytes { observed, maximum }) => {
                Error::LimitExceeded {
                    kind: LimitKind::WireBytes,
                    observed: observed as u64,
                    maximum: maximum as u64,
                }
            },
            Some(keynote_slide_number_codec::DecodeLimit::Fields {
                observed,
                maximum: _,
            }) => Error::LimitExceeded {
                kind: LimitKind::WireFields,
                observed: self.fields.saturating_add(observed) as u64,
                maximum: self.maximum_fields as u64,
            },
            Some(keynote_slide_number_codec::DecodeLimit::Work {
                observed,
                maximum: _,
            }) => Error::LimitExceeded {
                kind: LimitKind::WireWork,
                observed: self.work.saturating_add(observed) as u64,
                maximum: self.maximum_work as u64,
            },
            Some(keynote_slide_number_codec::DecodeLimit::Nesting { observed, maximum }) => {
                Error::LimitExceeded {
                    kind: LimitKind::WireNesting,
                    observed: u64::from(observed),
                    maximum: u64::from(maximum),
                }
            },
            None => Error::InvalidSource,
        }
    }
}

fn decode_options(
    payload: &[u8],
    limits: WireLimits,
    budget: &TransactionBudget,
) -> keynote_slide_number_codec::DecodeOptions {
    keynote_slide_number_codec::DecodeOptions::new(
        payload.len().min(limits.max_input_bytes()),
        budget.remaining_fields(),
        budget.remaining_work(),
        u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX),
    )
}

pub(super) fn node_visible(
    payload: &[u8],
    limits: WireLimits,
    budget: &mut TransactionBudget,
) -> Result<bool, Error> {
    let (projection, report) = keynote_slide_number_codec::decode_slide_number_node_with_report(
        payload,
        decode_options(payload, limits, budget),
    )
    .map_err(|error| budget.map_decode_error(error))?;
    budget.charge_decode_report(report)?;
    Ok(projection.visibility().unwrap_or(false))
}

pub(super) fn validate_node_references(
    node: &ArchiveObject,
    payload: &[u8],
    reserved_identifiers: &HashSet<u64>,
    limits: WireLimits,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    let (message_index, _) = selected_message(node, SLIDE_NODE_MESSAGE_TYPE)?;
    let scan_work = payload.len().checked_mul(2).ok_or(Error::InvalidSource)?;
    budget.charge_work(scan_work)?;
    charge_reference_metadata_scan(node, message_index, budget)?;
    let view = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let relevant_count = view
        .fields()
        .filter(|field| matches!(field.number(), 1 | 3 | 9))
        .count();
    budget.charge_references(relevant_count)?;
    let mut identifiers = HashMap::new();
    identifiers
        .try_reserve(relevant_count)
        .map_err(|_allocation| Error::Allocation {
            amount: relevant_count,
        })?;
    for field in view.fields() {
        if matches!(field.number(), 1 | 3 | 9) {
            let identifier = strict_reference(field, limits)?;
            if reserved_identifiers.contains(&identifier)
                || identifiers
                    .insert(identifier, (field.number(), 0usize, false))
                    .is_some()
            {
                return Err(Error::InvalidSource);
            }
        }
    }
    let info = node
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(Error::InvalidSource)?;
    for identifier in &info.object_references {
        if let Some((_role, count, _path_seen)) = identifiers.get_mut(identifier) {
            *count = count.checked_add(1).ok_or(Error::InvalidSource)?;
        }
    }
    for field in &info.field_infos {
        for identifier in &field.object_references {
            let Some((role, _count, path_seen)) = identifiers.get_mut(identifier) else {
                continue;
            };
            if field.path.as_slice() != [*role] || std::mem::replace(path_seen, true) {
                return Err(Error::InvalidSource);
            }
        }
    }
    if identifiers
        .values()
        .any(|(_role, count, _path_seen)| *count != 1)
    {
        return Err(Error::InvalidSource);
    }
    Ok(())
}

pub(super) fn validate_global_ownership(
    package: &Package,
    selected_slide_identifier: u64,
    placeholder_identifier: u64,
    limits: WireLimits,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    let catalog = physical_catalog(package)?;
    let mut selected_occurrences = 0usize;
    for component in catalog.components().iter() {
        budget.charge_work(1)?;
        for object in &component.archive().objects {
            budget.charge_work(1)?;
            for (message_index, message) in object.messages.iter().enumerate() {
                budget.charge_work(1)?;
                if message.type_ != SLIDE_MESSAGE_TYPE {
                    continue;
                }
                budget.charge_work(message.data.len())?;
                let view =
                    WireView::parse_with_limits(&message.data, limits).map_err(map_wire_error)?;
                let mut slide_number = None;
                for field in view.fields() {
                    field.validate_canonical_framing().map_err(map_wire_error)?;
                    if matches!(field.wire_type(), 3 | 4) {
                        return Err(Error::InvalidSource);
                    }
                    if field.number() == PLACEHOLDER_FIELD {
                        budget.charge_reference()?;
                        let identifier = strict_reference(field, limits)?;
                        if slide_number.replace(identifier).is_some() {
                            return Err(Error::InvalidSource);
                        }
                    }
                }
                if let Some(identifier) = slide_number {
                    charge_reference_metadata_scan(object, message_index, budget)?;
                    validate_reference_metadata(
                        object,
                        message_index,
                        identifier,
                        &[PLACEHOLDER_FIELD],
                    )?;
                    if identifier == placeholder_identifier {
                        if object.archive_info.identifier != Some(selected_slide_identifier) {
                            return Err(Error::InvalidSource);
                        }
                        selected_occurrences = selected_occurrences
                            .checked_add(1)
                            .ok_or(Error::InvalidSource)?;
                    }
                }
            }
        }
    }
    if selected_occurrences != 1 {
        return Err(Error::InvalidSource);
    }
    Ok(())
}

pub(super) fn placeholder_owner(
    payload: &[u8],
    limits: WireLimits,
) -> Result<(Option<i32>, u64), Error> {
    reject_groups(payload, limits)?;
    let kind = nested_unique_field(payload, &[2], limits)?
        .map(canonical_varint)
        .transpose()?
        .and_then(|value| i32::try_from(value).ok());
    let deprecated = nested_unique_field(payload, &[1, 2], limits)?
        .map(|field| strict_reference_with_zero(field, limits, true))
        .transpose()?;
    let modern = nested_unique_field(payload, &[1, 4], limits)?
        .map(|field| strict_reference_with_zero(field, limits, true))
        .transpose()?;
    if deprecated.is_some() && modern.is_some() && deprecated != modern {
        return Err(Error::InvalidSource);
    }
    Ok((kind, modern.or(deprecated).ok_or(Error::InvalidSource)?))
}

fn visit_storage_dependencies(
    payload: &[u8],
    limits: WireLimits,
    mut visit: impl FnMut(u64, StorageDependencyPath) -> Result<(), Error>,
) -> Result<(), Error> {
    const ATTRIBUTE_TABLE_FIELDS: [u32; 10] = [5, 7, 8, 11, 12, 15, 16, 17, 18, 21];

    let storage = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    for field in storage.fields() {
        field.validate_canonical_framing().map_err(map_wire_error)?;
        if matches!(field.wire_type(), 3 | 4) {
            return Err(Error::InvalidSource);
        }
        if field.number() == 2 {
            visit(
                strict_reference(field, limits)?,
                StorageDependencyPath::StyleSheet,
            )?;
            continue;
        }
        if !ATTRIBUTE_TABLE_FIELDS.contains(&field.number()) {
            continue;
        }
        if field.wire_type() != 2 {
            return Err(Error::InvalidSource);
        }
        let table = WireView::parse_with_limits(field.payload(), limits).map_err(map_wire_error)?;
        for entry in table.fields() {
            entry.validate_canonical_framing().map_err(map_wire_error)?;
            if matches!(entry.wire_type(), 3 | 4) {
                return Err(Error::InvalidSource);
            }
            if entry.number() != 1 {
                continue;
            }
            if entry.wire_type() != 2 {
                return Err(Error::InvalidSource);
            }
            let members =
                WireView::parse_with_limits(entry.payload(), limits).map_err(map_wire_error)?;
            for member in members.fields() {
                member
                    .validate_canonical_framing()
                    .map_err(map_wire_error)?;
                if matches!(member.wire_type(), 3 | 4) {
                    return Err(Error::InvalidSource);
                }
                if member.number() == 2 {
                    visit(
                        strict_reference(member, limits)?,
                        StorageDependencyPath::AttributeTable(field.number()),
                    )?;
                }
            }
        }
    }
    Ok(())
}

fn validate_storage_reference_metadata(
    storage: &ArchiveObject,
    message_index: usize,
    attachment_identifier: u64,
    dependencies: &mut HashMap<u64, StorageDependencyDeclaration>,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    charge_reference_metadata_scan(storage, message_index, budget)?;
    let info = storage
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(Error::InvalidSource)?;
    validate_merge_metadata(storage, info)?;
    let mut attachment_occurrences = 0usize;
    for identifier in &info.object_references {
        if *identifier == attachment_identifier {
            attachment_occurrences = attachment_occurrences
                .checked_add(1)
                .ok_or(Error::InvalidSource)?;
        } else if let Some(dependency) = dependencies.get_mut(identifier) {
            dependency.aggregate_occurrences = dependency
                .aggregate_occurrences
                .checked_add(1)
                .ok_or(Error::InvalidSource)?;
        }
    }
    let mut attachment_attributed = false;
    for field in &info.field_infos {
        for identifier in &field.object_references {
            if *identifier == attachment_identifier {
                if field.path.as_slice() != [9, 1, 2]
                    || std::mem::replace(&mut attachment_attributed, true)
                {
                    return Err(Error::InvalidSource);
                }
            } else if let Some(dependency) = dependencies.get_mut(identifier)
                && (!dependency.path.matches(field.path.as_slice())
                    || std::mem::replace(&mut dependency.field_attributed, true))
            {
                return Err(Error::InvalidSource);
            }
        }
    }
    if attachment_occurrences != 1
        || dependencies.values().any(|dependency| {
            dependency.aggregate_occurrences > 1
                || (dependency.field_attributed && dependency.aggregate_occurrences != 1)
        })
    {
        return Err(Error::InvalidSource);
    }
    Ok(())
}

pub(super) fn validate_storage(
    package: &Package,
    slide_component: &str,
    storage_identifier: u64,
    reserved_identifiers: &mut HashSet<u64>,
    limits: WireLimits,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    let (storage_component, storage) = package
        .object_with_component(storage_identifier)
        .ok_or(Error::InvalidSource)?;
    if storage_component != slide_component
        || storage.messages.len() != 1
        || storage.archive_info.message_infos.len() != 1
        || storage.messages[0].type_ != STORAGE_MESSAGE_TYPE
    {
        return Err(Error::UnsupportedSource);
    }
    validate_merge_metadata(storage, &storage.archive_info.message_infos[0])?;
    let payload = storage.messages[0].data.as_slice();
    let (storage_projection, storage_report) =
        keynote_slide_number_codec::decode_slide_number_storage_with_report(
            payload,
            decode_options(payload, limits, budget),
        )
        .map_err(|error| budget.map_decode_error(error))?;
    budget.charge_decode_report(storage_report)?;
    if !matches!(storage_projection.kind(), None | Some(3))
        || storage_projection.in_document() != Some(true)
    {
        return Err(Error::UnsupportedSource);
    }
    reject_groups(payload, limits)?;
    let dependency_scan_work = payload.len().checked_mul(6).ok_or(Error::InvalidSource)?;
    budget.charge_work(dependency_scan_work)?;
    let mut dependency_count = 0usize;
    visit_storage_dependencies(payload, limits, |_identifier, _path| {
        dependency_count = dependency_count
            .checked_add(1)
            .ok_or(Error::InvalidSource)?;
        Ok(())
    })?;
    budget.charge_references(dependency_count)?;
    let mut storage_dependencies = HashMap::new();
    storage_dependencies
        .try_reserve(dependency_count)
        .map_err(|_allocation| Error::Allocation {
            amount: dependency_count,
        })?;
    visit_storage_dependencies(payload, limits, |identifier, path| {
        if storage_dependencies
            .insert(
                identifier,
                StorageDependencyDeclaration {
                    path,
                    aggregate_occurrences: 0,
                    field_attributed: false,
                },
            )
            .is_some()
        {
            return Err(Error::InvalidSource);
        }
        Ok(())
    })?;
    let attachment_table_scan_work = payload.len().checked_mul(3).ok_or(Error::InvalidSource)?;
    budget.charge_work(attachment_table_scan_work)?;
    let table = nested_unique_field(payload, &[9], limits)?.ok_or(Error::InvalidSource)?;
    if table.wire_type() != 2 || storage_projection.attachment_table() != Some(table.payload()) {
        return Err(Error::InvalidSource);
    }
    let table_view =
        WireView::parse_with_limits(table.payload(), limits).map_err(map_wire_error)?;
    let mut entry = None;
    for field in table_view.fields() {
        field.validate_canonical_framing().map_err(map_wire_error)?;
        if matches!(field.wire_type(), 3 | 4) {
            return Err(Error::InvalidSource);
        }
        if field.number() == 1 && entry.replace(field).is_some() {
            return Err(Error::UnsupportedSource);
        }
    }
    let attachment_entry_field = entry.ok_or(Error::InvalidSource)?;
    if attachment_entry_field.wire_type() != 2 {
        return Err(Error::InvalidSource);
    }
    let entry_view = WireView::parse_with_limits(attachment_entry_field.payload(), limits)
        .map_err(map_wire_error)?;
    let mut character_index = None;
    let mut attachment_identifier = None;
    for field in entry_view.fields() {
        field.validate_canonical_framing().map_err(map_wire_error)?;
        if matches!(field.wire_type(), 3 | 4) {
            return Err(Error::InvalidSource);
        }
        match field.number() {
            1 if character_index.replace(canonical_varint(field)?).is_some() => {
                return Err(Error::InvalidSource);
            },
            2 => {
                budget.charge_reference()?;
                if attachment_identifier
                    .replace(strict_reference(field, limits)?)
                    .is_some()
                {
                    return Err(Error::InvalidSource);
                }
            },
            _ => {},
        }
    }
    if character_index != Some(0) {
        return Err(Error::InvalidSource);
    }
    let resolved_attachment_identifier = attachment_identifier.ok_or(Error::InvalidSource)?;
    validate_storage_reference_metadata(
        storage,
        0,
        resolved_attachment_identifier,
        &mut storage_dependencies,
        budget,
    )?;
    let (attachment_component, attachment) = package
        .object_with_component(resolved_attachment_identifier)
        .ok_or(Error::InvalidSource)?;
    if attachment_component != slide_component
        || attachment.messages.len() != 1
        || attachment.archive_info.message_infos.len() != 1
        || attachment.messages[0].type_ != ATTACHMENT_MESSAGE_TYPE
    {
        return Err(Error::UnsupportedSource);
    }
    validate_merge_metadata(attachment, &attachment.archive_info.message_infos[0])?;
    charge_reference_metadata_scan(attachment, 0, budget)?;
    let attachment_info = attachment
        .archive_info
        .message_infos
        .first()
        .ok_or(Error::InvalidSource)?;
    if !attachment_info.object_references.is_empty()
        || attachment_info
            .field_infos
            .iter()
            .any(|field| !field.object_references.is_empty())
    {
        return Err(Error::InvalidSource);
    }
    let attachment_payload = attachment.messages[0].data.as_slice();
    let (attachment_projection, attachment_report) =
        keynote_slide_number_codec::decode_slide_number_attachment_with_report(
            attachment_payload,
            decode_options(attachment_payload, limits, budget),
        )
        .map_err(|error| budget.map_decode_error(error))?;
    budget.charge_decode_report(attachment_report)?;
    if !matches!(attachment_projection.kind(), None | Some(0))
        || attachment_projection
            .string_equivalent()
            .is_some_and(|value| !value.is_empty())
    {
        return Err(Error::UnsupportedSource);
    }
    let attachment_scan_work = attachment_payload
        .len()
        .checked_mul(3)
        .ok_or(Error::InvalidSource)?;
    budget.charge_work(attachment_scan_work)?;
    reject_groups(attachment_payload, limits)?;
    let super_field =
        nested_unique_field(attachment_payload, &[1], limits)?.ok_or(Error::InvalidSource)?;
    if super_field.wire_type() != 2 {
        return Err(Error::InvalidSource);
    }
    let super_view =
        WireView::parse_with_limits(super_field.payload(), limits).map_err(map_wire_error)?;
    let mut attachment_kind = None;
    for field in super_view.fields() {
        field.validate_canonical_framing().map_err(map_wire_error)?;
        if matches!(field.wire_type(), 3 | 4) {
            return Err(Error::InvalidSource);
        }
        if field.number() == 2 && attachment_kind.replace(canonical_varint(field)?).is_some() {
            return Err(Error::InvalidSource);
        }
    }
    if !matches!(attachment_kind, None | Some(0))
        || reserved_identifiers.contains(&resolved_attachment_identifier)
        || storage_dependencies.contains_key(&resolved_attachment_identifier)
        || storage_dependencies
            .keys()
            .any(|identifier| reserved_identifiers.contains(identifier))
    {
        return Err(Error::InvalidSource);
    }
    let role_count = storage_dependencies
        .len()
        .checked_add(1)
        .ok_or(Error::InvalidSource)?;
    reserved_identifiers
        .try_reserve(role_count)
        .map_err(|_allocation| Error::Allocation { amount: role_count })?;
    reserved_identifiers.extend(storage_dependencies.into_keys());
    reserved_identifiers.insert(resolved_attachment_identifier);
    Ok(())
}
