//! Ordered CRUD for native value-axis chart reference lines.
//!
//! Pages, Numbers, and Keynote store the visible configuration in dedicated
//! `ReferenceLineNonStyleArchive` objects. The chart's reference-line
//! extension owns those objects through ordered references and stable UUIDs.

use std::collections::{HashMap, HashSet};

use litchi_iwa_common::chart::reference_line::{Kind, Line, Value};
use prost::Message;

use crate::archive::{ArchiveObject, RawMessage};
use crate::charts::IWorkChartArchive;
use crate::charts::source::CHART_MESSAGE_TYPE;
use crate::charts::unique_chart_object_archive_name;
use crate::package_metadata::{
    add_component_object_uuids, component_identifier_for_entry, component_uuid_identifiers,
    next_object_identifier, release_package_identifier_suffix, remove_component_object_uuids,
    set_package_last_object_identifier,
};
use crate::protobuf::{tsch, tsp, tss};
use crate::wire::{
    WireField, parse_wire_fields, patch_fixed64_field, patch_length_delimited_field,
    patch_varint_field,
};
use crate::{Error, IWorkPackage, Result};

pub(crate) const REFERENCE_LINE_STYLE_MESSAGE_TYPE: u32 = 5_030;
pub(crate) const REFERENCE_LINE_NON_STYLE_MESSAGE_TYPE: u32 = 5_031;
const GENERATED_REFERENCE_LINE_NON_STYLE_EXTENSION_FIELD: u32 = 10_000;
const TYPE_FIELD: u32 = 1;
const SHOW_LINE_FIELD: u32 = 2;
const SHOW_NAME_FIELD: u32 = 3;
const SHOW_VALUE_FIELD: u32 = 4;
const NAME_FIELD: u32 = 5;
const CUSTOM_VALUE_FIELD: u32 = 6;
const STANDARD_MESSAGE_VERSION: &[u32] = &[1, 0, 5];

const NATIVE_MINIMUM: i32 = 1;
const NATIVE_MAXIMUM: i32 = 2;
const NATIVE_AVERAGE: i32 = 3;
const NATIVE_MEDIAN: i32 = 4;
const NATIVE_CUSTOM: i32 = 5;
const MAX_REFERENCE_LINE_COUNT: usize = 5;

#[derive(Debug, Clone)]
struct ReferenceLineSlot {
    archive_name: String,
    object_id: u64,
    message_index: usize,
    uuid: tsp::Uuid,
    value: Line,
}

#[derive(Debug)]
struct ReferenceLineGraph {
    chart_message_index: usize,
    reference_lines: Option<tsch::ChartReferenceLinesArchive>,
    target_axis_index: Option<usize>,
    slots: Vec<ReferenceLineSlot>,
}

pub(crate) fn chart_reference_line_objects(
    chart: &IWorkChartArchive,
) -> Result<Vec<(u64, u32, &'static str)>> {
    let Some(reference_lines) = chart.reference_lines()? else {
        return Ok(Vec::new());
    };
    let mut objects = Vec::new();
    for axis in reference_lines.reference_line_non_styles_map {
        objects.extend(axis.reference_line_non_style_items.into_iter().map(|item| {
            (
                item.non_style.identifier,
                REFERENCE_LINE_NON_STYLE_MESSAGE_TYPE,
                "reference-line non-style",
            )
        }));
    }
    for axis in reference_lines.reference_line_styles_map {
        objects.extend(
            axis.reference_line_styles
                .into_iter()
                .flat_map(|styles| styles.entries)
                .map(|entry| {
                    (
                        entry.reference.identifier,
                        REFERENCE_LINE_STYLE_MESSAGE_TYPE,
                        "reference-line style",
                    )
                }),
        );
    }
    Ok(objects)
}

/// Read the ordered reference lines on a chart's primary value axis.
pub(crate) fn chart_reference_lines(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<Vec<Line>> {
    Ok(reference_line_graph(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?
    .slots
    .into_iter()
    .map(|slot| slot.value)
    .collect())
}

/// Replace the ordered reference lines on a chart's primary value axis.
pub(crate) fn set_chart_reference_lines(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    expected: &[Line],
) -> Result<()> {
    validate_reference_line_collection(expected)?;
    u32::try_from(expected.len())
        .map_err(|_| Error::InvalidFormat("chart reference-line count exceeds u32".to_owned()))?;

    let graph = reference_line_graph(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    let current = graph
        .slots
        .iter()
        .map(|slot| slot.value.clone())
        .collect::<Vec<_>>();
    if current == expected {
        return Ok(());
    }

    let retained_count = graph.slots.len().min(expected.len());
    for (slot, replacement) in graph.slots[..retained_count]
        .iter()
        .zip(&expected[..retained_count])
    {
        if &slot.value == replacement {
            continue;
        }
        ensure_reference_line_exclusive(
            package,
            slot.object_id,
            drawable_object_id,
            drawable_label,
        )?;
        let data = read_slot_data(package, slot)?;
        replace_slot_data(package, slot, patch_reference_line(&data, replacement)?)?;
    }

    let removed = graph.slots[retained_count..].to_vec();
    for slot in &removed {
        ensure_reference_line_exclusive(
            package,
            slot.object_id,
            drawable_object_id,
            drawable_label,
        )?;
    }
    remove_reference_line_objects(package, drawable_label, &removed)?;

    let mut additions = Vec::new();
    let mut next_identifier = next_object_identifier(package)?;
    let mut existing_uuids = reference_line_uuids(package)?;
    for replacement in &expected[retained_count..] {
        let identifier = next_identifier;
        next_identifier = next_identifier
            .checked_add(1)
            .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))?;
        let uuid = fresh_reference_line_uuid(&mut existing_uuids);
        additions.push((
            identifier,
            uuid,
            reference_line_object(identifier, replacement)?,
        ));
    }
    let created_ids = additions
        .iter()
        .map(|(identifier, _, _)| *identifier)
        .collect::<Vec<_>>();
    let created_items = additions
        .iter()
        .map(|(identifier, uuid, _)| (*identifier, *uuid))
        .collect::<Vec<_>>();
    if !additions.is_empty() {
        package.update_archive(chart_archive_name, |archive| {
            for (_, _, object) in additions.drain(..) {
                archive.insert_object(object)?;
            }
            Ok(())
        })?;
    }

    let retained = graph.slots[..retained_count]
        .iter()
        .map(|slot| (slot.object_id, slot.uuid))
        .chain(created_items)
        .collect::<Vec<_>>();
    patch_reference_line_graph(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        &graph,
        &retained,
    )?;
    update_component_registrations(package, chart_archive_name, &removed, &created_ids)?;

    if chart_reference_lines(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )? != expected
    {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} reference-line update failed validation"
        )));
    }
    Ok(())
}

fn reference_line_graph(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<ReferenceLineGraph> {
    let archive = package.archive(chart_archive_name)?;
    let object = archive.object(drawable_object_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} is missing"
        ))
    })?;
    let mut messages = object
        .messages
        .iter()
        .enumerate()
        .filter(|(_, message)| message.type_ == CHART_MESSAGE_TYPE);
    let Some((chart_message_index, message)) = messages.next() else {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} must have exactly one chart payload"
        )));
    };
    if messages.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} must have exactly one chart payload"
        )));
    }
    let chart = IWorkChartArchive::decode(message.data.as_slice())?;
    let reference_lines = chart.reference_lines()?;
    let target_axis_index = reference_lines.as_ref().and_then(|reference_lines| {
        reference_lines
            .reference_line_non_styles_map
            .iter()
            .position(|axis| is_primary_value_axis(axis.axis_id))
    });
    if reference_lines.as_ref().is_some_and(|reference_lines| {
        reference_lines
            .reference_line_non_styles_map
            .iter()
            .filter(|axis| is_primary_value_axis(axis.axis_id))
            .count()
            > 1
    }) {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} has duplicate primary value-axis reference-line maps"
        )));
    }

    let items = target_axis_index
        .and_then(|index| {
            reference_lines
                .as_ref()
                .map(|reference_lines| &reference_lines.reference_line_non_styles_map[index])
        })
        .map(|axis| axis.reference_line_non_style_items.as_slice())
        .unwrap_or_default();
    if items.len() > MAX_REFERENCE_LINE_COUNT {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} has more than {MAX_REFERENCE_LINE_COUNT} reference lines"
        )));
    }
    let mut object_ids = HashSet::with_capacity(items.len());
    let mut uuids = HashSet::with_capacity(items.len());
    let mut slots = Vec::with_capacity(items.len());
    for item in items {
        let identifier = item.non_style.identifier;
        let uuid_key = (item.uuid.lower, item.uuid.upper);
        if identifier == 0 || !object_ids.insert(identifier) || !uuids.insert(uuid_key) {
            return Err(Error::InvalidFormat(format!(
                "{drawable_label} chart {drawable_object_id} has an invalid or repeated reference line"
            )));
        }
        if object.archive_info.message_infos[chart_message_index]
            .object_references
            .iter()
            .filter(|candidate| **candidate == identifier)
            .count()
            != 1
        {
            return Err(Error::InvalidFormat(format!(
                "{drawable_label} chart {drawable_object_id} metadata does not reference line object {identifier} exactly once"
            )));
        }
        let archive_name = unique_chart_object_archive_name(
            package,
            identifier,
            "chart reference-line non-style object",
        )?;
        let line_archive = package.archive(&archive_name)?;
        let line_object = line_archive.object(identifier).ok_or_else(|| {
            Error::InvalidFormat(format!("chart reference line {identifier} is missing"))
        })?;
        let mut line_messages = line_object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == REFERENCE_LINE_NON_STYLE_MESSAGE_TYPE);
        let Some((message_index, line_message)) = line_messages.next() else {
            return Err(Error::InvalidFormat(format!(
                "chart reference line {identifier} must have exactly one non-style payload"
            )));
        };
        if line_messages.next().is_some() {
            return Err(Error::InvalidFormat(format!(
                "chart reference line {identifier} must have exactly one non-style payload"
            )));
        }
        slots.push(ReferenceLineSlot {
            archive_name,
            object_id: identifier,
            message_index,
            uuid: item.uuid,
            value: read_reference_line(&line_message.data)?,
        });
    }
    validate_reference_line_styles(
        reference_lines.as_ref(),
        drawable_object_id,
        drawable_label,
        items.len(),
    )?;
    validate_reference_line_collection(
        &slots
            .iter()
            .map(|slot| slot.value.clone())
            .collect::<Vec<_>>(),
    )?;
    Ok(ReferenceLineGraph {
        chart_message_index,
        reference_lines,
        target_axis_index,
        slots,
    })
}

fn validate_reference_line_collection(reference_lines: &[Line]) -> Result<()> {
    if reference_lines.len() > MAX_REFERENCE_LINE_COUNT {
        return Err(Error::InvalidFormat(format!(
            "iWork supports at most {MAX_REFERENCE_LINE_COUNT} reference lines per value axis"
        )));
    }
    let mut native_types = HashSet::with_capacity(reference_lines.len());
    for line in reference_lines {
        line.validate()?;
        let native_type = line.kind().native_value();
        if !native_types.insert(native_type) {
            return Err(Error::InvalidFormat(format!(
                "chart reference-line type {native_type} occurs more than once"
            )));
        }
    }
    Ok(())
}

fn is_primary_value_axis(axis_id: tsch::ChartAxisIdArchive) -> bool {
    axis_id.axis_type == Some(tsch::AxisType::Y as i32) && axis_id.ordinal.unwrap_or_default() == 0
}

fn validate_reference_line_styles(
    reference_lines: Option<&tsch::ChartReferenceLinesArchive>,
    drawable_object_id: u64,
    drawable_label: &str,
    line_count: usize,
) -> Result<()> {
    let Some(reference_lines) = reference_lines else {
        return Ok(());
    };
    let mut target = reference_lines
        .reference_line_styles_map
        .iter()
        .filter(|axis| is_primary_value_axis(axis.axis_id));
    let Some(axis) = target.next() else {
        return Ok(());
    };
    if target.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} has duplicate primary value-axis reference-line style maps"
        )));
    }
    if let Some(styles) = axis.reference_line_styles.as_ref() {
        if styles.entries.len() > MAX_REFERENCE_LINE_COUNT {
            return Err(Error::InvalidFormat(format!(
                "{drawable_label} chart {drawable_object_id} has more than {MAX_REFERENCE_LINE_COUNT} reference-line styles"
            )));
        }
        if usize::try_from(styles.count).ok() != Some(styles.entries.len()) {
            return Err(Error::InvalidFormat(format!(
                "{drawable_label} chart {drawable_object_id} reference-line style count is inconsistent"
            )));
        }
        let mut indexes = HashSet::with_capacity(styles.entries.len());
        for entry in &styles.entries {
            let index = usize::try_from(entry.index).map_err(|_| {
                Error::InvalidFormat("chart reference-line style index exceeds usize".to_owned())
            })?;
            if index >= line_count || entry.reference.identifier == 0 || !indexes.insert(index) {
                return Err(Error::InvalidFormat(format!(
                    "{drawable_label} chart {drawable_object_id} has an invalid reference-line style entry"
                )));
            }
        }
    }
    Ok(())
}

fn read_slot_data(package: &IWorkPackage, slot: &ReferenceLineSlot) -> Result<Vec<u8>> {
    let archive = package.archive(&slot.archive_name)?;
    let object = archive.object(slot.object_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "chart reference line {} is missing",
            slot.object_id
        ))
    })?;
    let message = object.messages.get(slot.message_index).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "chart reference line {} message index changed",
            slot.object_id
        ))
    })?;
    if message.type_ != REFERENCE_LINE_NON_STYLE_MESSAGE_TYPE {
        return Err(Error::InvalidFormat(format!(
            "chart reference line {} message type changed",
            slot.object_id
        )));
    }
    Ok(message.data.clone())
}

fn replace_slot_data(
    package: &mut IWorkPackage,
    slot: &ReferenceLineSlot,
    data: Vec<u8>,
) -> Result<()> {
    package.update_archive(&slot.archive_name, |archive| {
        let object = archive.object_mut(slot.object_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "chart reference line {} is missing",
                slot.object_id
            ))
        })?;
        object.replace_message(
            slot.message_index,
            RawMessage {
                type_: REFERENCE_LINE_NON_STYLE_MESSAGE_TYPE,
                data,
            },
        )?;
        Ok(())
    })
}

fn remove_reference_line_objects(
    package: &mut IWorkPackage,
    drawable_label: &str,
    removed: &[ReferenceLineSlot],
) -> Result<()> {
    let mut by_archive = HashMap::<String, Vec<u64>>::new();
    for slot in removed {
        by_archive
            .entry(slot.archive_name.clone())
            .or_default()
            .push(slot.object_id);
    }
    for (archive_name, identifiers) in by_archive {
        package.update_archive(&archive_name, |archive| {
            for identifier in &identifiers {
                archive.remove_object(*identifier).ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "{drawable_label} chart reference line {identifier} is missing"
                    ))
                })?;
            }
            Ok(())
        })?;
    }
    Ok(())
}

fn patch_reference_line_graph(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    graph: &ReferenceLineGraph,
    identifiers: &[(u64, tsp::Uuid)],
) -> Result<()> {
    let mut next = graph.reference_lines.clone().unwrap_or_default();
    let items = identifiers
        .iter()
        .map(|(identifier, uuid)| tsch::ChartReferenceLineNonStyleItem {
            non_style: tsp::Reference {
                identifier: *identifier,
                ..Default::default()
            },
            uuid: *uuid,
        })
        .collect::<Vec<_>>();
    match (graph.target_axis_index, items.is_empty()) {
        (Some(index), false) => {
            next.reference_line_non_styles_map[index].reference_line_non_style_items = items;
        },
        (Some(index), true) => {
            next.reference_line_non_styles_map.remove(index);
        },
        (None, false) => {
            next.reference_line_non_styles_map
                .push(tsch::ChartAxisReferenceLineNonStylesArchive {
                    axis_id: primary_value_axis_id(),
                    reference_line_non_style_items: items,
                });
        },
        (None, true) => {},
    }
    trim_primary_value_axis_styles(&mut next, identifiers.len())?;

    let previous_references = managed_reference_line_references(graph.reference_lines.as_ref());
    let next_references = managed_reference_line_references(Some(&next));
    let extension = (!next.reference_line_non_styles_map.is_empty()
        || !next.reference_line_styles_map.is_empty()
        || next.theme_preset_reference_line_style.is_some())
    .then_some(next);
    package.update_archive(chart_archive_name, |archive| {
        let object = archive.object_mut(drawable_object_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "{drawable_label} chart {drawable_object_id} is missing"
            ))
        })?;
        let message = object
            .messages
            .get(graph.chart_message_index)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "{drawable_label} chart {drawable_object_id} message index changed"
                ))
            })?;
        if message.type_ != CHART_MESSAGE_TYPE {
            return Err(Error::InvalidFormat(format!(
                "{drawable_label} chart {drawable_object_id} message type changed"
            )));
        }
        let mut chart = IWorkChartArchive::decode(&message.data)?;
        chart.set_reference_lines(extension.as_ref())?;
        object.replace_message(
            graph.chart_message_index,
            RawMessage {
                type_: CHART_MESSAGE_TYPE,
                data: chart.encode()?,
            },
        )?;
        let metadata =
            &mut object.archive_info.message_infos[graph.chart_message_index].object_references;
        metadata.retain(|identifier| !previous_references.contains(identifier));
        for identifier in next_references {
            if !metadata.contains(&identifier) {
                metadata.push(identifier);
            }
        }
        Ok(())
    })
}

fn trim_primary_value_axis_styles(
    reference_lines: &mut tsch::ChartReferenceLinesArchive,
    line_count: usize,
) -> Result<()> {
    let mut remove_axis = None;
    for (index, axis) in reference_lines
        .reference_line_styles_map
        .iter_mut()
        .enumerate()
    {
        if !is_primary_value_axis(axis.axis_id) {
            continue;
        }
        if let Some(styles) = axis.reference_line_styles.as_mut() {
            styles
                .entries
                .retain(|entry| usize::try_from(entry.index).is_ok_and(|index| index < line_count));
            styles.count = u32::try_from(styles.entries.len()).map_err(|_| {
                Error::InvalidFormat("chart reference-line style count exceeds u32".to_owned())
            })?;
            if styles.entries.is_empty() {
                axis.reference_line_styles = None;
            }
        }
        if axis.reference_line_styles.is_none() {
            remove_axis = Some(index);
        }
    }
    if let Some(index) = remove_axis {
        reference_lines.reference_line_styles_map.remove(index);
    }
    Ok(())
}

fn managed_reference_line_references(
    reference_lines: Option<&tsch::ChartReferenceLinesArchive>,
) -> HashSet<u64> {
    let mut identifiers = HashSet::new();
    let Some(reference_lines) = reference_lines else {
        return identifiers;
    };
    for axis in &reference_lines.reference_line_non_styles_map {
        for item in &axis.reference_line_non_style_items {
            if item.non_style.identifier != 0 {
                identifiers.insert(item.non_style.identifier);
            }
        }
    }
    for axis in &reference_lines.reference_line_styles_map {
        for entry in axis
            .reference_line_styles
            .as_ref()
            .into_iter()
            .flat_map(|styles| &styles.entries)
        {
            if entry.reference.identifier != 0 {
                identifiers.insert(entry.reference.identifier);
            }
        }
    }
    identifiers
}

fn primary_value_axis_id() -> tsch::ChartAxisIdArchive {
    tsch::ChartAxisIdArchive {
        axis_type: Some(tsch::AxisType::Y as i32),
        ordinal: Some(0),
    }
}

fn ensure_reference_line_exclusive(
    package: &IWorkPackage,
    reference_line_id: u64,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<()> {
    let mut owners = 0usize;
    for archive_name in package.iwa_entry_names() {
        for object in &package.archive(archive_name)?.objects {
            for message in object
                .messages
                .iter()
                .filter(|message| message.type_ == CHART_MESSAGE_TYPE)
            {
                let chart = IWorkChartArchive::decode(&message.data)?;
                if chart.reference_lines()?.is_some_and(|reference_lines| {
                    reference_lines
                        .reference_line_non_styles_map
                        .iter()
                        .flat_map(|axis| &axis.reference_line_non_style_items)
                        .any(|item| item.non_style.identifier == reference_line_id)
                }) {
                    owners = owners.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat("chart reference-line owner count overflow".to_owned())
                    })?;
                }
            }
        }
    }
    if owners != 1 {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} reference line {reference_line_id} is shared by {owners} charts"
        )));
    }
    Ok(())
}

fn update_component_registrations(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    removed: &[ReferenceLineSlot],
    created_ids: &[u64],
) -> Result<()> {
    let mut removed_by_component = HashMap::<u64, Vec<u64>>::new();
    for slot in removed {
        if let Some(component_id) = component_identifier_for_entry(package, &slot.archive_name)? {
            let registered = component_uuid_identifiers(package, component_id)?.unwrap_or_default();
            if registered.contains(&slot.object_id) {
                removed_by_component
                    .entry(component_id)
                    .or_default()
                    .push(slot.object_id);
            }
        }
    }
    for (component_id, identifiers) in removed_by_component {
        remove_component_object_uuids(package, component_id, &identifiers)?;
    }
    if !created_ids.is_empty() {
        if let Some(component_id) = component_identifier_for_entry(package, chart_archive_name)? {
            add_component_object_uuids(package, component_id, created_ids)?;
        }
        set_package_last_object_identifier(
            package,
            *created_ids.last().ok_or_else(|| {
                Error::InvalidFormat("reference-line creation lost identifiers".to_owned())
            })?,
        )?;
    }
    if !removed.is_empty() {
        release_package_identifier_suffix(
            package,
            &removed
                .iter()
                .map(|slot| slot.object_id)
                .collect::<Vec<_>>(),
        )?;
    }
    Ok(())
}

fn reference_line_uuids(package: &IWorkPackage) -> Result<HashSet<(u64, u64)>> {
    let mut uuids = HashSet::new();
    for archive_name in package.iwa_entry_names() {
        for object in &package.archive(archive_name)?.objects {
            for message in object
                .messages
                .iter()
                .filter(|message| message.type_ == CHART_MESSAGE_TYPE)
            {
                if let Some(reference_lines) =
                    IWorkChartArchive::decode(&message.data)?.reference_lines()?
                {
                    for item in reference_lines
                        .reference_line_non_styles_map
                        .iter()
                        .flat_map(|axis| &axis.reference_line_non_style_items)
                    {
                        if !uuids.insert((item.uuid.lower, item.uuid.upper)) {
                            return Err(Error::InvalidFormat(
                                "chart reference-line UUID is duplicated".to_owned(),
                            ));
                        }
                    }
                }
            }
        }
    }
    Ok(uuids)
}

fn fresh_reference_line_uuid(existing: &mut HashSet<(u64, u64)>) -> tsp::Uuid {
    loop {
        let bytes = litchi_core::id::generate_guid_bytes();
        let mut lower = [0; 8];
        lower.copy_from_slice(&bytes[..8]);
        let mut upper = [0; 8];
        upper.copy_from_slice(&bytes[8..]);
        let uuid = tsp::Uuid {
            lower: u64::from_le_bytes(lower),
            upper: u64::from_le_bytes(upper),
        };
        if existing.insert((uuid.lower, uuid.upper)) {
            return uuid;
        }
    }
}

fn reference_line_object(identifier: u64, reference_line: &Line) -> Result<ArchiveObject> {
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: REFERENCE_LINE_NON_STYLE_MESSAGE_TYPE,
            data: canonical_reference_line(reference_line)?,
        }],
    )?;
    object.archive_info.message_infos[0].versions = STANDARD_MESSAGE_VERSION.to_vec();
    Ok(object)
}

fn canonical_reference_line(reference_line: &Line) -> Result<Vec<u8>> {
    reference_line.validate()?;
    let mut data = tsch::ReferenceLineNonStyleArchive {
        super_: Some(tss::StyleArchive::default()),
    }
    .encode_to_vec();
    let kind = reference_line.kind();
    let generated = tsch::generated::ReferenceLineNonStyleArchive {
        tschreferencelinedefaulttype: Some(kind.native_value()),
        tschreferencelinedefaultshowline: Some(true),
        tschreferencelinedefaultshowlabel: Some(reference_line.shows_name()),
        tschreferencelinedefaultshowvaluelabel: Some(reference_line.shows_value()),
        tschreferencelinedefaultlabel: Some(reference_line.name().to_owned()),
        tschreferencelinedefaultcustomvalue: kind.custom_value().map(|value| {
            tsch::ChartsNsNumberDoubleArchive {
                number_archive: Some(value.value()),
            }
        }),
    };
    crate::wire::append_length_delimited_field(
        &mut data,
        GENERATED_REFERENCE_LINE_NON_STYLE_EXTENSION_FIELD,
        &generated.encode_to_vec(),
    )?;
    if read_reference_line(&data)? != *reference_line {
        return Err(Error::InvalidFormat(
            "canonical chart reference-line encoding failed validation".to_owned(),
        ));
    }
    Ok(data)
}

fn read_reference_line(data: &[u8]) -> Result<Line> {
    tsch::ReferenceLineNonStyleArchive::decode(data)?;
    let extension = generated_reference_line_extension(data)?.ok_or_else(|| {
        Error::InvalidFormat("chart reference-line non-style has no generated payload".to_owned())
    })?;
    tsch::generated::ReferenceLineNonStyleArchive::decode(extension)?;
    let native_type = strict_optional_varint(extension, TYPE_FIELD)?
        .map(raw_i32)
        .transpose()?
        .unwrap_or_default();
    let custom_value = strict_optional_message(extension, CUSTOM_VALUE_FIELD)?
        .map(read_custom_value)
        .transpose()?;
    let kind = match native_type {
        NATIVE_MINIMUM => known_kind_without_custom(Kind::Minimum, custom_value)?,
        NATIVE_MAXIMUM => known_kind_without_custom(Kind::Maximum, custom_value)?,
        NATIVE_AVERAGE => known_kind_without_custom(Kind::Average, custom_value)?,
        NATIVE_MEDIAN => known_kind_without_custom(Kind::Median, custom_value)?,
        NATIVE_CUSTOM => Kind::Custom(custom_value.ok_or_else(|| {
            Error::InvalidFormat("custom chart reference line has no finite value".to_owned())
        })?),
        native_type => Kind::unsupported(native_type, custom_value)?,
    };
    if !strict_optional_bool(extension, SHOW_LINE_FIELD)?.unwrap_or(true) {
        return Err(Error::InvalidFormat(
            "stored chart reference line is hidden instead of removed".to_owned(),
        ));
    }
    let show_name = strict_optional_bool(extension, SHOW_NAME_FIELD)?.unwrap_or(true);
    let show_value = strict_optional_bool(extension, SHOW_VALUE_FIELD)?
        .unwrap_or_else(|| kind.default_shows_value());
    let mut value = Line::from_kind(kind)?
        .with_name_visibility(show_name)
        .with_value_visibility(show_value);
    if let Some(bytes) = strict_optional_message(extension, NAME_FIELD)? {
        let name =
            std::str::from_utf8(bytes).map_err(|error| Error::InvalidFormat(error.to_string()))?;
        value = value.try_with_name(name)?;
    }
    value.validate()?;
    Ok(value)
}

fn known_kind_without_custom(kind: Kind, custom_value: Option<Value>) -> Result<Kind> {
    if custom_value.is_some() {
        return Err(Error::InvalidFormat(format!(
            "chart reference-line type {} unexpectedly stores a custom value",
            kind.native_value()
        )));
    }
    Ok(kind)
}

fn patch_reference_line(data: &[u8], reference_line: &Line) -> Result<Vec<u8>> {
    reference_line.validate()?;
    tsch::ReferenceLineNonStyleArchive::decode(data)?;
    let existing = generated_reference_line_extension(data)?.ok_or_else(|| {
        Error::InvalidFormat("chart reference-line non-style has no generated payload".to_owned())
    })?;
    let extension = existing;
    tsch::generated::ReferenceLineNonStyleArchive::decode(extension)?;
    let _ = read_reference_line(data)?;

    let kind = reference_line.kind();
    let mut patched = patch_varint(
        extension,
        TYPE_FIELD,
        Some(kind.native_value() as i64 as u64),
    )?;
    patched = patch_varint(&patched, SHOW_LINE_FIELD, Some(1))?;
    patched = patch_varint(
        &patched,
        SHOW_NAME_FIELD,
        Some(u64::from(reference_line.shows_name())),
    )?;
    patched = patch_varint(
        &patched,
        SHOW_VALUE_FIELD,
        Some(u64::from(reference_line.shows_value())),
    )?;
    patched = patch_message(&patched, NAME_FIELD, Some(reference_line.name().as_bytes()))?;
    let custom_value = kind
        .custom_value()
        .map(|value| {
            let existing_custom = strict_optional_message(extension, CUSTOM_VALUE_FIELD)?;
            if let Some(existing_custom) = existing_custom {
                return patch_fixed64_field(
                    existing_custom,
                    1,
                    strict_optional_fixed64(existing_custom, 1)?.is_some(),
                    Some(value.value().to_bits()),
                );
            }
            Ok(tsch::ChartsNsNumberDoubleArchive {
                number_archive: Some(value.value()),
            }
            .encode_to_vec())
        })
        .transpose()?;
    patched = patch_message(&patched, CUSTOM_VALUE_FIELD, custom_value.as_deref())?;
    let output = patch_length_delimited_field(
        data,
        GENERATED_REFERENCE_LINE_NON_STYLE_EXTENSION_FIELD,
        true,
        Some(&patched),
    )?;
    if read_reference_line(&output)? != *reference_line {
        return Err(Error::InvalidFormat(
            "chart reference-line wire patch failed validation".to_owned(),
        ));
    }
    Ok(output)
}

fn generated_reference_line_extension(data: &[u8]) -> Result<Option<&[u8]>> {
    strict_optional_message(data, GENERATED_REFERENCE_LINE_NON_STYLE_EXTENSION_FIELD)
}

fn patch_varint(data: &[u8], field_number: u32, replacement: Option<u64>) -> Result<Vec<u8>> {
    patch_varint_field(
        data,
        field_number,
        strict_optional_varint(data, field_number)?.is_some(),
        replacement,
    )
}

fn patch_message(data: &[u8], field_number: u32, replacement: Option<&[u8]>) -> Result<Vec<u8>> {
    patch_length_delimited_field(
        data,
        field_number,
        strict_optional_message(data, field_number)?.is_some(),
        replacement,
    )
}

fn strict_optional_bool(data: &[u8], field_number: u32) -> Result<Option<bool>> {
    strict_optional_varint(data, field_number)?
        .map(|value| match value {
            0 => Ok(false),
            1 => Ok(true),
            value => Err(Error::InvalidFormat(format!(
                "chart reference-line bool field {field_number} must be 0 or 1, found {value}"
            ))),
        })
        .transpose()
}

fn strict_optional_varint(data: &[u8], field_number: u32) -> Result<Option<u64>> {
    let fields = parse_wire_fields(data)?;
    let mut matches = fields.iter().filter(|field| field.number() == field_number);
    let Some(field) = matches.next() else {
        return Ok(None);
    };
    if matches.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "singular chart reference-line field {field_number} occurs more than once"
        )));
    }
    if field.wire_type() != 0 {
        return Err(Error::InvalidFormat(format!(
            "chart reference-line field {field_number} is not a varint"
        )));
    }
    validate_canonical_field(data, field, false)?;
    let payload = field.checked_payload(data)?;
    let (value, consumed) =
        litchi_iwa_common::varint::decode_varint_from_bytes(payload).map_err(|error| {
            Error::InvalidFormat(format!(
                "chart reference-line field {field_number} is invalid: {error}"
            ))
        })?;
    if consumed != payload.len() || litchi_iwa_common::varint::encoded_len(value) != consumed {
        return Err(Error::InvalidFormat(format!(
            "chart reference-line field {field_number} is not canonically encoded"
        )));
    }
    Ok(Some(value))
}

fn strict_optional_message(data: &[u8], field_number: u32) -> Result<Option<&[u8]>> {
    let fields = parse_wire_fields(data)?;
    let mut matches = fields.iter().filter(|field| field.number() == field_number);
    let Some(field) = matches.next() else {
        return Ok(None);
    };
    if matches.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "singular chart reference-line field {field_number} occurs more than once"
        )));
    }
    if field.wire_type() != 2 {
        return Err(Error::InvalidFormat(format!(
            "chart reference-line field {field_number} is not length-delimited"
        )));
    }
    validate_canonical_field(data, field, true)?;
    Ok(Some(field.checked_payload(data)?))
}

fn read_custom_value(data: &[u8]) -> Result<Value> {
    let value = strict_optional_fixed64(data, 1)?.ok_or_else(|| {
        Error::InvalidFormat("chart reference-line custom value is missing its number".to_owned())
    })?;
    Value::new(f64::from_le_bytes(value.to_le_bytes()))
        .map_err(|error| Error::InvalidFormat(error.to_string()))
}

fn strict_optional_fixed64(data: &[u8], field_number: u32) -> Result<Option<u64>> {
    let fields = parse_wire_fields(data)?;
    let mut matches = fields.iter().filter(|field| field.number() == field_number);
    let Some(field) = matches.next() else {
        return Ok(None);
    };
    if matches.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "singular chart reference-line field {field_number} occurs more than once"
        )));
    }
    if field.wire_type() != 1 {
        return Err(Error::InvalidFormat(format!(
            "chart reference-line field {field_number} is not fixed64"
        )));
    }
    validate_canonical_field(data, field, false)?;
    let payload = field.checked_payload(data)?;
    let bytes: [u8; 8] = payload.try_into().map_err(|_| {
        Error::InvalidFormat(format!(
            "chart reference-line field {field_number} has an invalid fixed64 payload"
        ))
    })?;
    Ok(Some(u64::from_le_bytes(bytes)))
}

fn validate_canonical_field(data: &[u8], field: &WireField, length_delimited: bool) -> Result<()> {
    field.validate_canonical_key(data)?;
    if length_delimited {
        field.validate_canonical_length(data)?;
    }
    Ok(())
}

fn raw_i32(value: u64) -> Result<i32> {
    let signed = value as i64;
    i32::try_from(signed).map_err(|_| {
        Error::InvalidFormat(format!(
            "chart reference-line type varint {value} is not an i32"
        ))
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn custom_values_must_be_finite() {
        assert!(Value::new(17.5).is_ok());
        assert!(Value::new(f64::NAN).is_err());
        assert!(Value::new(f64::INFINITY).is_err());
    }

    #[test]
    fn canonical_reference_lines_round_trip() {
        let lines = [
            Line::minimum(),
            Line::maximum().with_name_visibility(false),
            Line::average().with_value_visibility(true),
            Line::median().try_with_name("Middle").unwrap(),
            Line::custom(Value::new(17.5).unwrap())
                .try_with_name("Threshold")
                .unwrap(),
        ];
        for line in lines {
            let encoded = canonical_reference_line(&line).unwrap();
            assert_eq!(read_reference_line(&encoded).unwrap(), line);
        }
    }

    #[test]
    fn malformed_reference_line_fields_are_rejected() {
        let line = Line::average();
        let encoded = canonical_reference_line(&line).unwrap();
        let extension = generated_reference_line_extension(&encoded)
            .unwrap()
            .unwrap();
        let mut duplicate = extension.to_vec();
        crate::wire::append_varint_field(&mut duplicate, TYPE_FIELD, NATIVE_MEDIAN as u64).unwrap();
        let malformed = patch_length_delimited_field(
            &encoded,
            GENERATED_REFERENCE_LINE_NON_STYLE_EXTENSION_FIELD,
            true,
            Some(&duplicate),
        )
        .unwrap();
        assert!(read_reference_line(&malformed).is_err());
    }

    #[test]
    fn recognized_reference_line_fields_require_canonical_framing() {
        let encoded = canonical_reference_line(&Line::average()).unwrap();

        let noncanonical_key = [0x88, 0x00, NATIVE_AVERAGE as u8];
        let malformed = patch_length_delimited_field(
            &encoded,
            GENERATED_REFERENCE_LINE_NON_STYLE_EXTENSION_FIELD,
            true,
            Some(&noncanonical_key),
        )
        .unwrap();
        assert!(read_reference_line(&malformed).is_err());

        let mut noncanonical_length = Vec::new();
        crate::wire::append_varint_field(
            &mut noncanonical_length,
            TYPE_FIELD,
            NATIVE_AVERAGE as u64,
        )
        .unwrap();
        noncanonical_length.extend([0x2a, 0x87, 0x00]);
        noncanonical_length.extend_from_slice(b"Average");
        let malformed = patch_length_delimited_field(
            &encoded,
            GENERATED_REFERENCE_LINE_NON_STYLE_EXTENSION_FIELD,
            true,
            Some(&noncanonical_length),
        )
        .unwrap();
        assert!(read_reference_line(&malformed).is_err());
    }

    #[test]
    fn unknown_reference_line_fields_remain_permissive_and_preserved() {
        let original = canonical_reference_line(&Line::average()).unwrap();
        let extension = generated_reference_line_extension(&original)
            .unwrap()
            .unwrap();
        // Field 99 uses an intentionally overlong key varint. Unknown fields
        // remain opaque and permissive even while recognized fields are strict.
        let unknown = [0x98, 0x86, 0x00, 7];
        let mut extension_with_unknown = extension.to_vec();
        extension_with_unknown.extend(unknown);
        let input = patch_length_delimited_field(
            &original,
            GENERATED_REFERENCE_LINE_NON_STYLE_EXTENSION_FIELD,
            true,
            Some(&extension_with_unknown),
        )
        .unwrap();

        assert_eq!(read_reference_line(&input).unwrap(), Line::average());
        let replacement = Line::average().with_name_visibility(false);
        let output = patch_reference_line(&input, &replacement).unwrap();
        let output_extension = generated_reference_line_extension(&output)
            .unwrap()
            .unwrap();
        assert!(output_extension.ends_with(&unknown));
        assert_eq!(read_reference_line(&output).unwrap(), replacement);
    }

    #[test]
    fn custom_value_patching_retains_nested_unknown_fields() {
        let original = canonical_reference_line(&Line::custom(Value::new(17.5).unwrap())).unwrap();
        let extension = generated_reference_line_extension(&original)
            .unwrap()
            .unwrap();
        let custom = strict_optional_message(extension, CUSTOM_VALUE_FIELD)
            .unwrap()
            .unwrap();
        let mut custom_with_unknown = custom.to_vec();
        crate::wire::append_varint_field(&mut custom_with_unknown, 99, 7).unwrap();
        let extension_with_unknown = patch_length_delimited_field(
            extension,
            CUSTOM_VALUE_FIELD,
            true,
            Some(&custom_with_unknown),
        )
        .unwrap();
        let input = patch_length_delimited_field(
            &original,
            GENERATED_REFERENCE_LINE_NON_STYLE_EXTENSION_FIELD,
            true,
            Some(&extension_with_unknown),
        )
        .unwrap();
        let replacement = Line::custom(Value::new(23.0).unwrap());
        let output = patch_reference_line(&input, &replacement).unwrap();
        let output_extension = generated_reference_line_extension(&output)
            .unwrap()
            .unwrap();
        let output_custom = strict_optional_message(output_extension, CUSTOM_VALUE_FIELD)
            .unwrap()
            .unwrap();
        assert_eq!(strict_optional_varint(output_custom, 99).unwrap(), Some(7));
        assert_eq!(read_reference_line(&output).unwrap(), replacement);
    }

    #[test]
    fn native_reference_line_kinds_are_unique_and_bounded() {
        assert!(
            validate_reference_line_collection(&[
                Line::minimum(),
                Line::maximum(),
                Line::average(),
                Line::median(),
                Line::custom(Value::new(1.0).unwrap()),
            ])
            .is_ok()
        );
        assert!(validate_reference_line_collection(&[Line::average(), Line::average(),]).is_err());
    }
}
