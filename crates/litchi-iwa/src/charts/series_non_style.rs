//! Shared lossless access to sparse chart-series non-style objects.
//!
//! Per-series behavioral overrides are stored in private
//! `TSCH.ChartSeriesNonStyleArchive` objects referenced by a sparse array on
//! the chart. This module owns graph validation, object allocation/removal,
//! component UUID bookkeeping, and sparse-reference rewrites so focused chart
//! features only need to encode and decode their own property.

use std::collections::{HashMap, HashSet};

use prost::Message;

use crate::archive::{ArchiveObject, RawMessage};
use crate::charts::IWorkChartArchive;
use crate::charts::object_container::{
    ObjectContainerAllocation, insert_object_container, is_object_container_archive,
    remove_object_container_objects, reserve_object_container,
};
use crate::charts::source::{
    CHART_MESSAGE_TYPE, SERIES_NON_STYLE_MESSAGE_TYPE, register_chart_styles,
    unregister_chart_styles,
};
use crate::charts::unique_chart_object_archive_name;
use crate::package_metadata::{
    add_component_object_uuids, component_identifier_for_entry, component_uuid_identifiers,
    next_object_identifier, release_package_identifier_suffix, remove_component_object_uuids,
    set_package_last_object_identifier,
};
use crate::protobuf::{tsch, tsp, tss};
use crate::wire::{append_varint_field, parse_wire_fields, patch_length_delimited_field};
use crate::{Error, IWorkPackage, Result};

/// Proto2 extension holding generated chart-series non-style properties.
pub(crate) const GENERATED_CHART_SERIES_NON_STYLE_EXTENSION_FIELD: u32 = 10_000;

const SUPPORTS_CUSTOM_NUMBER_FORMAT_FIELD: u32 = 10_001;
const SUPPORTS_CUSTOM_DATE_FORMAT_FIELD: u32 = 10_002;
const SUPPORTS_CALLOUT_LINES_FIELD: u32 = 10_003;
const STANDARD_MESSAGE_VERSION: &[u32] = &[1, 0, 5];

/// Native parent used when allocating a previously absent series non-style.
///
/// Most label and formatter overrides are document styles. Pure geometry
/// overrides, however, use an empty style parent in files saved by iWork.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum NewChartSeriesNonStyleBase {
    Styled,
    Unstyled,
}

#[derive(Debug, Clone)]
struct ChartSeriesNonStyleSlot {
    archive_name: String,
    object_id: u64,
    message_index: usize,
}

#[derive(Debug)]
struct ChartSeriesNonStyleGraph {
    chart_message_index: usize,
    slots: Vec<Option<ChartSeriesNonStyleSlot>>,
}

/// Read one typed value for each series in chart series order.
pub(crate) fn chart_series_non_style_values<T>(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    series_count: usize,
    default: T,
    read: impl Fn(&[u8]) -> Result<T>,
) -> Result<Vec<T>>
where
    T: Clone,
{
    let graph = chart_series_non_style_graph(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        series_count,
    )?;
    read_values_from_graph(package, &graph, &default, &read)
}

/// Set one typed value for each series in chart series order.
pub(crate) fn set_chart_series_non_style_values<T>(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    property_label: &str,
    new_object_base: NewChartSeriesNonStyleBase,
    expected: &[T],
    default: T,
    read: impl Fn(&[u8]) -> Result<T> + Copy,
    patch: impl Fn(&[u8], &T) -> Result<Vec<u8>>,
) -> Result<()>
where
    T: Clone + PartialEq,
{
    if expected.is_empty() {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} has no series for {property_label}"
        )));
    }
    let graph = chart_series_non_style_graph(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        expected.len(),
    )?;
    let stylesheet_id = chart_stylesheet_id(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    let current = read_values_from_graph(package, &graph, &default, &read)?;
    if current == expected {
        return Ok(());
    }

    let styled_empty = canonical_empty_chart_series_non_style_data_with_stylesheet(stylesheet_id)?;
    let unstyled_empty = canonical_empty_chart_series_non_style_data()?;
    let unstyled_new_base = chart_series_non_style_data_with_style(tss::StyleArchive::default());
    let canonical_empty = match new_object_base {
        NewChartSeriesNonStyleBase::Styled => styled_empty.as_slice(),
        NewChartSeriesNonStyleBase::Unstyled => unstyled_new_base.as_slice(),
    };
    let mut next_identifier = next_object_identifier(package)?;
    let mut final_ids = graph
        .slots
        .iter()
        .map(|slot| slot.as_ref().map(|slot| slot.object_id))
        .collect::<Vec<_>>();
    let mut updates = Vec::new();
    let mut removals = Vec::new();
    let mut styled_creations = Vec::new();
    let mut chart_creation_ids = Vec::new();
    let mut unstyled_creations = Vec::new();
    let mut object_container = None;
    let mut created_ids = Vec::new();
    let mut registrations_by_archive = HashMap::<String, Vec<u64>>::new();
    let mut unregistrations_by_archive = HashMap::<String, Vec<u64>>::new();
    for (index, ((slot, current), replacement)) in
        graph.slots.iter().zip(&current).zip(expected).enumerate()
    {
        if current == replacement {
            continue;
        }
        if let Some(slot) = slot {
            let (mut patched, registered_stylesheet_id) = slot.read(package, |data| {
                Ok((
                    patch(data, replacement)?,
                    chart_series_non_style_stylesheet_id(data)?,
                ))
            })?;
            if let Some(registered_stylesheet_id) = registered_stylesheet_id
                && registered_stylesheet_id != stylesheet_id
            {
                return Err(Error::InvalidFormat(format!(
                    "{drawable_label} chart series non-style {} belongs to stylesheet {registered_stylesheet_id}, expected {stylesheet_id}",
                    slot.object_id
                )));
            }
            let needs_stylesheet_registration = new_object_base
                == NewChartSeriesNonStyleBase::Styled
                && replacement != &default
                && registered_stylesheet_id.is_none();
            if needs_stylesheet_registration {
                patched = set_chart_series_non_style_stylesheet(&patched, Some(stylesheet_id))?;
            }
            let is_empty = patched.as_slice() == styled_empty.as_slice()
                || patched.as_slice() == unstyled_empty.as_slice()
                || patched.as_slice() == unstyled_new_base.as_slice();
            let owner_count = slot.owner_count(package)?;
            if owner_count == 0 {
                return Err(Error::InvalidFormat(format!(
                    "{drawable_label} chart {drawable_object_id} series non-style {} has no owner",
                    slot.object_id
                )));
            }
            if owner_count > 1 {
                if is_empty {
                    final_ids[index] = None;
                    continue;
                }
                let allocation_base = if registered_stylesheet_id.is_some() {
                    NewChartSeriesNonStyleBase::Styled
                } else {
                    new_object_base
                };
                let identifier = allocate_series_non_style_identifier(
                    package,
                    &mut next_identifier,
                    allocation_base,
                    &mut object_container,
                )?;
                final_ids[index] = Some(identifier);
                let object = chart_series_non_style_object(identifier, patched)?;
                match allocation_base {
                    NewChartSeriesNonStyleBase::Styled => {
                        registrations_by_archive
                            .entry(chart_archive_name.to_owned())
                            .or_default()
                            .push(identifier);
                        chart_creation_ids.push(identifier);
                        styled_creations.push(object);
                    },
                    NewChartSeriesNonStyleBase::Unstyled => unstyled_creations.push(object),
                }
                created_ids.push(identifier);
            } else if is_empty {
                final_ids[index] = None;
                if registered_stylesheet_id.is_some() {
                    unregistrations_by_archive
                        .entry(slot.archive_name.clone())
                        .or_default()
                        .push(slot.object_id);
                }
                removals.push(slot.clone());
            } else {
                if needs_stylesheet_registration {
                    registrations_by_archive
                        .entry(slot.archive_name.clone())
                        .or_default()
                        .push(slot.object_id);
                }
                updates.push((slot.clone(), patched));
            }
        } else if replacement != &default {
            let mut data = patch(canonical_empty, replacement)?;
            if data.as_slice() == canonical_empty {
                return Err(Error::InvalidFormat(format!(
                    "{drawable_label} chart {drawable_object_id} {property_label} patch produced no native override"
                )));
            }
            if new_object_base == NewChartSeriesNonStyleBase::Unstyled {
                append_chart_series_non_style_capabilities(&mut data)?;
            }
            let identifier = allocate_series_non_style_identifier(
                package,
                &mut next_identifier,
                new_object_base,
                &mut object_container,
            )?;
            final_ids[index] = Some(identifier);
            if new_object_base == NewChartSeriesNonStyleBase::Styled {
                registrations_by_archive
                    .entry(chart_archive_name.to_owned())
                    .or_default()
                    .push(identifier);
                chart_creation_ids.push(identifier);
            }
            let object = chart_series_non_style_object(identifier, data)?;
            match new_object_base {
                NewChartSeriesNonStyleBase::Styled => styled_creations.push(object),
                NewChartSeriesNonStyleBase::Unstyled => unstyled_creations.push(object),
            }
            created_ids.push(identifier);
        }
    }

    for (slot, data) in updates {
        slot.replace(package, data)?;
    }
    for (archive_name, identifiers) in &unregistrations_by_archive {
        unregister_chart_styles(package, stylesheet_id, archive_name, identifiers)?;
    }
    let mut container_removals = HashMap::<String, Vec<u64>>::new();
    let mut ordinary_removals = Vec::new();
    for slot in &removals {
        if is_object_container_archive(package, &slot.archive_name)? {
            container_removals
                .entry(slot.archive_name.clone())
                .or_default()
                .push(slot.object_id);
        } else {
            ordinary_removals.push(slot);
        }
    }
    for slot in ordinary_removals {
        package.update_archive(&slot.archive_name, |archive| {
            archive.remove_object(slot.object_id).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "{drawable_label} chart series non-style {} is missing",
                    slot.object_id
                ))
            })?;
            Ok(())
        })?;
    }
    let mut released_container_ids = Vec::new();
    for (archive_name, identifiers) in container_removals {
        if let Some(container_id) =
            remove_object_container_objects(package, &archive_name, &identifiers)?
        {
            released_container_ids.push(container_id);
        }
    }
    if !styled_creations.is_empty() {
        package.update_archive(chart_archive_name, |archive| {
            for object in styled_creations.drain(..) {
                archive.insert_object(object)?;
            }
            Ok(())
        })?;
    }
    if !unstyled_creations.is_empty() {
        let allocation = object_container.take().ok_or_else(|| {
            Error::InvalidFormat(
                "unstyled chart series creation lost its object container".to_owned(),
            )
        })?;
        insert_object_container(package, chart_archive_name, allocation, unstyled_creations)?;
    }
    for (archive_name, identifiers) in &registrations_by_archive {
        register_chart_styles(package, stylesheet_id, archive_name, identifiers)?;
    }

    patch_chart_series_non_style_references(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        graph.chart_message_index,
        &graph
            .slots
            .iter()
            .flatten()
            .map(|slot| slot.object_id)
            .collect::<Vec<_>>(),
        &final_ids,
    )?;
    update_component_registrations(package, chart_archive_name, &removals, &chart_creation_ids)?;
    if let Some(last_identifier) = created_ids.last().copied() {
        set_package_last_object_identifier(package, last_identifier)?;
    }
    let mut released_ids = removals
        .iter()
        .map(|slot| slot.object_id)
        .collect::<Vec<_>>();
    released_ids.extend(released_container_ids);
    if !released_ids.is_empty() {
        release_package_identifier_suffix(package, &released_ids)?;
    }

    if chart_series_non_style_values(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        expected.len(),
        default.clone(),
        read,
    )? != expected
    {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} {property_label} update failed validation"
        )));
    }
    Ok(())
}

fn allocate_series_non_style_identifier(
    package: &IWorkPackage,
    next_identifier: &mut u64,
    base: NewChartSeriesNonStyleBase,
    object_container: &mut Option<ObjectContainerAllocation>,
) -> Result<u64> {
    if base == NewChartSeriesNonStyleBase::Unstyled && object_container.is_none() {
        *object_container = Some(reserve_object_container(package, next_identifier)?);
    }
    let identifier = *next_identifier;
    *next_identifier = next_identifier
        .checked_add(1)
        .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))?;
    Ok(identifier)
}

fn chart_stylesheet_id(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<u64> {
    let chart_archive = package.archive(chart_archive_name)?;
    let chart_object = chart_archive.object(drawable_object_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} is missing"
        ))
    })?;
    let messages = chart_object
        .messages
        .iter()
        .filter(|message| message.type_ == CHART_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    let [message] = messages.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} must have exactly one chart payload"
        )));
    };
    let chart = IWorkChartArchive::decode(message.data.as_slice())?;
    let chart = chart.chart.as_ref().ok_or_else(|| {
        Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} has no chart archive"
        ))
    })?;
    let style_id = chart
        .series_theme_styles
        .first()
        .map(|reference| reference.identifier)
        .filter(|identifier| *identifier != 0)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "{drawable_label} chart {drawable_object_id} has no series theme style"
            ))
        })?;
    let style_archive_name =
        unique_chart_object_archive_name(package, style_id, "chart series style object")?;
    let style_archive = package.archive(&style_archive_name)?;
    let style_object = style_archive
        .object(style_id)
        .ok_or_else(|| Error::InvalidFormat(format!("chart series style {style_id} is missing")))?;
    let style_messages = style_object
        .messages
        .iter()
        .filter(|message| message.type_ == crate::charts::source::SERIES_STYLE_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    let [style_message] = style_messages.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "chart series style {style_id} must have exactly one series-style payload"
        )));
    };
    tsch::ChartSeriesStyleArchive::decode(style_message.data.as_slice())?
        .super_
        .and_then(|style| style.stylesheet)
        .map(|reference| reference.identifier)
        .filter(|identifier| *identifier != 0)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "chart series style {style_id} has no document stylesheet"
            ))
        })
}

fn read_values_from_graph<T>(
    package: &IWorkPackage,
    graph: &ChartSeriesNonStyleGraph,
    default: &T,
    read: &impl Fn(&[u8]) -> Result<T>,
) -> Result<Vec<T>>
where
    T: Clone,
{
    graph
        .slots
        .iter()
        .map(|slot| {
            slot.as_ref()
                .map_or_else(|| Ok(default.clone()), |slot| slot.read(package, read))
        })
        .collect()
}

fn chart_series_non_style_graph(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    series_count: usize,
) -> Result<ChartSeriesNonStyleGraph> {
    let chart_archive = package.archive(chart_archive_name)?;
    let chart_object = chart_archive.object(drawable_object_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} is missing"
        ))
    })?;
    let mut chart_messages = chart_object
        .messages
        .iter()
        .enumerate()
        .filter(|(_, message)| message.type_ == CHART_MESSAGE_TYPE);
    let Some((chart_message_index, chart_message)) = chart_messages.next() else {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} must have exactly one chart payload"
        )));
    };
    if chart_messages.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} must have exactly one chart payload"
        )));
    }
    let chart = IWorkChartArchive::decode(chart_message.data.as_slice())?;
    let chart = chart.chart.as_ref().ok_or_else(|| {
        Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} has no chart archive"
        ))
    })?;
    let sparse = chart.series_non_styles.as_ref();
    if let Some(sparse) = sparse
        && usize::try_from(sparse.count).ok() != Some(sparse.entries.len())
    {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} declares {} series non-styles but stores {} entries",
            sparse.count,
            sparse.entries.len()
        )));
    }
    let mut slots = vec![None; series_count];
    let mut identifiers = HashSet::new();
    for entry in sparse.into_iter().flat_map(|sparse| &sparse.entries) {
        let index = usize::try_from(entry.index)
            .map_err(|_| Error::InvalidFormat("chart series index exceeds usize".to_owned()))?;
        if index >= series_count {
            return Err(Error::InvalidFormat(format!(
                "{drawable_label} chart {drawable_object_id} series non-style index {index} exceeds series count {series_count}"
            )));
        }
        let identifier = entry.reference.identifier;
        if identifier == 0 || !identifiers.insert(identifier) || slots[index].is_some() {
            return Err(Error::InvalidFormat(format!(
                "{drawable_label} chart {drawable_object_id} has an invalid or repeated series non-style {identifier}"
            )));
        }
        let archive_name =
            unique_chart_object_archive_name(package, identifier, "chart series non-style object")?;
        let archive = package.archive(&archive_name)?;
        let object = archive.object(identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "{drawable_label} chart series non-style {identifier} is missing"
            ))
        })?;
        let mut messages = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == SERIES_NON_STYLE_MESSAGE_TYPE);
        let Some((message_index, _)) = messages.next() else {
            return Err(Error::InvalidFormat(format!(
                "{drawable_label} chart series non-style {identifier} must have exactly one series non-style payload"
            )));
        };
        if messages.next().is_some() {
            return Err(Error::InvalidFormat(format!(
                "{drawable_label} chart series non-style {identifier} must have exactly one series non-style payload"
            )));
        }
        slots[index] = Some(ChartSeriesNonStyleSlot {
            archive_name,
            object_id: identifier,
            message_index,
        });
    }
    Ok(ChartSeriesNonStyleGraph {
        chart_message_index,
        slots,
    })
}

impl ChartSeriesNonStyleSlot {
    fn read<T>(&self, package: &IWorkPackage, read: impl FnOnce(&[u8]) -> Result<T>) -> Result<T> {
        let archive = package.archive(&self.archive_name)?;
        let object = archive.object(self.object_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "chart series non-style {} is missing",
                self.object_id
            ))
        })?;
        let message = object.messages.get(self.message_index).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "chart series non-style {} message index changed unexpectedly",
                self.object_id
            ))
        })?;
        if message.type_ != SERIES_NON_STYLE_MESSAGE_TYPE {
            return Err(Error::InvalidFormat(format!(
                "chart series non-style {} message type changed unexpectedly",
                self.object_id
            )));
        }
        read(message.data.as_slice())
    }

    fn replace(&self, package: &mut IWorkPackage, data: Vec<u8>) -> Result<()> {
        package.update_archive(&self.archive_name, |archive| {
            let object = archive.object_mut(self.object_id).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "chart series non-style {} is missing",
                    self.object_id
                ))
            })?;
            let message = object.messages.get(self.message_index).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "chart series non-style {} message index changed unexpectedly",
                    self.object_id
                ))
            })?;
            if message.type_ != SERIES_NON_STYLE_MESSAGE_TYPE {
                return Err(Error::InvalidFormat(format!(
                    "chart series non-style {} message type changed unexpectedly",
                    self.object_id
                )));
            }
            object.replace_message(
                self.message_index,
                RawMessage {
                    type_: SERIES_NON_STYLE_MESSAGE_TYPE,
                    data,
                },
            )?;
            Ok(())
        })
    }

    fn owner_count(&self, package: &IWorkPackage) -> Result<usize> {
        let mut owner_count = 0usize;
        for archive_name in package.iwa_entry_names() {
            let archive = package.archive(archive_name)?;
            for object in &archive.objects {
                for message in object
                    .messages
                    .iter()
                    .filter(|message| message.type_ == CHART_MESSAGE_TYPE)
                {
                    let chart = IWorkChartArchive::decode(message.data.as_slice())?;
                    if chart.chart.as_ref().is_some_and(|chart| {
                        chart.series_non_styles.as_ref().is_some_and(|sparse| {
                            sparse
                                .entries
                                .iter()
                                .any(|entry| entry.reference.identifier == self.object_id)
                        })
                    }) {
                        owner_count = owner_count.checked_add(1).ok_or_else(|| {
                            Error::InvalidFormat(
                                "chart series non-style owner count overflow".to_owned(),
                            )
                        })?;
                    }
                }
            }
        }
        Ok(owner_count)
    }
}

fn patch_chart_series_non_style_references(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    chart_message_index: usize,
    previous_series_non_style_ids: &[u64],
    identifiers: &[Option<u64>],
) -> Result<()> {
    package.update_archive(chart_archive_name, |archive| {
        let object = archive.object_mut(drawable_object_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "{drawable_label} chart {drawable_object_id} is missing"
            ))
        })?;
        let message = object.messages.get(chart_message_index).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "{drawable_label} chart {drawable_object_id} message index changed unexpectedly"
            ))
        })?;
        if message.type_ != CHART_MESSAGE_TYPE {
            return Err(Error::InvalidFormat(format!(
                "{drawable_label} chart {drawable_object_id} message type changed unexpectedly"
            )));
        }
        let mut chart = IWorkChartArchive::decode(message.data.as_slice())?;
        let chart_payload = chart.chart.as_mut().ok_or_else(|| {
            Error::InvalidFormat(format!(
                "{drawable_label} chart {drawable_object_id} has no chart archive"
            ))
        })?;
        let entries = identifiers
            .iter()
            .enumerate()
            .filter_map(|(index, identifier)| identifier.map(|identifier| (index, identifier)))
            .map(|(index, identifier)| {
                Ok(tsp::sparse_reference_array::Entry {
                    index: u32::try_from(index).map_err(|_| {
                        Error::InvalidFormat("chart series index exceeds u32".to_owned())
                    })?,
                    reference: tsp::Reference {
                        identifier,
                        ..Default::default()
                    },
                })
            })
            .collect::<Result<Vec<_>>>()?;
        chart_payload.series_non_styles = Some(tsp::SparseReferenceArray {
            count: u32::try_from(entries.len()).map_err(|_| {
                Error::InvalidFormat("chart series non-style count exceeds u32".to_owned())
            })?,
            entries,
        });
        let data = chart.encode()?;
        let previous_ids = object.archive_info.message_infos[chart_message_index]
            .object_references
            .clone();
        object.replace_message(
            chart_message_index,
            RawMessage {
                type_: CHART_MESSAGE_TYPE,
                data,
            },
        )?;
        let references =
            &mut object.archive_info.message_infos[chart_message_index].object_references;
        references.clear();
        references.extend(
            previous_ids
                .into_iter()
                .filter(|identifier| !previous_series_non_style_ids.contains(identifier)),
        );
        for identifier in identifiers.iter().copied().flatten() {
            if !references.contains(&identifier) {
                references.push(identifier);
            }
        }
        Ok(())
    })
}

fn update_component_registrations(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    removals: &[ChartSeriesNonStyleSlot],
    chart_created_ids: &[u64],
) -> Result<()> {
    let mut removed_by_component = HashMap::<u64, Vec<u64>>::new();
    for slot in removals {
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

    if !chart_created_ids.is_empty() {
        if let Some(component_id) = component_identifier_for_entry(package, chart_archive_name)? {
            add_component_object_uuids(package, component_id, chart_created_ids)?;
        }
    }
    Ok(())
}

fn chart_series_non_style_object(identifier: u64, data: Vec<u8>) -> Result<ArchiveObject> {
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: SERIES_NON_STYLE_MESSAGE_TYPE,
            data,
        }],
    )?;
    object.archive_info.message_infos[0].versions = STANDARD_MESSAGE_VERSION.to_vec();
    Ok(object)
}

/// Return the native empty private series non-style payload.
pub(crate) fn canonical_empty_chart_series_non_style_data() -> Result<Vec<u8>> {
    canonical_empty_chart_series_non_style_data_with_style(tss::StyleArchive::default())
}

fn canonical_empty_chart_series_non_style_data_with_stylesheet(
    stylesheet_id: u64,
) -> Result<Vec<u8>> {
    canonical_empty_chart_series_non_style_data_with_style(tss::StyleArchive {
        stylesheet: Some(tsp::Reference {
            identifier: stylesheet_id,
            ..Default::default()
        }),
        ..Default::default()
    })
}

fn canonical_empty_chart_series_non_style_data_with_style(
    style: tss::StyleArchive,
) -> Result<Vec<u8>> {
    let mut data = chart_series_non_style_data_with_style(style);
    append_chart_series_non_style_capabilities(&mut data)?;
    Ok(data)
}

fn chart_series_non_style_data_with_style(style: tss::StyleArchive) -> Vec<u8> {
    tsch::ChartSeriesNonStyleArchive {
        super_: Some(style),
    }
    .encode_to_vec()
}

fn append_chart_series_non_style_capabilities(data: &mut Vec<u8>) -> Result<()> {
    append_varint_field(data, SUPPORTS_CUSTOM_NUMBER_FORMAT_FIELD, 1)?;
    append_varint_field(data, SUPPORTS_CUSTOM_DATE_FORMAT_FIELD, 1)?;
    append_varint_field(data, SUPPORTS_CALLOUT_LINES_FIELD, 1)?;
    Ok(())
}

fn chart_series_non_style_stylesheet_id(data: &[u8]) -> Result<Option<u64>> {
    let archive = tsch::ChartSeriesNonStyleArchive::decode(data)?;
    let style = archive.super_.ok_or_else(|| {
        Error::InvalidFormat("chart series non-style has no native style parent".to_owned())
    })?;
    Ok(style.stylesheet.map(|reference| reference.identifier))
}

fn set_chart_series_non_style_stylesheet(
    data: &[u8],
    stylesheet_id: Option<u64>,
) -> Result<Vec<u8>> {
    let current = chart_series_non_style_stylesheet_id(data)?;
    if current == stylesheet_id {
        return Ok(data.to_vec());
    }
    let style = tss::StyleArchive {
        stylesheet: stylesheet_id.map(|identifier| tsp::Reference {
            identifier,
            ..Default::default()
        }),
        ..Default::default()
    }
    .encode_to_vec();
    patch_length_delimited_field(data, 1, true, Some(&style))
}

/// Decode an outer series non-style payload and locate its generated extension.
pub(crate) fn generated_chart_series_non_style_extension(data: &[u8]) -> Result<Option<&[u8]>> {
    tsch::ChartSeriesNonStyleArchive::decode(data)?;
    let fields = parse_wire_fields(data)?;
    let mut extensions = fields
        .iter()
        .filter(|field| field.number == GENERATED_CHART_SERIES_NON_STYLE_EXTENSION_FIELD);
    let Some(extension) = extensions.next() else {
        return Ok(None);
    };
    if extensions.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "chart series non-style extension {GENERATED_CHART_SERIES_NON_STYLE_EXTENSION_FIELD} occurs more than once"
        )));
    }
    if extension.wire_type != 2 {
        return Err(Error::InvalidFormat(format!(
            "chart series non-style extension {GENERATED_CHART_SERIES_NON_STYLE_EXTENSION_FIELD} is not length-delimited"
        )));
    }
    Ok(Some(&data[extension.payload_start..extension.end]))
}

/// Rewrite the generated extension while preserving all unrelated outer bytes.
pub(crate) fn patch_chart_series_non_style_extension(
    data: &[u8],
    expected_present: bool,
    extension: Option<&[u8]>,
) -> Result<Vec<u8>> {
    patch_length_delimited_field(
        data,
        GENERATED_CHART_SERIES_NON_STYLE_EXTENSION_FIELD,
        expected_present,
        extension,
    )
}
