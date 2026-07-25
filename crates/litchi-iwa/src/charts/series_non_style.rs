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
use crate::charts::source::{CHART_MESSAGE_TYPE, SERIES_NON_STYLE_MESSAGE_TYPE};
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
    let current = read_values_from_graph(package, &graph, &default, &read)?;
    if current == expected {
        return Ok(());
    }

    let canonical_empty = canonical_empty_chart_series_non_style_data()?;
    let mut next_identifier = next_object_identifier(package)?;
    let mut final_ids = graph
        .slots
        .iter()
        .map(|slot| slot.as_ref().map(|slot| slot.object_id))
        .collect::<Vec<_>>();
    let mut updates = Vec::new();
    let mut removals = Vec::new();
    let mut creations = Vec::new();
    for (index, ((slot, current), replacement)) in
        graph.slots.iter().zip(&current).zip(expected).enumerate()
    {
        if current == replacement {
            continue;
        }
        if let Some(slot) = slot {
            slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
            let patched = slot.read(package, |data| patch(data, replacement))?;
            if patched == canonical_empty {
                final_ids[index] = None;
                removals.push(slot.clone());
            } else {
                updates.push((slot.clone(), patched));
            }
        } else if replacement != &default {
            let data = patch(canonical_empty.as_slice(), replacement)?;
            if data == canonical_empty {
                return Err(Error::InvalidFormat(format!(
                    "{drawable_label} chart {drawable_object_id} {property_label} patch produced no native override"
                )));
            }
            let identifier = next_identifier;
            next_identifier = next_identifier
                .checked_add(1)
                .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))?;
            final_ids[index] = Some(identifier);
            creations.push((identifier, chart_series_non_style_object(identifier, data)?));
        }
    }

    for (slot, data) in updates {
        slot.replace(package, data)?;
    }
    for slot in &removals {
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
    let created_ids = creations
        .iter()
        .map(|(identifier, _)| *identifier)
        .collect::<Vec<_>>();
    if !creations.is_empty() {
        package.update_archive(chart_archive_name, |archive| {
            for (_, object) in creations.drain(..) {
                archive.insert_object(object)?;
            }
            Ok(())
        })?;
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
    update_component_registrations(package, chart_archive_name, &removals, &created_ids)?;

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

    fn ensure_exclusive(
        &self,
        package: &IWorkPackage,
        drawable_object_id: u64,
        drawable_label: &str,
    ) -> Result<()> {
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
        if owner_count != 1 {
            return Err(Error::InvalidFormat(format!(
                "{drawable_label} chart {drawable_object_id} series non-style {} is shared by {owner_count} charts",
                self.object_id
            )));
        }
        Ok(())
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
    created_ids: &[u64],
) -> Result<()> {
    let removed_ids = removals
        .iter()
        .map(|slot| slot.object_id)
        .collect::<Vec<_>>();
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

    if !created_ids.is_empty() {
        if let Some(component_id) = component_identifier_for_entry(package, chart_archive_name)? {
            add_component_object_uuids(package, component_id, created_ids)?;
        }
        let last_identifier = *created_ids.last().ok_or_else(|| {
            Error::InvalidFormat("chart series creation lost allocated identifiers".to_owned())
        })?;
        set_package_last_object_identifier(package, last_identifier)?;
    }
    if !removed_ids.is_empty() {
        release_package_identifier_suffix(package, &removed_ids)?;
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
    let mut data = tsch::ChartSeriesNonStyleArchive {
        super_: Some(tss::StyleArchive::default()),
    }
    .encode_to_vec();
    append_varint_field(&mut data, SUPPORTS_CUSTOM_NUMBER_FORMAT_FIELD, 1)?;
    append_varint_field(&mut data, SUPPORTS_CUSTOM_DATE_FORMAT_FIELD, 1)?;
    append_varint_field(&mut data, SUPPORTS_CALLOUT_LINES_FIELD, 1)?;
    Ok(data)
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
