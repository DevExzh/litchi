//! Lossless per-wedge position CRUD for native pie and donut charts.
//!
//! iWork stores each wedge's distance from the chart center in a sparse array
//! of private `TSCH.ChartSeriesNonStyleArchive` objects. This module validates
//! that graph, creates and removes private objects transactionally, and patches
//! only the generated wedge-explosion scalar.

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
use crate::wire::{
    append_length_delimited_field, append_varint_field, parse_wire_fields, patch_fixed32_field,
    patch_length_delimited_field,
};
use crate::{Error, IWorkPackage, Result};

const GENERATED_SERIES_NON_STYLE_EXTENSION_FIELD: u32 = 10_000;
/// `tschchartseriespiewedgeexplosion` in the generated series non-style.
const PIE_WEDGE_EXPLOSION_FIELD: u32 = 63;
const SUPPORTS_CUSTOM_NUMBER_FORMAT_FIELD: u32 = 10_001;
const SUPPORTS_CUSTOM_DATE_FORMAT_FIELD: u32 = 10_002;
const SUPPORTS_CALLOUT_LINES_FIELD: u32 = 10_003;
const MINIMUM_WEDGE_EXPLOSION_PERCENT: f32 = 0.0;
const MAXIMUM_WEDGE_EXPLOSION_PERCENT: f32 = 100.0;
const MAXIMUM_WEDGE_EXPLOSION_FRACTION: f32 = 1.0;
const PERCENT_SCALE: f32 = 100.0;
const STANDARD_MESSAGE_VERSION: &[u32] = &[1, 0, 5];

/// Distance of one pie or donut wedge from the chart center.
///
/// Values use the percentage displayed by the Wedges inspector and must be
/// finite in the inclusive range `0%..=100%`.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct ChartPieWedgeExplosion(f32);

impl ChartPieWedgeExplosion {
    /// Wedge at its native, unseparated position.
    pub const ZERO: Self = Self(MINIMUM_WEDGE_EXPLOSION_PERCENT);
    /// Wedge at the inspector's maximum distance from the center.
    pub const MAXIMUM: Self = Self(MAXIMUM_WEDGE_EXPLOSION_FRACTION);

    /// Construct a wedge position from an inspector percentage.
    pub fn from_percent(percent: f32) -> Result<Self> {
        if !percent.is_finite()
            || !(MINIMUM_WEDGE_EXPLOSION_PERCENT..=MAXIMUM_WEDGE_EXPLOSION_PERCENT)
                .contains(&percent)
        {
            return Err(Error::InvalidFormat(format!(
                "chart pie wedge explosion must be finite and within {MINIMUM_WEDGE_EXPLOSION_PERCENT}%..={MAXIMUM_WEDGE_EXPLOSION_PERCENT}%"
            )));
        }
        Ok(Self(percent / PERCENT_SCALE))
    }

    /// Return the percentage displayed by iWork.
    pub fn percent(self) -> f32 {
        self.0 * PERCENT_SCALE
    }

    fn from_native(fraction: f32) -> Result<Self> {
        if !fraction.is_finite() || !(0.0..=1.0).contains(&fraction) {
            return Err(Error::InvalidFormat(format!(
                "native chart pie wedge explosion {fraction} must be finite and within 0.0..=1.0"
            )));
        }
        Ok(Self(fraction))
    }

    const fn native_fraction(self) -> f32 {
        self.0
    }
}

impl Default for ChartPieWedgeExplosion {
    fn default() -> Self {
        Self::ZERO
    }
}

impl TryFrom<f32> for ChartPieWedgeExplosion {
    type Error = Error;

    fn try_from(percent: f32) -> Result<Self> {
        Self::from_percent(percent)
    }
}

/// Zero-based index of one pie or donut wedge in chart series order.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct ChartPieWedgeIndex(usize);

impl ChartPieWedgeIndex {
    /// Construct a zero-based wedge index.
    pub const fn from_zero_based(index: usize) -> Self {
        Self(index)
    }

    /// Return the zero-based wedge index.
    pub const fn zero_based(self) -> usize {
        self.0
    }
}

#[derive(Debug, Clone)]
struct SeriesNonStyleSlot {
    archive_name: String,
    object_id: u64,
    message_index: usize,
}

#[derive(Debug)]
struct SeriesNonStyleGraph {
    chart_message_index: usize,
    slots: Vec<Option<SeriesNonStyleSlot>>,
}

/// Read every wedge position in chart series order.
pub(crate) fn chart_pie_wedge_explosions(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    series_count: usize,
) -> Result<Vec<ChartPieWedgeExplosion>> {
    let graph = series_non_style_graph(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        series_count,
    )?;
    graph
        .slots
        .iter()
        .map(|slot| {
            slot.as_ref()
                .map_or(Ok(ChartPieWedgeExplosion::ZERO), |slot| {
                    slot.read(package, read_series_non_style_explosion)
                })
        })
        .collect()
}

/// Set every wedge position in chart series order.
pub(crate) fn set_chart_pie_wedge_explosions(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    expected: &[ChartPieWedgeExplosion],
) -> Result<()> {
    if expected.is_empty() {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} has no pie wedges"
        )));
    }
    let graph = series_non_style_graph(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        expected.len(),
    )?;
    let current = graph
        .slots
        .iter()
        .map(|slot| {
            slot.as_ref()
                .map_or(Ok(ChartPieWedgeExplosion::ZERO), |slot| {
                    slot.read(package, read_series_non_style_explosion)
                })
        })
        .collect::<Result<Vec<_>>>()?;
    if current == expected {
        return Ok(());
    }

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
            let patched = slot.read(package, |data| {
                patch_series_non_style_explosion(data, *replacement)
            })?;
            if *replacement == ChartPieWedgeExplosion::ZERO
                && patched == canonical_empty_series_non_style_data()?
            {
                final_ids[index] = None;
                removals.push(slot.clone());
            } else {
                updates.push((slot.clone(), patched));
            }
        } else if *replacement != ChartPieWedgeExplosion::ZERO {
            let identifier = next_identifier;
            next_identifier = next_identifier
                .checked_add(1)
                .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))?;
            final_ids[index] = Some(identifier);
            creations.push((
                identifier,
                series_non_style_object(identifier, *replacement)?,
            ));
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

    let removed_ids = removals
        .iter()
        .map(|slot| slot.object_id)
        .collect::<Vec<_>>();
    let mut removed_by_component = HashMap::<u64, Vec<u64>>::new();
    for slot in &removals {
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

    let created_ids = final_ids
        .iter()
        .copied()
        .flatten()
        .filter(|identifier| {
            graph
                .slots
                .iter()
                .flatten()
                .all(|slot| slot.object_id != *identifier)
        })
        .collect::<Vec<_>>();
    if !created_ids.is_empty() {
        if let Some(component_id) = component_identifier_for_entry(package, chart_archive_name)? {
            add_component_object_uuids(package, component_id, &created_ids)?;
        }
        let last_identifier = *created_ids.last().ok_or_else(|| {
            Error::InvalidFormat("chart wedge creation lost allocated identifiers".to_owned())
        })?;
        set_package_last_object_identifier(package, last_identifier)?;
    }
    if !removed_ids.is_empty() {
        release_package_identifier_suffix(package, &removed_ids)?;
    }

    if chart_pie_wedge_explosions(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        expected.len(),
    )? != expected
    {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} pie wedge-explosion update failed validation"
        )));
    }
    Ok(())
}

fn series_non_style_graph(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    series_count: usize,
) -> Result<SeriesNonStyleGraph> {
    let series_count_u32 = u32::try_from(series_count)
        .map_err(|_| Error::InvalidFormat("chart series count exceeds u32".to_owned()))?;
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
        && sparse.count != 0
        && sparse.count != series_count_u32
    {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} series non-style count {} does not match series count {series_count}",
            sparse.count
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
        slots[index] = Some(SeriesNonStyleSlot {
            archive_name,
            object_id: identifier,
            message_index,
        });
    }
    Ok(SeriesNonStyleGraph {
        chart_message_index,
        slots,
    })
}

impl SeriesNonStyleSlot {
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
            count: if entries.is_empty() {
                0
            } else {
                u32::try_from(identifiers.len()).map_err(|_| {
                    Error::InvalidFormat("chart series count exceeds u32".to_owned())
                })?
            },
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

fn series_non_style_object(
    identifier: u64,
    explosion: ChartPieWedgeExplosion,
) -> Result<ArchiveObject> {
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: SERIES_NON_STYLE_MESSAGE_TYPE,
            data: canonical_series_non_style_data(explosion)?,
        }],
    )?;
    object.archive_info.message_infos[0].versions = STANDARD_MESSAGE_VERSION.to_vec();
    Ok(object)
}

fn canonical_series_non_style_data(explosion: ChartPieWedgeExplosion) -> Result<Vec<u8>> {
    if explosion == ChartPieWedgeExplosion::ZERO {
        return canonical_empty_series_non_style_data();
    }
    let generated = tsch::generated::ChartSeriesNonStyleArchive {
        tschchartseriespiewedgeexplosion: Some(explosion.native_fraction()),
        ..Default::default()
    };
    let mut data = tsch::ChartSeriesNonStyleArchive {
        super_: Some(tss::StyleArchive::default()),
    }
    .encode_to_vec();
    append_length_delimited_field(
        &mut data,
        GENERATED_SERIES_NON_STYLE_EXTENSION_FIELD,
        &generated.encode_to_vec(),
    )?;
    append_support_fields(&mut data)?;
    Ok(data)
}

fn canonical_empty_series_non_style_data() -> Result<Vec<u8>> {
    let mut data = tsch::ChartSeriesNonStyleArchive {
        super_: Some(tss::StyleArchive::default()),
    }
    .encode_to_vec();
    append_support_fields(&mut data)?;
    Ok(data)
}

fn append_support_fields(data: &mut Vec<u8>) -> Result<()> {
    append_varint_field(data, SUPPORTS_CUSTOM_NUMBER_FORMAT_FIELD, 1)?;
    append_varint_field(data, SUPPORTS_CUSTOM_DATE_FORMAT_FIELD, 1)?;
    append_varint_field(data, SUPPORTS_CALLOUT_LINES_FIELD, 1)
}

fn read_series_non_style_explosion(data: &[u8]) -> Result<ChartPieWedgeExplosion> {
    let Some(extension) = generated_series_non_style_extension(data)? else {
        return Ok(ChartPieWedgeExplosion::ZERO);
    };
    let generated = tsch::generated::ChartSeriesNonStyleArchive::decode(extension)?;
    generated
        .tschchartseriespiewedgeexplosion
        .map(ChartPieWedgeExplosion::from_native)
        .transpose()
        .map(|value| value.unwrap_or(ChartPieWedgeExplosion::ZERO))
}

fn patch_series_non_style_explosion(
    data: &[u8],
    explosion: ChartPieWedgeExplosion,
) -> Result<Vec<u8>> {
    let Some(extension) = generated_series_non_style_extension(data)? else {
        if explosion == ChartPieWedgeExplosion::ZERO {
            return Ok(data.to_vec());
        }
        let generated = tsch::generated::ChartSeriesNonStyleArchive {
            tschchartseriespiewedgeexplosion: Some(explosion.native_fraction()),
            ..Default::default()
        };
        let patched = patch_length_delimited_field(
            data,
            GENERATED_SERIES_NON_STYLE_EXTENSION_FIELD,
            false,
            Some(generated.encode_to_vec().as_slice()),
        )?;
        validate_patched_explosion(&patched, explosion)?;
        return Ok(patched);
    };

    let generated = tsch::generated::ChartSeriesNonStyleArchive::decode(extension)?;
    let present = generated.tschchartseriespiewedgeexplosion.is_some();
    let native =
        (explosion != ChartPieWedgeExplosion::ZERO).then(|| explosion.native_fraction().to_bits());
    let extension = patch_fixed32_field(extension, PIE_WEDGE_EXPLOSION_FIELD, present, native)?;
    let patched = patch_length_delimited_field(
        data,
        GENERATED_SERIES_NON_STYLE_EXTENSION_FIELD,
        true,
        (!extension.is_empty()).then_some(extension.as_slice()),
    )?;
    validate_patched_explosion(&patched, explosion)?;
    Ok(patched)
}

fn generated_series_non_style_extension(data: &[u8]) -> Result<Option<&[u8]>> {
    tsch::ChartSeriesNonStyleArchive::decode(data)?;
    let fields = parse_wire_fields(data)?;
    let mut extensions = fields
        .iter()
        .filter(|field| field.number == GENERATED_SERIES_NON_STYLE_EXTENSION_FIELD);
    let Some(extension) = extensions.next() else {
        return Ok(None);
    };
    if extensions.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "chart series non-style extension {GENERATED_SERIES_NON_STYLE_EXTENSION_FIELD} occurs more than once"
        )));
    }
    if extension.wire_type != 2 {
        return Err(Error::InvalidFormat(format!(
            "chart series non-style extension {GENERATED_SERIES_NON_STYLE_EXTENSION_FIELD} is not length-delimited"
        )));
    }
    Ok(Some(&data[extension.payload_start..extension.end]))
}

fn validate_patched_explosion(data: &[u8], expected: ChartPieWedgeExplosion) -> Result<()> {
    if read_series_non_style_explosion(data)? != expected {
        return Err(Error::InvalidFormat(
            "chart pie wedge-explosion wire patch failed validation".to_owned(),
        ));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    const UNMAPPED_OUTER_FIELD: u32 = 4_096;
    const UNMAPPED_GENERATED_FIELD: u32 = 4_097;
    const UNMAPPED_VALUE: u64 = 42;

    #[test]
    fn wedge_explosions_are_strict_percentages() {
        assert_eq!(
            ChartPieWedgeExplosion::default(),
            ChartPieWedgeExplosion::ZERO
        );
        assert_eq!(
            ChartPieWedgeExplosion::from_percent(25.0)
                .unwrap()
                .percent(),
            25.0
        );
        assert_eq!(
            ChartPieWedgeExplosion::from_percent(100.0).unwrap(),
            ChartPieWedgeExplosion::MAXIMUM
        );
        for invalid in [f32::NEG_INFINITY, -0.1, 100.1, f32::INFINITY, f32::NAN] {
            assert!(ChartPieWedgeExplosion::from_percent(invalid).is_err());
        }
    }

    #[test]
    fn wedge_explosion_patch_is_lossless_and_resets_exactly() {
        let mut generated = tsch::generated::ChartSeriesNonStyleArchive::default().encode_to_vec();
        append_varint_field(&mut generated, UNMAPPED_GENERATED_FIELD, UNMAPPED_VALUE).unwrap();
        let mut original = tsch::ChartSeriesNonStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();
        append_length_delimited_field(
            &mut original,
            GENERATED_SERIES_NON_STYLE_EXTENSION_FIELD,
            &generated,
        )
        .unwrap();
        append_varint_field(&mut original, UNMAPPED_OUTER_FIELD, UNMAPPED_VALUE).unwrap();
        let customized = ChartPieWedgeExplosion::from_percent(25.0).unwrap();

        let patched = patch_series_non_style_explosion(&original, customized).unwrap();
        assert_eq!(
            read_series_non_style_explosion(&patched).unwrap(),
            customized
        );
        assert_eq!(
            raw_field(&patched, UNMAPPED_OUTER_FIELD),
            raw_field(&original, UNMAPPED_OUTER_FIELD)
        );
        assert_eq!(
            raw_field(
                generated_series_non_style_extension(&patched)
                    .unwrap()
                    .unwrap(),
                UNMAPPED_GENERATED_FIELD,
            ),
            raw_field(
                generated_series_non_style_extension(&original)
                    .unwrap()
                    .unwrap(),
                UNMAPPED_GENERATED_FIELD,
            )
        );
        assert_eq!(
            patch_series_non_style_explosion(&patched, ChartPieWedgeExplosion::ZERO).unwrap(),
            original
        );
    }

    #[test]
    fn malformed_native_wedge_explosions_are_rejected() {
        for invalid in [-0.1, 1.1, f32::INFINITY, f32::NAN] {
            let mut outer = canonical_empty_series_non_style_data().unwrap();
            let generated = tsch::generated::ChartSeriesNonStyleArchive {
                tschchartseriespiewedgeexplosion: Some(invalid),
                ..Default::default()
            };
            outer = patch_length_delimited_field(
                &outer,
                GENERATED_SERIES_NON_STYLE_EXTENSION_FIELD,
                false,
                Some(generated.encode_to_vec().as_slice()),
            )
            .unwrap();
            assert!(read_series_non_style_explosion(&outer).is_err());
        }
    }

    fn raw_field(data: &[u8], number: u32) -> Vec<Vec<u8>> {
        parse_wire_fields(data)
            .unwrap()
            .into_iter()
            .filter(|field| field.number == number)
            .map(|field| data[field.start..field.end].to_vec())
            .collect()
    }
}
