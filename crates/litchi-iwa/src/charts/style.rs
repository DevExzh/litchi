//! Shared lossless access to a native chart-style payload.
//!
//! Chart-level presentation controls are stored in the generated extension of
//! a chart's `TSCH.ChartStyleArchive`. This module resolves the one private
//! style object referenced by a chart and provides guarded wire-level access
//! for focused chart-style feature modules.

use prost::Message;

use crate::archive::RawMessage;
use crate::charts::IWorkChartArchive;
use crate::charts::source::{CHART_MESSAGE_TYPE, CHART_STYLE_MESSAGE_TYPE};
use crate::charts::unique_chart_object_archive_name;
use crate::protobuf::tsch;
use crate::wire::parse_wire_fields;
use crate::{Error, IWorkPackage, Result};

/// Proto2 extension holding the generated chart-style properties.
pub(crate) const GENERATED_CHART_STYLE_EXTENSION_FIELD: u32 = 10_000;

/// The single mutable native chart-style payload for one chart.
#[derive(Debug)]
pub(crate) struct ChartStyleSlot {
    archive_name: String,
    object_id: u64,
    message_index: usize,
}

/// Resolve the native chart-style payload referenced by one chart.
pub(crate) fn chart_style_slot(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<ChartStyleSlot> {
    let chart_archive = package.archive(chart_archive_name)?;
    let chart_object = chart_archive.object(drawable_object_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} is missing"
        ))
    })?;
    let mut chart_messages = chart_object
        .messages
        .iter()
        .filter(|message| message.type_ == CHART_MESSAGE_TYPE);
    let Some(chart_message) = chart_messages.next() else {
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
    let style_id = chart
        .chart
        .as_ref()
        .and_then(|payload| payload.chart_style.as_ref())
        .map(|reference| reference.identifier)
        .filter(|identifier| *identifier != 0)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "{drawable_label} chart {drawable_object_id} has no chart style"
            ))
        })?;
    let archive_name = unique_chart_object_archive_name(package, style_id, "chart style object")?;
    let archive = package.archive(&archive_name)?;
    let style_object = archive.object(style_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "{drawable_label} chart style {style_id} is missing"
        ))
    })?;
    let mut messages = style_object
        .messages
        .iter()
        .enumerate()
        .filter(|(_, message)| message.type_ == CHART_STYLE_MESSAGE_TYPE);
    let Some((message_index, _)) = messages.next() else {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart style {style_id} must have exactly one chart-style payload"
        )));
    };
    if messages.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart style {style_id} must have exactly one chart-style payload"
        )));
    }
    Ok(ChartStyleSlot {
        archive_name,
        object_id: style_id,
        message_index,
    })
}

impl ChartStyleSlot {
    pub(crate) fn archive_name(&self) -> &str {
        &self.archive_name
    }

    pub(crate) const fn object_id(&self) -> u64 {
        self.object_id
    }

    /// Read the resolved chart-style bytes without allocating a rewritten archive.
    pub(crate) fn read<T>(
        &self,
        package: &IWorkPackage,
        read: impl FnOnce(&[u8]) -> Result<T>,
    ) -> Result<T> {
        let archive = package.archive(&self.archive_name)?;
        let object = archive.object(self.object_id).ok_or_else(|| {
            Error::InvalidFormat(format!("chart style {} is missing", self.object_id))
        })?;
        let message = object.messages.get(self.message_index).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "chart style {} message index changed unexpectedly",
                self.object_id
            ))
        })?;
        if message.type_ != CHART_STYLE_MESSAGE_TYPE {
            return Err(Error::InvalidFormat(format!(
                "chart style {} message type changed unexpectedly",
                self.object_id
            )));
        }
        read(message.data.as_slice())
    }

    /// Reject a mutation that would silently affect another chart.
    pub(crate) fn ensure_exclusive(
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
                    if chart
                        .chart
                        .as_ref()
                        .and_then(|payload| payload.chart_style.as_ref())
                        .is_some_and(|reference| reference.identifier == self.object_id)
                    {
                        owner_count = owner_count.checked_add(1).ok_or_else(|| {
                            Error::InvalidFormat("chart style owner count overflow".to_owned())
                        })?;
                    }
                }
            }
        }
        if owner_count != 1 {
            return Err(Error::InvalidFormat(format!(
                "{drawable_label} chart {drawable_object_id} style {} is shared by {owner_count} charts",
                self.object_id
            )));
        }
        Ok(())
    }

    /// Transactionally rewrite the resolved style message.
    pub(crate) fn update(
        &self,
        package: &mut IWorkPackage,
        patch: impl FnOnce(&[u8]) -> Result<Vec<u8>>,
    ) -> Result<()> {
        package.update_archive(&self.archive_name, |archive| {
            let object = archive.object_mut(self.object_id).ok_or_else(|| {
                Error::InvalidFormat(format!("chart style {} is missing", self.object_id))
            })?;
            let original = object.messages.get(self.message_index).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "chart style {} message index changed unexpectedly",
                    self.object_id
                ))
            })?;
            if original.type_ != CHART_STYLE_MESSAGE_TYPE {
                return Err(Error::InvalidFormat(format!(
                    "chart style {} message type changed unexpectedly",
                    self.object_id
                )));
            }
            let data = patch(original.data.as_slice())?;
            object.replace_message(
                self.message_index,
                RawMessage {
                    type_: CHART_STYLE_MESSAGE_TYPE,
                    data,
                },
            )?;
            Ok(())
        })
    }
}

/// Decode the outer chart-style payload and locate its generated extension.
pub(crate) fn generated_chart_style_extension(data: &[u8]) -> Result<Option<&[u8]>> {
    tsch::ChartStyleArchive::decode(data)?;
    let fields = parse_wire_fields(data)?;
    let mut extensions = fields
        .iter()
        .filter(|field| field.number() == GENERATED_CHART_STYLE_EXTENSION_FIELD);
    let Some(extension) = extensions.next() else {
        return Ok(None);
    };
    if extensions.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "chart style extension {GENERATED_CHART_STYLE_EXTENSION_FIELD} occurs more than once"
        )));
    }
    if extension.wire_type() != 2 {
        return Err(Error::InvalidFormat(format!(
            "chart style extension {GENERATED_CHART_STYLE_EXTENSION_FIELD} is not length-delimited"
        )));
    }
    extension.validate_canonical_framing(data)?;
    Ok(Some(extension.checked_payload(data)?))
}

/// Replace schema-known fields in one selected message while retaining every
/// unknown field's original wire bytes. The selected message is intentionally
/// treated as a narrow projection: callers provide the complete replacement
/// for the known field set, while producer extensions remain opaque.
pub(crate) fn replace_known_wire_fields_preserving_unknown(
    existing: &[u8],
    replacement: &[u8],
    known_fields: &[u32],
) -> Result<Vec<u8>> {
    if existing == replacement {
        return Ok(existing.to_vec());
    }
    let existing_fields = parse_wire_fields(existing)?;
    let replacement_fields = parse_wire_fields(replacement)?;
    let is_known = |number: u32| known_fields.contains(&number);
    let first_existing_known = existing_fields
        .iter()
        .position(|field| is_known(field.number()));
    let has_known_overlap = existing_fields.iter().any(|existing_field| {
        is_known(existing_field.number())
            && replacement_fields
                .iter()
                .any(|replacement_field| replacement_field.number() == existing_field.number())
    });

    let replacement_size = replacement_fields
        .iter()
        .try_fold(0usize, |size, field| {
            size.checked_add(field.end() - field.start())
        })
        .ok_or_else(|| {
            Error::InvalidFormat("selected wire replacement size overflow".to_owned())
        })?;
    let unknown_size = existing_fields
        .iter()
        .filter(|field| !is_known(field.number()))
        .try_fold(0usize, |size, field| {
            size.checked_add(field.end() - field.start())
        })
        .ok_or_else(|| Error::InvalidFormat("selected wire payload size overflow".to_owned()))?;
    let capacity = unknown_size
        .checked_add(replacement_size)
        .ok_or_else(|| Error::InvalidFormat("selected wire output size overflow".to_owned()))?;
    let mut output = Vec::new();
    output.try_reserve_exact(capacity).map_err(|_| {
        Error::IwaCommon(litchi_iwa_common::Error::Allocation {
            resource: "selected chart-style wire output",
            amount: capacity,
        })
    })?;
    if !has_known_overlap {
        let mut inserted = false;
        for (index, field) in existing_fields.iter().enumerate() {
            if is_known(field.number()) {
                if first_existing_known == Some(index) && !inserted {
                    output.extend_from_slice(replacement);
                    inserted = true;
                }
                continue;
            }
            output.extend_from_slice(field.raw(existing)?);
        }
        if !inserted {
            output.extend_from_slice(replacement);
        }
        return Ok(output);
    }

    let mut emitted_known = Vec::new();
    for field in &existing_fields {
        if is_known(field.number()) {
            if emitted_known.contains(&field.number()) {
                continue;
            }
            if let Some(replacement_field) = replacement_fields
                .iter()
                .find(|candidate| candidate.number() == field.number())
            {
                output.extend_from_slice(replacement_field.raw(replacement)?);
                emitted_known.push(field.number());
            }
            continue;
        }
        output.extend_from_slice(field.raw(existing)?);
    }
    for replacement_field in &replacement_fields {
        if !is_known(replacement_field.number())
            || !existing_fields
                .iter()
                .any(|field| field.number() == replacement_field.number())
        {
            output.extend_from_slice(replacement_field.raw(replacement)?);
        }
    }
    Ok(output)
}
