//! Shared lossless access to a native chart non-style payload.
//!
//! Chart-level behavioral controls are stored in the generated extension of a
//! chart's `TSCH.ChartNonStyleArchive`. This module resolves the one private
//! non-style object referenced by a chart and provides guarded wire-level
//! access for focused chart feature modules.

use prost::Message;

use crate::archive::RawMessage;
use crate::charts::IWorkChartArchive;
use crate::charts::source::{CHART_MESSAGE_TYPE, CHART_NON_STYLE_MESSAGE_TYPE};
use crate::charts::unique_chart_object_archive_name;
use crate::protobuf::tsch;
use crate::wire::parse_wire_fields;
use crate::{Error, IWorkPackage, Result};

/// Proto2 extension holding the generated chart non-style properties.
pub(crate) const GENERATED_CHART_NON_STYLE_EXTENSION_FIELD: u32 = 10_000;

/// The single mutable native chart non-style payload for one chart.
#[derive(Debug)]
pub(crate) struct ChartNonStyleSlot {
    archive_name: String,
    object_id: u64,
    message_index: usize,
}

/// Resolve the native chart non-style payload referenced by one chart.
pub(crate) fn chart_non_style_slot(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<ChartNonStyleSlot> {
    let chart_archive = package.archive(chart_archive_name)?;
    let chart_object = chart_archive.object(drawable_object_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} is missing"
        ))
    })?;
    let mut messages = chart_object
        .messages
        .iter()
        .filter(|message| message.type_ == CHART_MESSAGE_TYPE);
    let Some(message) = messages.next() else {
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
    let non_style_id = chart
        .chart
        .as_ref()
        .and_then(|chart| chart.chart_non_style.as_ref())
        .map(|reference| reference.identifier)
        .filter(|identifier| *identifier != 0)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "{drawable_label} chart {drawable_object_id} has no chart non-style"
            ))
        })?;
    let archive_name =
        unique_chart_object_archive_name(package, non_style_id, "chart non-style object")?;
    let archive = package.archive(&archive_name)?;
    let non_style_object = archive.object(non_style_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "{drawable_label} chart non-style {non_style_id} is missing"
        ))
    })?;
    let mut messages = non_style_object
        .messages
        .iter()
        .enumerate()
        .filter(|(_, message)| message.type_ == CHART_NON_STYLE_MESSAGE_TYPE);
    let Some((message_index, _)) = messages.next() else {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart non-style {non_style_id} must have exactly one chart non-style payload"
        )));
    };
    if messages.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart non-style {non_style_id} must have exactly one chart non-style payload"
        )));
    }
    Ok(ChartNonStyleSlot {
        archive_name,
        object_id: non_style_id,
        message_index,
    })
}

impl ChartNonStyleSlot {
    /// Read the resolved chart non-style bytes without allocating a rewrite.
    pub(crate) fn read<T>(
        &self,
        package: &IWorkPackage,
        read: impl FnOnce(&[u8]) -> Result<T>,
    ) -> Result<T> {
        let archive = package.archive(&self.archive_name)?;
        let object = archive.object(self.object_id).ok_or_else(|| {
            Error::InvalidFormat(format!("chart non-style {} is missing", self.object_id))
        })?;
        let message = object.messages.get(self.message_index).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "chart non-style {} message index changed unexpectedly",
                self.object_id
            ))
        })?;
        if message.type_ != CHART_NON_STYLE_MESSAGE_TYPE {
            return Err(Error::InvalidFormat(format!(
                "chart non-style {} message type changed unexpectedly",
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
                        .and_then(|chart| chart.chart_non_style.as_ref())
                        .is_some_and(|reference| reference.identifier == self.object_id)
                    {
                        owner_count = owner_count.checked_add(1).ok_or_else(|| {
                            Error::InvalidFormat("chart non-style owner count overflow".to_owned())
                        })?;
                    }
                }
            }
        }
        if owner_count != 1 {
            return Err(Error::InvalidFormat(format!(
                "{drawable_label} chart {drawable_object_id} non-style {} is shared by {owner_count} charts",
                self.object_id
            )));
        }
        Ok(())
    }

    /// Transactionally rewrite the resolved non-style message.
    pub(crate) fn update(
        &self,
        package: &mut IWorkPackage,
        patch: impl FnOnce(&[u8]) -> Result<Vec<u8>>,
    ) -> Result<()> {
        package.update_archive(&self.archive_name, |archive| {
            let object = archive.object_mut(self.object_id).ok_or_else(|| {
                Error::InvalidFormat(format!("chart non-style {} is missing", self.object_id))
            })?;
            let original = object.messages.get(self.message_index).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "chart non-style {} message index changed unexpectedly",
                    self.object_id
                ))
            })?;
            if original.type_ != CHART_NON_STYLE_MESSAGE_TYPE {
                return Err(Error::InvalidFormat(format!(
                    "chart non-style {} message type changed unexpectedly",
                    self.object_id
                )));
            }
            let data = patch(original.data.as_slice())?;
            object.replace_message(
                self.message_index,
                RawMessage {
                    type_: CHART_NON_STYLE_MESSAGE_TYPE,
                    data,
                },
            )?;
            Ok(())
        })
    }
}

/// Decode the outer non-style payload and locate its generated extension.
pub(crate) fn generated_chart_non_style_extension(data: &[u8]) -> Result<Option<&[u8]>> {
    tsch::ChartNonStyleArchive::decode(data)?;
    let fields = parse_wire_fields(data)?;
    let mut extensions = fields
        .iter()
        .filter(|field| field.number == GENERATED_CHART_NON_STYLE_EXTENSION_FIELD);
    let Some(extension) = extensions.next() else {
        return Ok(None);
    };
    if extensions.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "chart non-style extension {GENERATED_CHART_NON_STYLE_EXTENSION_FIELD} occurs more than once"
        )));
    }
    if extension.wire_type != 2 {
        return Err(Error::InvalidFormat(format!(
            "chart non-style extension {GENERATED_CHART_NON_STYLE_EXTENSION_FIELD} is not length-delimited"
        )));
    }
    Ok(Some(&data[extension.payload_start..extension.end]))
}
