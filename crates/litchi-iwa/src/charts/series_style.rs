//! Shared lossless access to chart-series style payloads.

use std::collections::HashSet;

use prost::Message;

use crate::archive::RawMessage;
use crate::charts::IWorkChartArchive;
use crate::charts::source::{CHART_MESSAGE_TYPE, SERIES_STYLE_MESSAGE_TYPE};
use crate::charts::unique_chart_object_archive_name;
use crate::protobuf::tsch;
use crate::wire::parse_wire_fields;
use crate::{Error, IWorkPackage, Result};

/// Proto2 extension holding generated chart-series style properties.
pub(crate) const GENERATED_CHART_SERIES_STYLE_EXTENSION_FIELD: u32 = 10_000;

/// One mutable native series-style payload referenced by a chart.
#[derive(Debug)]
pub(crate) struct ChartSeriesStyleSlot {
    archive_name: String,
    object_id: u64,
    message_index: usize,
}

/// Resolve every unique theme/private series style referenced by one chart.
pub(crate) fn chart_series_style_slots(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<Vec<ChartSeriesStyleSlot>> {
    let archive = package.archive(chart_archive_name)?;
    let object = archive.object(drawable_object_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} is missing"
        ))
    })?;
    let messages = object
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
    let private_styles = chart
        .series_private_styles
        .as_ref()
        .into_iter()
        .flat_map(|styles| styles.entries.iter().map(|entry| &entry.reference));
    let mut seen = HashSet::new();
    let identifiers = chart
        .series_theme_styles
        .iter()
        .chain(private_styles)
        .map(|reference| reference.identifier)
        .map(|identifier| {
            if identifier == 0 || !seen.insert(identifier) {
                return Err(Error::InvalidFormat(format!(
                    "{drawable_label} chart {drawable_object_id} has an invalid or repeated series style {identifier}"
                )));
            }
            Ok(identifier)
        })
        .collect::<Result<Vec<_>>>()?;
    if identifiers.is_empty() {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} has no series styles"
        )));
    }

    identifiers
        .into_iter()
        .map(|identifier| {
            let archive_name =
                unique_chart_object_archive_name(package, identifier, "chart series style object")?;
            let archive = package.archive(&archive_name)?;
            let object = archive.object(identifier).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "{drawable_label} chart series style {identifier} is missing"
                ))
            })?;
            let messages = object
                .messages
                .iter()
                .enumerate()
                .filter(|(_, message)| message.type_ == SERIES_STYLE_MESSAGE_TYPE)
                .collect::<Vec<_>>();
            let [(message_index, _)] = messages.as_slice() else {
                return Err(Error::InvalidFormat(format!(
                    "{drawable_label} chart series style {identifier} must have exactly one series-style payload"
                )));
            };
            Ok(ChartSeriesStyleSlot {
                archive_name,
                object_id: identifier,
                message_index: *message_index,
            })
        })
        .collect()
}

impl ChartSeriesStyleSlot {
    pub(crate) fn read<T>(
        &self,
        package: &IWorkPackage,
        read: impl FnOnce(&[u8]) -> Result<T>,
    ) -> Result<T> {
        let archive = package.archive(&self.archive_name)?;
        let object = archive.object(self.object_id).ok_or_else(|| {
            Error::InvalidFormat(format!("chart series style {} is missing", self.object_id))
        })?;
        let message = object.messages.get(self.message_index).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "chart series style {} message index changed unexpectedly",
                self.object_id
            ))
        })?;
        if message.type_ != SERIES_STYLE_MESSAGE_TYPE {
            return Err(Error::InvalidFormat(format!(
                "chart series style {} message type changed unexpectedly",
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
                    let Some(chart) = chart.chart.as_ref() else {
                        continue;
                    };
                    let private_owns = chart.series_private_styles.as_ref().is_some_and(|styles| {
                        styles
                            .entries
                            .iter()
                            .any(|entry| entry.reference.identifier == self.object_id)
                    });
                    if private_owns
                        || chart
                            .series_theme_styles
                            .iter()
                            .any(|reference| reference.identifier == self.object_id)
                    {
                        owner_count = owner_count.checked_add(1).ok_or_else(|| {
                            Error::InvalidFormat(
                                "chart series-style owner count overflow".to_owned(),
                            )
                        })?;
                    }
                }
            }
        }
        if owner_count != 1 {
            return Err(Error::InvalidFormat(format!(
                "{drawable_label} chart {drawable_object_id} series style {} is shared by {owner_count} charts",
                self.object_id
            )));
        }
        Ok(())
    }

    pub(crate) fn update(
        &self,
        package: &mut IWorkPackage,
        patch: impl FnOnce(&[u8]) -> Result<Vec<u8>>,
    ) -> Result<()> {
        package.update_archive(&self.archive_name, |archive| {
            let object = archive.object_mut(self.object_id).ok_or_else(|| {
                Error::InvalidFormat(format!("chart series style {} is missing", self.object_id))
            })?;
            let original = object.messages.get(self.message_index).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "chart series style {} message index changed unexpectedly",
                    self.object_id
                ))
            })?;
            if original.type_ != SERIES_STYLE_MESSAGE_TYPE {
                return Err(Error::InvalidFormat(format!(
                    "chart series style {} message type changed unexpectedly",
                    self.object_id
                )));
            }
            let data = patch(original.data.as_slice())?;
            object.replace_message(
                self.message_index,
                RawMessage {
                    type_: SERIES_STYLE_MESSAGE_TYPE,
                    data,
                },
            )?;
            Ok(())
        })
    }
}

/// Decode an outer series-style payload and locate its generated extension.
pub(crate) fn generated_chart_series_style_extension(data: &[u8]) -> Result<Option<&[u8]>> {
    tsch::ChartSeriesStyleArchive::decode(data)?;
    let fields = parse_wire_fields(data)?;
    let extensions = fields
        .iter()
        .filter(|field| field.number == GENERATED_CHART_SERIES_STYLE_EXTENSION_FIELD)
        .collect::<Vec<_>>();
    let [extension] = extensions.as_slice() else {
        if extensions.is_empty() {
            return Ok(None);
        }
        return Err(Error::InvalidFormat(format!(
            "chart series-style extension {GENERATED_CHART_SERIES_STYLE_EXTENSION_FIELD} occurs more than once"
        )));
    };
    if extension.wire_type != 2 {
        return Err(Error::InvalidFormat(format!(
            "chart series-style extension {GENERATED_CHART_SERIES_STYLE_EXTENSION_FIELD} is not length-delimited"
        )));
    }
    Ok(Some(&data[extension.payload_start..extension.end]))
}
