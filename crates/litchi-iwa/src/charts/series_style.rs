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

/// Resolve one effective series style per native series index.
///
/// Native documents keep the baseline styles in `series_theme_styles` and
/// overlay sparse `series_private_styles` entries by index. Source-built
/// charts generally own their theme-style objects directly and have no
/// private overlays.
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
    if chart.series_theme_styles.is_empty() {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} has no series styles"
        )));
    }
    let mut identifiers = chart
        .series_theme_styles
        .iter()
        .map(|reference| reference.identifier)
        .collect::<Vec<_>>();
    let mut seen = HashSet::new();
    for &identifier in &identifiers {
        if identifier == 0 || !seen.insert(identifier) {
            return Err(Error::InvalidFormat(format!(
                "{drawable_label} chart {drawable_object_id} has an invalid or repeated series theme style {identifier}"
            )));
        }
    }
    if let Some(private_styles) = chart.series_private_styles.as_ref() {
        let private_count = usize::try_from(private_styles.count).map_err(|_| {
            Error::InvalidFormat("chart private series-style count exceeds usize".to_owned())
        })?;
        if private_count != private_styles.entries.len() {
            return Err(Error::InvalidFormat(format!(
                "{drawable_label} chart {drawable_object_id} declares {private_count} private series styles but stores {} entries",
                private_styles.entries.len()
            )));
        }
        let mut private_indices = HashSet::new();
        for entry in &private_styles.entries {
            let index = usize::try_from(entry.index).map_err(|_| {
                Error::InvalidFormat("chart private series-style index exceeds usize".to_owned())
            })?;
            let identifier = entry.reference.identifier;
            if index >= identifiers.len()
                || !private_indices.insert(index)
                || identifier == 0
                || !seen.insert(identifier)
            {
                return Err(Error::InvalidFormat(format!(
                    "{drawable_label} chart {drawable_object_id} has an invalid private series style {identifier} at index {index}"
                )));
            }
            identifiers[index] = identifier;
        }
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

/// Resolve the effective style slot for every data series.
pub(crate) fn effective_chart_series_style_slots(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    series_count: usize,
) -> Result<Vec<ChartSeriesStyleSlot>> {
    let mut slots = chart_series_style_slots(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    if slots.len() < series_count {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} has {} series styles for {series_count} series",
            slots.len()
        )));
    }
    slots.truncate(series_count);
    Ok(slots)
}

impl ChartSeriesStyleSlot {
    pub(crate) fn archive_name(&self) -> &str {
        &self.archive_name
    }

    pub(crate) const fn object_id(&self) -> u64 {
        self.object_id
    }

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

    /// Resolve one property through the native series-style parent chain.
    ///
    /// Sparse private styles store only their overrides and point at a theme
    /// style through `TSS.StyleArchive.parent`. Callers return `None` when the
    /// current style does not define the requested property.
    pub(crate) fn read_inherited<T>(
        &self,
        package: &IWorkPackage,
        read: impl Fn(&[u8]) -> Result<Option<T>> + Copy,
    ) -> Result<Option<T>> {
        let mut identifier = self.object_id;
        let mut archive_name = self.archive_name.clone();
        let mut visited = HashSet::new();
        loop {
            if !visited.insert(identifier) {
                return Err(Error::InvalidFormat(format!(
                    "chart series style parent cycle contains {identifier}"
                )));
            }
            let archive = package.archive(&archive_name)?;
            let object = archive.object(identifier).ok_or_else(|| {
                Error::InvalidFormat(format!("chart series style {identifier} is missing"))
            })?;
            let messages = object
                .messages
                .iter()
                .filter(|message| message.type_ == SERIES_STYLE_MESSAGE_TYPE)
                .collect::<Vec<_>>();
            let [message] = messages.as_slice() else {
                return Err(Error::InvalidFormat(format!(
                    "chart series style {identifier} must have exactly one series-style payload"
                )));
            };
            if let Some(value) = read(message.data.as_slice())? {
                return Ok(Some(value));
            }
            let parent = tsch::ChartSeriesStyleArchive::decode(message.data.as_slice())?
                .super_
                .and_then(|style| style.parent)
                .map(|reference| reference.identifier)
                .filter(|identifier| *identifier != 0);
            let Some(parent) = parent else {
                return Ok(None);
            };
            identifier = parent;
            archive_name =
                unique_chart_object_archive_name(package, parent, "chart series parent style")?;
        }
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
