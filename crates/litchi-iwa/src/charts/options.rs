//! Lossless native chart-option storage and mutation.
//!
//! iWork chart title and legend controls live in the generated extension of a
//! chart's `TSCH.ChartNonStyleArchive`. This module keeps both the outer
//! non-style payload and the generated extension lossless while updating only
//! their requested native option fields.

use prost::Message;

use crate::archive::RawMessage;
use crate::charts::IWorkChartArchive;
use crate::charts::source::{CHART_MESSAGE_TYPE, CHART_NON_STYLE_MESSAGE_TYPE};
use crate::protobuf::tsch;
use crate::wire::{parse_wire_fields, patch_length_delimited_field, patch_varint_field};
use crate::{Error, IWorkPackage, Result};

/// Proto2 extension holding the generated chart non-style properties.
const GENERATED_CHART_NON_STYLE_EXTENSION_FIELD: u32 = 10_000;
/// `tschchartinfodefaultshowlegend` in `TSCH.Generated.ChartNonStyleArchive`.
const CHART_LEGEND_VISIBLE_FIELD: u32 = 20;
/// `tschchartinfodefaultshowtitle` in `TSCH.Generated.ChartNonStyleArchive`.
const CHART_TITLE_VISIBLE_FIELD: u32 = 21;
/// `tschchartinfodefaulttitle` in `TSCH.Generated.ChartNonStyleArchive`.
const CHART_TITLE_TEXT_FIELD: u32 = 23;

/// The single mutable native chart non-style payload for one chart.
#[derive(Debug)]
struct ChartNonStyleSlot {
    archive_name: String,
    object_id: u64,
    message_index: usize,
}

/// Read one chart title from its native chart non-style extension.
pub(crate) fn chart_title(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<Option<String>> {
    chart_non_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?
    .read(package, read_chart_non_style_title)
}

/// Set one chart title, enabling the native title switch when necessary.
pub(crate) fn set_chart_title(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    title: &str,
) -> Result<()> {
    let slot = chart_non_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    if slot.read(package, read_chart_non_style_title)?.as_deref() == Some(title) {
        return Ok(());
    }
    slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
    slot.update(package, |data| {
        patch_chart_non_style_title(data, Some(title))
    })?;
    if slot.read(package, read_chart_non_style_title)?.as_deref() != Some(title) {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} title update failed validation"
        )));
    }
    Ok(())
}

/// Remove one visible chart title.
///
/// Returns whether the title was visible. A chart non-style shared by more than
/// one chart is rejected rather than silently changing another chart.
pub(crate) fn remove_chart_title(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<bool> {
    let slot = chart_non_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    if slot.read(package, read_chart_non_style_title)?.is_none() {
        return Ok(false);
    }
    slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
    slot.update(package, |data| patch_chart_non_style_title(data, None))?;
    if slot.read(package, read_chart_non_style_title)?.is_some() {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} title removal failed validation"
        )));
    }
    Ok(true)
}

/// Read whether one chart shows its native legend.
pub(crate) fn chart_legend_visible(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<bool> {
    chart_non_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?
    .read(package, read_chart_non_style_legend_visible)
}

/// Set whether one chart shows its native legend.
pub(crate) fn set_chart_legend_visible(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    visible: bool,
) -> Result<()> {
    let slot = chart_non_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    if slot.read(package, read_chart_non_style_legend_visible)? == visible {
        return Ok(());
    }
    slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
    slot.update(package, |data| {
        patch_chart_non_style_legend_visibility(data, visible)
    })?;
    if slot.read(package, read_chart_non_style_legend_visible)? != visible {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} legend update failed validation"
        )));
    }
    Ok(())
}

/// Decode a `TSCH.ChartNonStyleArchive` and return its visible native title.
pub(crate) fn read_chart_non_style_title(data: &[u8]) -> Result<Option<String>> {
    let Some(extension) = generated_chart_non_style_extension(data)? else {
        return Ok(None);
    };
    let generated = tsch::generated::ChartNonStyleArchive::decode(extension)?;
    if generated.tschchartinfodefaultshowtitle != Some(true) {
        return Ok(None);
    }
    Ok(Some(
        generated.tschchartinfodefaulttitle.unwrap_or_default(),
    ))
}

/// Decode a `TSCH.ChartNonStyleArchive` and return its native legend switch.
pub(crate) fn read_chart_non_style_legend_visible(data: &[u8]) -> Result<bool> {
    let Some(extension) = generated_chart_non_style_extension(data)? else {
        return Ok(false);
    };
    let generated = tsch::generated::ChartNonStyleArchive::decode(extension)?;
    Ok(generated.tschchartinfodefaultshowlegend.unwrap_or(false))
}

fn chart_non_style_slot(
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
    let mut chart_messages = chart_object
        .messages
        .iter()
        .enumerate()
        .filter(|(_, message)| message.type_ == CHART_MESSAGE_TYPE);
    let Some((_, chart_message)) = chart_messages.next() else {
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
    let archive_name = unique_object_archive_name(package, non_style_id)?;
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
    fn read<T>(&self, package: &IWorkPackage, read: impl FnOnce(&[u8]) -> Result<T>) -> Result<T> {
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

    fn update(
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

fn patch_chart_non_style_title(data: &[u8], title: Option<&str>) -> Result<Vec<u8>> {
    let Some(extension) = generated_chart_non_style_extension(data)? else {
        let Some(title) = title else {
            return Ok(data.to_vec());
        };
        let generated = tsch::generated::ChartNonStyleArchive {
            tschchartinfodefaultshowtitle: Some(true),
            tschchartinfodefaulttitle: Some(title.to_owned()),
            ..Default::default()
        };
        let patched = patch_length_delimited_field(
            data,
            GENERATED_CHART_NON_STYLE_EXTENSION_FIELD,
            false,
            Some(generated.encode_to_vec().as_slice()),
        )?;
        validate_patched_title(&patched, Some(title))?;
        return Ok(patched);
    };

    let generated = tsch::generated::ChartNonStyleArchive::decode(extension)?;
    let visible_present = generated.tschchartinfodefaultshowtitle.is_some();
    let title_present = generated.tschchartinfodefaulttitle.is_some();
    let extension = patch_varint_field(
        extension,
        CHART_TITLE_VISIBLE_FIELD,
        visible_present,
        Some(u64::from(title.is_some())),
    )?;
    let extension = patch_length_delimited_field(
        &extension,
        CHART_TITLE_TEXT_FIELD,
        title_present,
        title.map(str::as_bytes),
    )?;
    let patched = patch_length_delimited_field(
        data,
        GENERATED_CHART_NON_STYLE_EXTENSION_FIELD,
        true,
        Some(extension.as_slice()),
    )?;
    validate_patched_title(&patched, title)?;
    Ok(patched)
}

fn patch_chart_non_style_legend_visibility(data: &[u8], visible: bool) -> Result<Vec<u8>> {
    let Some(extension) = generated_chart_non_style_extension(data)? else {
        let generated = tsch::generated::ChartNonStyleArchive {
            tschchartinfodefaultshowlegend: Some(visible),
            ..Default::default()
        };
        let patched = patch_length_delimited_field(
            data,
            GENERATED_CHART_NON_STYLE_EXTENSION_FIELD,
            false,
            Some(generated.encode_to_vec().as_slice()),
        )?;
        validate_patched_legend_visibility(&patched, visible)?;
        return Ok(patched);
    };

    let generated = tsch::generated::ChartNonStyleArchive::decode(extension)?;
    let visible_present = generated.tschchartinfodefaultshowlegend.is_some();
    let extension = patch_varint_field(
        extension,
        CHART_LEGEND_VISIBLE_FIELD,
        visible_present,
        Some(u64::from(visible)),
    )?;
    let patched = patch_length_delimited_field(
        data,
        GENERATED_CHART_NON_STYLE_EXTENSION_FIELD,
        true,
        Some(extension.as_slice()),
    )?;
    validate_patched_legend_visibility(&patched, visible)?;
    Ok(patched)
}

fn validate_patched_title(data: &[u8], expected: Option<&str>) -> Result<()> {
    if read_chart_non_style_title(data)?.as_deref() != expected {
        return Err(Error::InvalidFormat(
            "chart non-style title wire patch failed validation".to_owned(),
        ));
    }
    Ok(())
}

fn validate_patched_legend_visibility(data: &[u8], expected: bool) -> Result<()> {
    if read_chart_non_style_legend_visible(data)? != expected {
        return Err(Error::InvalidFormat(
            "chart non-style legend wire patch failed validation".to_owned(),
        ));
    }
    Ok(())
}

fn generated_chart_non_style_extension(data: &[u8]) -> Result<Option<&[u8]>> {
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

fn unique_object_archive_name(package: &IWorkPackage, identifier: u64) -> Result<String> {
    let mut archive_name = None;
    for name in package.iwa_entry_names() {
        if package.archive(name)?.object(identifier).is_none() {
            continue;
        }
        if archive_name.replace(name.to_owned()).is_some() {
            return Err(Error::Archive(format!(
                "chart non-style object {identifier} occurs in multiple IWA components"
            )));
        }
    }
    archive_name.ok_or_else(|| {
        Error::InvalidFormat(format!("chart non-style object {identifier} is missing"))
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::protobuf::tss;
    use crate::wire::{append_length_delimited_field, append_varint_field};

    const UNMAPPED_OUTER_FIELD: u32 = 4_096;
    const UNMAPPED_GENERATED_FIELD: u32 = 4_097;
    const UNMAPPED_VALUE: u64 = 42;

    #[test]
    fn title_patch_retains_unmapped_chart_non_style_fields() {
        let generated = tsch::generated::ChartNonStyleArchive {
            tschchartinfodefaultshowlegend: Some(true),
            tschchartinfodefaultshowtitle: Some(false),
            ..Default::default()
        };
        let mut extension = generated.encode_to_vec();
        append_varint_field(&mut extension, UNMAPPED_GENERATED_FIELD, UNMAPPED_VALUE).unwrap();
        let base = tsch::ChartNonStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        };
        let mut original = base.encode_to_vec();
        append_length_delimited_field(
            &mut original,
            GENERATED_CHART_NON_STYLE_EXTENSION_FIELD,
            &extension,
        )
        .unwrap();
        append_varint_field(&mut original, UNMAPPED_OUTER_FIELD, UNMAPPED_VALUE).unwrap();

        let titled = patch_chart_non_style_title(&original, Some("Revenue by region")).unwrap();
        assert_eq!(
            read_chart_non_style_title(&titled).unwrap(),
            Some("Revenue by region".to_owned())
        );
        assert_eq!(
            raw_field(&titled, UNMAPPED_OUTER_FIELD),
            raw_field(&original, UNMAPPED_OUTER_FIELD)
        );
        assert_eq!(
            raw_field(
                generated_chart_non_style_extension(&titled)
                    .unwrap()
                    .unwrap(),
                UNMAPPED_GENERATED_FIELD
            ),
            raw_field(
                generated_chart_non_style_extension(&original)
                    .unwrap()
                    .unwrap(),
                UNMAPPED_GENERATED_FIELD
            )
        );

        let removed = patch_chart_non_style_title(&titled, None).unwrap();
        assert_eq!(read_chart_non_style_title(&removed).unwrap(), None);
        assert_eq!(removed, original);
    }

    #[test]
    fn legend_patch_retains_title_and_unmapped_chart_non_style_fields() {
        let generated = tsch::generated::ChartNonStyleArchive {
            tschchartinfodefaultshowlegend: Some(true),
            tschchartinfodefaultshowtitle: Some(true),
            tschchartinfodefaulttitle: Some("Revenue by region".to_owned()),
            ..Default::default()
        };
        let mut extension = generated.encode_to_vec();
        append_varint_field(&mut extension, UNMAPPED_GENERATED_FIELD, UNMAPPED_VALUE).unwrap();
        let base = tsch::ChartNonStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        };
        let mut original = base.encode_to_vec();
        append_length_delimited_field(
            &mut original,
            GENERATED_CHART_NON_STYLE_EXTENSION_FIELD,
            &extension,
        )
        .unwrap();
        append_varint_field(&mut original, UNMAPPED_OUTER_FIELD, UNMAPPED_VALUE).unwrap();

        let hidden = patch_chart_non_style_legend_visibility(&original, false).unwrap();
        assert!(!read_chart_non_style_legend_visible(&hidden).unwrap());
        assert_eq!(
            read_chart_non_style_title(&hidden).unwrap(),
            Some("Revenue by region".to_owned())
        );
        assert_eq!(
            raw_field(&hidden, UNMAPPED_OUTER_FIELD),
            raw_field(&original, UNMAPPED_OUTER_FIELD)
        );
        assert_eq!(
            raw_field(
                generated_chart_non_style_extension(&hidden)
                    .unwrap()
                    .unwrap(),
                UNMAPPED_GENERATED_FIELD
            ),
            raw_field(
                generated_chart_non_style_extension(&original)
                    .unwrap()
                    .unwrap(),
                UNMAPPED_GENERATED_FIELD
            )
        );

        let visible = patch_chart_non_style_legend_visibility(&hidden, true).unwrap();
        assert!(read_chart_non_style_legend_visible(&visible).unwrap());
        assert_eq!(visible, original);
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
