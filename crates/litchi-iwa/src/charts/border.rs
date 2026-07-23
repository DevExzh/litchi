//! Lossless native chart-border storage and mutation.
//!
//! iWork stores the Chart Options `Border` switch in the generated extension
//! of a chart's `TSCH.ChartStyleArchive`. This module resolves the private
//! style object, preserves both protobuf layers losslessly, and changes only
//! the native border switch.

use prost::Message;

use crate::archive::RawMessage;
use crate::charts::IWorkChartArchive;
use crate::charts::source::{CHART_MESSAGE_TYPE, CHART_STYLE_MESSAGE_TYPE};
use crate::charts::unique_chart_object_archive_name;
use crate::protobuf::tsch;
use crate::wire::{parse_wire_fields, patch_length_delimited_field, patch_varint_field};
use crate::{Error, IWorkPackage, Result};

/// Proto2 extension holding the generated chart-style properties.
const GENERATED_CHART_STYLE_EXTENSION_FIELD: u32 = 10_000;
/// `tschchartinfodefaultshowborder` in `TSCH.Generated.ChartStyleArchive`.
const CHART_BORDER_VISIBLE_FIELD: u32 = 18;

/// The single mutable native chart-style payload for one chart.
#[derive(Debug)]
struct ChartStyleSlot {
    archive_name: String,
    object_id: u64,
    message_index: usize,
}

/// Read whether one native chart shows its chart-area border.
pub(crate) fn chart_border_visible(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<bool> {
    chart_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?
    .read(package, read_chart_border_visible)
}

/// Set whether one native chart shows its chart-area border.
pub(crate) fn set_chart_border_visible(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    visible: bool,
) -> Result<()> {
    let slot = chart_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    if slot.read(package, read_chart_border_visible)? == visible {
        return Ok(());
    }
    slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
    slot.update(package, |data| patch_chart_border_visibility(data, visible))?;
    if slot.read(package, read_chart_border_visible)? != visible {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} border update failed validation"
        )));
    }
    Ok(())
}

fn chart_style_slot(
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
    fn read<T>(&self, package: &IWorkPackage, read: impl FnOnce(&[u8]) -> Result<T>) -> Result<T> {
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

    fn update(
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

fn read_chart_border_visible(data: &[u8]) -> Result<bool> {
    let Some(extension) = generated_chart_style_extension(data)? else {
        return Ok(false);
    };
    let generated = tsch::generated::ChartStyleArchive::decode(extension)?;
    Ok(generated.tschchartinfodefaultshowborder.unwrap_or(false))
}

fn patch_chart_border_visibility(data: &[u8], visible: bool) -> Result<Vec<u8>> {
    let Some(extension) = generated_chart_style_extension(data)? else {
        if !visible {
            return Ok(data.to_vec());
        }
        let generated = tsch::generated::ChartStyleArchive {
            tschchartinfodefaultshowborder: Some(true),
            ..Default::default()
        };
        let extension = generated.encode_to_vec();
        let patched = patch_length_delimited_field(
            data,
            GENERATED_CHART_STYLE_EXTENSION_FIELD,
            false,
            Some(extension.as_slice()),
        )?;
        validate_patched_chart_border_visibility(&patched, visible)?;
        return Ok(patched);
    };

    let generated = tsch::generated::ChartStyleArchive::decode(extension)?;
    let visible_present = generated.tschchartinfodefaultshowborder.is_some();
    let value = (visible_present || visible).then_some(u64::from(visible));
    let extension = patch_varint_field(
        extension,
        CHART_BORDER_VISIBLE_FIELD,
        visible_present,
        value,
    )?;
    let patched = patch_length_delimited_field(
        data,
        GENERATED_CHART_STYLE_EXTENSION_FIELD,
        true,
        Some(extension.as_slice()),
    )?;
    validate_patched_chart_border_visibility(&patched, visible)?;
    Ok(patched)
}

fn validate_patched_chart_border_visibility(data: &[u8], expected: bool) -> Result<()> {
    if read_chart_border_visible(data)? != expected {
        return Err(Error::InvalidFormat(
            "chart border wire patch failed validation".to_owned(),
        ));
    }
    Ok(())
}

fn generated_chart_style_extension(data: &[u8]) -> Result<Option<&[u8]>> {
    tsch::ChartStyleArchive::decode(data)?;
    let fields = parse_wire_fields(data)?;
    let mut extensions = fields
        .iter()
        .filter(|field| field.number == GENERATED_CHART_STYLE_EXTENSION_FIELD);
    let Some(extension) = extensions.next() else {
        return Ok(None);
    };
    if extensions.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "chart style extension {GENERATED_CHART_STYLE_EXTENSION_FIELD} occurs more than once"
        )));
    }
    if extension.wire_type != 2 {
        return Err(Error::InvalidFormat(format!(
            "chart style extension {GENERATED_CHART_STYLE_EXTENSION_FIELD} is not length-delimited"
        )));
    }
    Ok(Some(&data[extension.payload_start..extension.end]))
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
    fn border_patch_retains_other_style_fields_and_unmapped_data() {
        let generated = tsch::generated::ChartStyleArchive {
            tschchartinfodefaultshowborder: Some(false),
            tschchartinfodefaultgridbackgroundopacity: Some(1.0),
            tschchartinfodefaultinterbargap: Some(0.2),
            ..Default::default()
        };
        let mut extension = generated.encode_to_vec();
        append_varint_field(&mut extension, UNMAPPED_GENERATED_FIELD, UNMAPPED_VALUE).unwrap();
        let base = tsch::ChartStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        };
        let mut original = base.encode_to_vec();
        append_length_delimited_field(
            &mut original,
            GENERATED_CHART_STYLE_EXTENSION_FIELD,
            &extension,
        )
        .unwrap();
        append_varint_field(&mut original, UNMAPPED_OUTER_FIELD, UNMAPPED_VALUE).unwrap();

        let visible = patch_chart_border_visibility(&original, true).unwrap();
        assert!(read_chart_border_visible(&visible).unwrap());
        let patched_generated = tsch::generated::ChartStyleArchive::decode(
            generated_chart_style_extension(&visible).unwrap().unwrap(),
        )
        .unwrap();
        assert_eq!(
            patched_generated.tschchartinfodefaultgridbackgroundopacity,
            Some(1.0)
        );
        assert_eq!(patched_generated.tschchartinfodefaultinterbargap, Some(0.2));
        assert_eq!(
            raw_field(&visible, UNMAPPED_OUTER_FIELD),
            raw_field(&original, UNMAPPED_OUTER_FIELD)
        );
        assert_eq!(
            raw_field(
                generated_chart_style_extension(&visible).unwrap().unwrap(),
                UNMAPPED_GENERATED_FIELD,
            ),
            raw_field(
                generated_chart_style_extension(&original).unwrap().unwrap(),
                UNMAPPED_GENERATED_FIELD,
            )
        );

        let restored = patch_chart_border_visibility(&visible, false).unwrap();
        assert_eq!(restored, original);
    }

    #[test]
    fn borders_default_hidden_and_create_an_extension_when_needed() {
        let original = tsch::ChartStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();
        assert!(!read_chart_border_visible(&original).unwrap());
        assert_eq!(
            patch_chart_border_visibility(&original, false).unwrap(),
            original
        );

        let visible = patch_chart_border_visibility(&original, true).unwrap();
        assert!(read_chart_border_visible(&visible).unwrap());
        assert!(generated_chart_style_extension(&visible).unwrap().is_some());
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
