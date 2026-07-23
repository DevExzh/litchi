//! Lossless native chart-axis line visibility storage and mutation.
//!
//! iWork stores category- and value-axis line switches in the generated
//! extension of a chart's `TSCH.ChartAxisStyleArchive`. This module identifies
//! the primary native axis-style object, preserves both protobuf layers
//! losslessly, and changes only the requested axis-line field.

use prost::Message;

use crate::archive::RawMessage;
use crate::charts::ChartAxis;
use crate::charts::IWorkChartArchive;
use crate::charts::source::{AXIS_STYLE_MESSAGE_TYPE, CHART_MESSAGE_TYPE};
use crate::charts::unique_chart_object_archive_name;
use crate::protobuf::tsch;
use crate::wire::{parse_wire_fields, patch_length_delimited_field, patch_varint_field};
use crate::{Error, IWorkPackage, Result};

/// Proto2 extension holding the generated chart-axis style properties.
const GENERATED_CHART_AXIS_STYLE_EXTENSION_FIELD: u32 = 10_000;
/// `tschchartaxiscategoryshowaxis` in `TSCH.Generated.ChartAxisStyleArchive`.
const CATEGORY_AXIS_LINE_VISIBLE_FIELD: u32 = 24;
/// `tschchartaxisvalueshowaxis` in `TSCH.Generated.ChartAxisStyleArchive`.
const VALUE_AXIS_LINE_VISIBLE_FIELD: u32 = 25;

/// The single mutable native axis-style payload for one chart axis.
#[derive(Debug)]
struct AxisStyleSlot {
    archive_name: String,
    object_id: u64,
    message_index: usize,
}

/// Read whether iWork shows one native chart-axis line.
pub(crate) fn chart_axis_line_visible(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    axis: ChartAxis,
) -> Result<bool> {
    axis_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        axis,
    )?
    .read(package, |data| read_axis_style_line_visibility(data, axis))
}

/// Set whether iWork shows one native chart-axis line.
pub(crate) fn set_chart_axis_line_visible(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    axis: ChartAxis,
    visible: bool,
) -> Result<()> {
    let slot = axis_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        axis,
    )?;
    if slot.read(package, |data| read_axis_style_line_visibility(data, axis))? == visible {
        return Ok(());
    }
    slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
    slot.update(package, |data| {
        patch_axis_style_line_visibility(data, axis, visible)
    })?;
    if slot.read(package, |data| read_axis_style_line_visibility(data, axis))? != visible {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} {}-axis line update failed validation",
            axis.label()
        )));
    }
    Ok(())
}

/// Decode a `TSCH.ChartAxisStyleArchive` and return its native axis-line switch.
fn read_axis_style_line_visibility(data: &[u8], axis: ChartAxis) -> Result<bool> {
    let Some(extension) = generated_axis_style_extension(data)? else {
        return Ok(false);
    };
    let generated = tsch::generated::ChartAxisStyleArchive::decode(extension)?;
    let visible = match axis {
        ChartAxis::Category => generated.tschchartaxiscategoryshowaxis,
        ChartAxis::Value => generated.tschchartaxisvalueshowaxis,
    };
    Ok(visible.unwrap_or(false))
}

fn axis_style_slot(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    axis: ChartAxis,
) -> Result<AxisStyleSlot> {
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
    let payload = chart.chart.as_ref().ok_or_else(|| {
        Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} has no chart payload"
        ))
    })?;
    let style_id = axis.primary_style_identifier(payload).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} has no primary {}-axis style",
            axis.label()
        ))
    })?;
    let archive_name =
        unique_chart_object_archive_name(package, style_id, "chart axis-style object")?;
    let archive = package.archive(&archive_name)?;
    let style_object = archive.object(style_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "{drawable_label} chart {}-axis style {style_id} is missing",
            axis.label()
        ))
    })?;
    let mut messages = style_object
        .messages
        .iter()
        .enumerate()
        .filter(|(_, message)| message.type_ == AXIS_STYLE_MESSAGE_TYPE);
    let Some((message_index, _)) = messages.next() else {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {}-axis style {style_id} must have exactly one axis-style payload",
            axis.label()
        )));
    };
    if messages.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {}-axis style {style_id} must have exactly one axis-style payload",
            axis.label()
        )));
    }
    Ok(AxisStyleSlot {
        archive_name,
        object_id: style_id,
        message_index,
    })
}

impl AxisStyleSlot {
    fn read<T>(&self, package: &IWorkPackage, read: impl FnOnce(&[u8]) -> Result<T>) -> Result<T> {
        let archive = package.archive(&self.archive_name)?;
        let object = archive.object(self.object_id).ok_or_else(|| {
            Error::InvalidFormat(format!("chart axis-style {} is missing", self.object_id))
        })?;
        let message = object.messages.get(self.message_index).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "chart axis-style {} message index changed unexpectedly",
                self.object_id
            ))
        })?;
        if message.type_ != AXIS_STYLE_MESSAGE_TYPE {
            return Err(Error::InvalidFormat(format!(
                "chart axis-style {} message type changed unexpectedly",
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
        let mut reference_count = 0usize;
        for archive_name in package.iwa_entry_names() {
            let archive = package.archive(archive_name)?;
            for object in &archive.objects {
                for message in object
                    .messages
                    .iter()
                    .filter(|message| message.type_ == CHART_MESSAGE_TYPE)
                {
                    let chart = IWorkChartArchive::decode(message.data.as_slice())?;
                    let Some(payload) = chart.chart.as_ref() else {
                        continue;
                    };
                    let references = payload
                        .value_axis_styles
                        .iter()
                        .chain(payload.category_axis_styles.iter());
                    for reference in references {
                        if reference.identifier == self.object_id {
                            reference_count = reference_count.checked_add(1).ok_or_else(|| {
                                Error::InvalidFormat(
                                    "chart axis-style reference count overflow".to_owned(),
                                )
                            })?;
                        }
                    }
                }
            }
        }
        if reference_count != 1 {
            return Err(Error::InvalidFormat(format!(
                "{drawable_label} chart {drawable_object_id} axis-style {} is referenced by {reference_count} chart axes",
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
                Error::InvalidFormat(format!("chart axis-style {} is missing", self.object_id))
            })?;
            let original = object.messages.get(self.message_index).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "chart axis-style {} message index changed unexpectedly",
                    self.object_id
                ))
            })?;
            if original.type_ != AXIS_STYLE_MESSAGE_TYPE {
                return Err(Error::InvalidFormat(format!(
                    "chart axis-style {} message type changed unexpectedly",
                    self.object_id
                )));
            }
            let data = patch(original.data.as_slice())?;
            object.replace_message(
                self.message_index,
                RawMessage {
                    type_: AXIS_STYLE_MESSAGE_TYPE,
                    data,
                },
            )?;
            Ok(())
        })
    }
}

fn patch_axis_style_line_visibility(
    data: &[u8],
    axis: ChartAxis,
    visible: bool,
) -> Result<Vec<u8>> {
    let Some(extension) = generated_axis_style_extension(data)? else {
        let mut generated = tsch::generated::ChartAxisStyleArchive::default();
        match axis {
            ChartAxis::Category => generated.tschchartaxiscategoryshowaxis = Some(visible),
            ChartAxis::Value => generated.tschchartaxisvalueshowaxis = Some(visible),
        }
        let encoded = generated.encode_to_vec();
        let patched = patch_length_delimited_field(
            data,
            GENERATED_CHART_AXIS_STYLE_EXTENSION_FIELD,
            false,
            Some(encoded.as_slice()),
        )?;
        validate_patched_axis_line_visibility(&patched, axis, visible)?;
        return Ok(patched);
    };

    let generated = tsch::generated::ChartAxisStyleArchive::decode(extension)?;
    let visible_present = match axis {
        ChartAxis::Category => generated.tschchartaxiscategoryshowaxis.is_some(),
        ChartAxis::Value => generated.tschchartaxisvalueshowaxis.is_some(),
    };
    let extension = patch_varint_field(
        extension,
        axis_line_visible_field(axis),
        visible_present,
        Some(u64::from(visible)),
    )?;
    let patched = patch_length_delimited_field(
        data,
        GENERATED_CHART_AXIS_STYLE_EXTENSION_FIELD,
        true,
        Some(extension.as_slice()),
    )?;
    validate_patched_axis_line_visibility(&patched, axis, visible)?;
    Ok(patched)
}

fn axis_line_visible_field(axis: ChartAxis) -> u32 {
    match axis {
        ChartAxis::Category => CATEGORY_AXIS_LINE_VISIBLE_FIELD,
        ChartAxis::Value => VALUE_AXIS_LINE_VISIBLE_FIELD,
    }
}

fn validate_patched_axis_line_visibility(
    data: &[u8],
    axis: ChartAxis,
    expected: bool,
) -> Result<()> {
    if read_axis_style_line_visibility(data, axis)? != expected {
        return Err(Error::InvalidFormat(format!(
            "{}-axis line wire patch failed validation",
            axis.label()
        )));
    }
    Ok(())
}

fn generated_axis_style_extension(data: &[u8]) -> Result<Option<&[u8]>> {
    tsch::ChartAxisStyleArchive::decode(data)?;
    let fields = parse_wire_fields(data)?;
    let mut extensions = fields
        .iter()
        .filter(|field| field.number == GENERATED_CHART_AXIS_STYLE_EXTENSION_FIELD);
    let Some(extension) = extensions.next() else {
        return Ok(None);
    };
    if extensions.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "chart axis-style extension {GENERATED_CHART_AXIS_STYLE_EXTENSION_FIELD} occurs more than once"
        )));
    }
    if extension.wire_type != 2 {
        return Err(Error::InvalidFormat(format!(
            "chart axis-style extension {GENERATED_CHART_AXIS_STYLE_EXTENSION_FIELD} is not length-delimited"
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
    fn category_axis_line_patch_retains_value_line_and_unmapped_fields() {
        let generated = tsch::generated::ChartAxisStyleArchive {
            tschchartaxiscategoryshowaxis: Some(false),
            tschchartaxisvalueshowaxis: Some(true),
            ..Default::default()
        };
        let original = axis_style_with_unknown_fields(generated);

        let visible =
            patch_axis_style_line_visibility(&original, ChartAxis::Category, true).unwrap();
        assert!(read_axis_style_line_visibility(&visible, ChartAxis::Category).unwrap());
        assert!(read_axis_style_line_visibility(&visible, ChartAxis::Value).unwrap());
        assert_unknown_fields_retained(&original, &visible);

        let restored =
            patch_axis_style_line_visibility(&visible, ChartAxis::Category, false).unwrap();
        assert_eq!(restored, original);
    }

    #[test]
    fn value_axis_line_patch_retains_category_line_and_unmapped_fields() {
        let generated = tsch::generated::ChartAxisStyleArchive {
            tschchartaxiscategoryshowaxis: Some(true),
            tschchartaxisvalueshowaxis: Some(false),
            ..Default::default()
        };
        let original = axis_style_with_unknown_fields(generated);

        let visible = patch_axis_style_line_visibility(&original, ChartAxis::Value, true).unwrap();
        assert!(read_axis_style_line_visibility(&visible, ChartAxis::Category).unwrap());
        assert!(read_axis_style_line_visibility(&visible, ChartAxis::Value).unwrap());
        assert_unknown_fields_retained(&original, &visible);

        let restored = patch_axis_style_line_visibility(&visible, ChartAxis::Value, false).unwrap();
        assert_eq!(restored, original);
    }

    #[test]
    fn axis_line_patch_creates_a_style_extension_when_missing() {
        let original = tsch::ChartAxisStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();

        let visible = patch_axis_style_line_visibility(&original, ChartAxis::Value, true).unwrap();
        assert!(read_axis_style_line_visibility(&visible, ChartAxis::Value).unwrap());
        assert!(!read_axis_style_line_visibility(&visible, ChartAxis::Category).unwrap());
    }

    fn axis_style_with_unknown_fields(
        generated: tsch::generated::ChartAxisStyleArchive,
    ) -> Vec<u8> {
        let mut extension = generated.encode_to_vec();
        append_varint_field(&mut extension, UNMAPPED_GENERATED_FIELD, UNMAPPED_VALUE).unwrap();
        let base = tsch::ChartAxisStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        };
        let mut data = base.encode_to_vec();
        append_length_delimited_field(
            &mut data,
            GENERATED_CHART_AXIS_STYLE_EXTENSION_FIELD,
            &extension,
        )
        .unwrap();
        append_varint_field(&mut data, UNMAPPED_OUTER_FIELD, UNMAPPED_VALUE).unwrap();
        data
    }

    fn assert_unknown_fields_retained(original: &[u8], patched: &[u8]) {
        assert_eq!(
            raw_field(patched, UNMAPPED_OUTER_FIELD),
            raw_field(original, UNMAPPED_OUTER_FIELD)
        );
        assert_eq!(
            raw_field(
                generated_axis_style_extension(patched).unwrap().unwrap(),
                UNMAPPED_GENERATED_FIELD
            ),
            raw_field(
                generated_axis_style_extension(original).unwrap().unwrap(),
                UNMAPPED_GENERATED_FIELD
            )
        );
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
