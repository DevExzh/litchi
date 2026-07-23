//! Lossless native chart-axis title storage and mutation.
//!
//! iWork stores category- and value-axis titles in the generated extension of
//! a chart's `TSCH.ChartAxisNonStyleArchive`. This module identifies the
//! primary native axis object, preserves both protobuf layers losslessly, and
//! changes only the requested title fields.

use prost::Message;

use crate::archive::RawMessage;
use crate::charts::IWorkChartArchive;
use crate::charts::source::{AXIS_NON_STYLE_MESSAGE_TYPE, CHART_MESSAGE_TYPE};
use crate::charts::unique_chart_object_archive_name;
use crate::protobuf::tsch;
use crate::wire::{parse_wire_fields, patch_length_delimited_field, patch_varint_field};
use crate::{Error, IWorkPackage, Result};

/// Proto2 extension holding the generated chart-axis non-style properties.
const GENERATED_CHART_AXIS_NON_STYLE_EXTENSION_FIELD: u32 = 10_000;
/// `tschchartaxiscategoryshowtitle` in `TSCH.Generated.ChartAxisNonStyleArchive`.
const CATEGORY_AXIS_TITLE_VISIBLE_FIELD: u32 = 13;
/// `tschchartaxisvalueshowtitle` in `TSCH.Generated.ChartAxisNonStyleArchive`.
const VALUE_AXIS_TITLE_VISIBLE_FIELD: u32 = 14;
/// `tschchartaxiscategorytitle` in `TSCH.Generated.ChartAxisNonStyleArchive`.
const CATEGORY_AXIS_TITLE_TEXT_FIELD: u32 = 15;
/// `tschchartaxisvaluetitle` in `TSCH.Generated.ChartAxisNonStyleArchive`.
const VALUE_AXIS_TITLE_TEXT_FIELD: u32 = 16;

/// A native chart axis exposed by iWork's Axis formatter.
///
/// [`Self::Value`] addresses the primary value-axis object. iWork charts can
/// retain additional value-axis objects for specialized chart types, but the
/// standard Axis formatter exposes this primary axis as `Value (Y)`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ChartAxis {
    /// The chart's category axis, shown as `Category (X)` by the standard formatter.
    Category,
    /// The chart's primary value axis, shown as `Value (Y)` by the standard formatter.
    Value,
}

#[derive(Debug, Clone, Copy)]
struct AxisTitleFields {
    visible: u32,
    text: u32,
}

impl ChartAxis {
    const fn title_fields(self) -> AxisTitleFields {
        match self {
            Self::Category => AxisTitleFields {
                visible: CATEGORY_AXIS_TITLE_VISIBLE_FIELD,
                text: CATEGORY_AXIS_TITLE_TEXT_FIELD,
            },
            Self::Value => AxisTitleFields {
                visible: VALUE_AXIS_TITLE_VISIBLE_FIELD,
                text: VALUE_AXIS_TITLE_TEXT_FIELD,
            },
        }
    }

    pub(crate) const fn label(self) -> &'static str {
        match self {
            Self::Category => "category",
            Self::Value => "value",
        }
    }

    fn primary_non_style_identifier(self, chart: &tsch::ChartArchive) -> Option<u64> {
        let references = match self {
            Self::Category => &chart.category_axis_nonstyles,
            Self::Value => &chart.value_axis_nonstyles,
        };
        references
            .first()
            .map(|reference| reference.identifier)
            .filter(|identifier| *identifier != 0)
    }

    pub(crate) fn primary_style_identifier(self, chart: &tsch::ChartArchive) -> Option<u64> {
        let references = match self {
            Self::Category => &chart.category_axis_styles,
            Self::Value => &chart.value_axis_styles,
        };
        references
            .first()
            .map(|reference| reference.identifier)
            .filter(|identifier| *identifier != 0)
    }
}

/// The single mutable native axis non-style payload for one chart axis.
#[derive(Debug)]
struct AxisNonStyleSlot {
    archive_name: String,
    object_id: u64,
    message_index: usize,
}

/// Read one native chart-axis title.
pub(crate) fn chart_axis_title(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    axis: ChartAxis,
) -> Result<Option<String>> {
    axis_non_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        axis,
    )?
    .read(package, |data| read_axis_non_style_title(data, axis))
}

/// Set one native chart-axis title and enable its Axis Name switch.
pub(crate) fn set_chart_axis_title(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    axis: ChartAxis,
    title: &str,
) -> Result<()> {
    let slot = axis_non_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        axis,
    )?;
    if slot
        .read(package, |data| read_axis_non_style_title(data, axis))?
        .as_deref()
        == Some(title)
    {
        return Ok(());
    }
    slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
    slot.update(package, |data| {
        patch_axis_non_style_title(data, axis, Some(title))
    })?;
    if slot
        .read(package, |data| read_axis_non_style_title(data, axis))?
        .as_deref()
        != Some(title)
    {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} {}-axis title update failed validation",
            axis.label()
        )));
    }
    Ok(())
}

/// Remove one visible native chart-axis title.
///
/// Returns whether the title was visible. An axis non-style shared by another
/// chart axis is rejected rather than silently changing that axis.
pub(crate) fn remove_chart_axis_title(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    axis: ChartAxis,
) -> Result<bool> {
    let slot = axis_non_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        axis,
    )?;
    if slot
        .read(package, |data| read_axis_non_style_title(data, axis))?
        .is_none()
    {
        return Ok(false);
    }
    slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
    slot.update(package, |data| patch_axis_non_style_title(data, axis, None))?;
    if slot
        .read(package, |data| read_axis_non_style_title(data, axis))?
        .is_some()
    {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} {}-axis title removal failed validation",
            axis.label()
        )));
    }
    Ok(true)
}

/// Decode a `TSCH.ChartAxisNonStyleArchive` and return one visible title.
fn read_axis_non_style_title(data: &[u8], axis: ChartAxis) -> Result<Option<String>> {
    let Some(extension) = generated_axis_non_style_extension(data)? else {
        return Ok(None);
    };
    let generated = tsch::generated::ChartAxisNonStyleArchive::decode(extension)?;
    let (visible, title) = match axis {
        ChartAxis::Category => (
            generated.tschchartaxiscategoryshowtitle,
            generated.tschchartaxiscategorytitle,
        ),
        ChartAxis::Value => (
            generated.tschchartaxisvalueshowtitle,
            generated.tschchartaxisvaluetitle,
        ),
    };
    if visible != Some(true) {
        return Ok(None);
    }
    Ok(Some(title.unwrap_or_default()))
}

fn axis_non_style_slot(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    axis: ChartAxis,
) -> Result<AxisNonStyleSlot> {
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
    let non_style_id = axis.primary_non_style_identifier(payload).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} has no primary {}-axis non-style",
            axis.label()
        ))
    })?;
    let archive_name =
        unique_chart_object_archive_name(package, non_style_id, "chart axis non-style object")?;
    let archive = package.archive(&archive_name)?;
    let non_style_object = archive.object(non_style_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "{drawable_label} chart {}-axis non-style {non_style_id} is missing",
            axis.label()
        ))
    })?;
    let mut messages = non_style_object
        .messages
        .iter()
        .enumerate()
        .filter(|(_, message)| message.type_ == AXIS_NON_STYLE_MESSAGE_TYPE);
    let Some((message_index, _)) = messages.next() else {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {}-axis non-style {non_style_id} must have exactly one axis non-style payload",
            axis.label()
        )));
    };
    if messages.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {}-axis non-style {non_style_id} must have exactly one axis non-style payload",
            axis.label()
        )));
    }
    Ok(AxisNonStyleSlot {
        archive_name,
        object_id: non_style_id,
        message_index,
    })
}

impl AxisNonStyleSlot {
    fn read<T>(&self, package: &IWorkPackage, read: impl FnOnce(&[u8]) -> Result<T>) -> Result<T> {
        let archive = package.archive(&self.archive_name)?;
        let object = archive.object(self.object_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "chart axis non-style {} is missing",
                self.object_id
            ))
        })?;
        let message = object.messages.get(self.message_index).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "chart axis non-style {} message index changed unexpectedly",
                self.object_id
            ))
        })?;
        if message.type_ != AXIS_NON_STYLE_MESSAGE_TYPE {
            return Err(Error::InvalidFormat(format!(
                "chart axis non-style {} message type changed unexpectedly",
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
                        .value_axis_nonstyles
                        .iter()
                        .chain(payload.category_axis_nonstyles.iter());
                    for reference in references {
                        if reference.identifier == self.object_id {
                            reference_count = reference_count.checked_add(1).ok_or_else(|| {
                                Error::InvalidFormat(
                                    "chart axis non-style reference count overflow".to_owned(),
                                )
                            })?;
                        }
                    }
                }
            }
        }
        if reference_count != 1 {
            return Err(Error::InvalidFormat(format!(
                "{drawable_label} chart {drawable_object_id} axis non-style {} is referenced by {reference_count} chart axes",
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
                Error::InvalidFormat(format!(
                    "chart axis non-style {} is missing",
                    self.object_id
                ))
            })?;
            let original = object.messages.get(self.message_index).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "chart axis non-style {} message index changed unexpectedly",
                    self.object_id
                ))
            })?;
            if original.type_ != AXIS_NON_STYLE_MESSAGE_TYPE {
                return Err(Error::InvalidFormat(format!(
                    "chart axis non-style {} message type changed unexpectedly",
                    self.object_id
                )));
            }
            let data = patch(original.data.as_slice())?;
            object.replace_message(
                self.message_index,
                RawMessage {
                    type_: AXIS_NON_STYLE_MESSAGE_TYPE,
                    data,
                },
            )?;
            Ok(())
        })
    }
}

fn patch_axis_non_style_title(
    data: &[u8],
    axis: ChartAxis,
    title: Option<&str>,
) -> Result<Vec<u8>> {
    let fields = axis.title_fields();
    let Some(extension) = generated_axis_non_style_extension(data)? else {
        let Some(title) = title else {
            return Ok(data.to_vec());
        };
        let mut generated = tsch::generated::ChartAxisNonStyleArchive::default();
        match axis {
            ChartAxis::Category => {
                generated.tschchartaxiscategoryshowtitle = Some(true);
                generated.tschchartaxiscategorytitle = Some(title.to_owned());
            },
            ChartAxis::Value => {
                generated.tschchartaxisvalueshowtitle = Some(true);
                generated.tschchartaxisvaluetitle = Some(title.to_owned());
            },
        }
        let patched = patch_length_delimited_field(
            data,
            GENERATED_CHART_AXIS_NON_STYLE_EXTENSION_FIELD,
            false,
            Some(generated.encode_to_vec().as_slice()),
        )?;
        validate_patched_axis_title(&patched, axis, Some(title))?;
        return Ok(patched);
    };

    let generated = tsch::generated::ChartAxisNonStyleArchive::decode(extension)?;
    let (visible_present, title_present) = match axis {
        ChartAxis::Category => (
            generated.tschchartaxiscategoryshowtitle.is_some(),
            generated.tschchartaxiscategorytitle.is_some(),
        ),
        ChartAxis::Value => (
            generated.tschchartaxisvalueshowtitle.is_some(),
            generated.tschchartaxisvaluetitle.is_some(),
        ),
    };
    let extension = patch_varint_field(
        extension,
        fields.visible,
        visible_present,
        Some(u64::from(title.is_some())),
    )?;
    let extension = patch_length_delimited_field(
        &extension,
        fields.text,
        title_present,
        title.map(str::as_bytes),
    )?;
    let patched = patch_length_delimited_field(
        data,
        GENERATED_CHART_AXIS_NON_STYLE_EXTENSION_FIELD,
        true,
        Some(extension.as_slice()),
    )?;
    validate_patched_axis_title(&patched, axis, title)?;
    Ok(patched)
}

fn validate_patched_axis_title(data: &[u8], axis: ChartAxis, expected: Option<&str>) -> Result<()> {
    if read_axis_non_style_title(data, axis)?.as_deref() != expected {
        return Err(Error::InvalidFormat(format!(
            "{}-axis title wire patch failed validation",
            axis.label()
        )));
    }
    Ok(())
}

fn generated_axis_non_style_extension(data: &[u8]) -> Result<Option<&[u8]>> {
    tsch::ChartAxisNonStyleArchive::decode(data)?;
    let fields = parse_wire_fields(data)?;
    let mut extensions = fields
        .iter()
        .filter(|field| field.number == GENERATED_CHART_AXIS_NON_STYLE_EXTENSION_FIELD);
    let Some(extension) = extensions.next() else {
        return Ok(None);
    };
    if extensions.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "chart axis non-style extension {GENERATED_CHART_AXIS_NON_STYLE_EXTENSION_FIELD} occurs more than once"
        )));
    }
    if extension.wire_type != 2 {
        return Err(Error::InvalidFormat(format!(
            "chart axis non-style extension {GENERATED_CHART_AXIS_NON_STYLE_EXTENSION_FIELD} is not length-delimited"
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
    fn category_title_patch_retains_value_title_and_unmapped_fields() {
        let generated = tsch::generated::ChartAxisNonStyleArchive {
            tschchartaxiscategoryshowtitle: Some(false),
            tschchartaxisvalueshowtitle: Some(true),
            tschchartaxisvaluetitle: Some("Revenue".to_owned()),
            ..Default::default()
        };
        let original = axis_non_style_with_unknown_fields(generated);

        let titled =
            patch_axis_non_style_title(&original, ChartAxis::Category, Some("Month")).unwrap();
        assert_eq!(
            read_axis_non_style_title(&titled, ChartAxis::Category).unwrap(),
            Some("Month".to_owned())
        );
        assert_eq!(
            read_axis_non_style_title(&titled, ChartAxis::Value).unwrap(),
            Some("Revenue".to_owned())
        );
        assert_unknown_fields_retained(&original, &titled);

        let removed = patch_axis_non_style_title(&titled, ChartAxis::Category, None).unwrap();
        assert_eq!(
            read_axis_non_style_title(&removed, ChartAxis::Category).unwrap(),
            None
        );
        assert_eq!(removed, original);
    }

    #[test]
    fn value_title_patch_retains_category_title_and_unmapped_fields() {
        let generated = tsch::generated::ChartAxisNonStyleArchive {
            tschchartaxiscategoryshowtitle: Some(true),
            tschchartaxiscategorytitle: Some("Month".to_owned()),
            tschchartaxisvalueshowtitle: Some(false),
            ..Default::default()
        };
        let original = axis_non_style_with_unknown_fields(generated);

        let titled =
            patch_axis_non_style_title(&original, ChartAxis::Value, Some("Revenue")).unwrap();
        assert_eq!(
            read_axis_non_style_title(&titled, ChartAxis::Category).unwrap(),
            Some("Month".to_owned())
        );
        assert_eq!(
            read_axis_non_style_title(&titled, ChartAxis::Value).unwrap(),
            Some("Revenue".to_owned())
        );
        assert_unknown_fields_retained(&original, &titled);

        let removed = patch_axis_non_style_title(&titled, ChartAxis::Value, None).unwrap();
        assert_eq!(
            read_axis_non_style_title(&removed, ChartAxis::Value).unwrap(),
            None
        );
        assert_eq!(removed, original);
    }

    #[test]
    fn title_patch_creates_an_axis_extension_when_missing() {
        let original = tsch::ChartAxisNonStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();

        let titled =
            patch_axis_non_style_title(&original, ChartAxis::Value, Some("Revenue")).unwrap();
        assert_eq!(
            read_axis_non_style_title(&titled, ChartAxis::Value).unwrap(),
            Some("Revenue".to_owned())
        );
        assert_eq!(
            read_axis_non_style_title(&titled, ChartAxis::Category).unwrap(),
            None
        );
    }

    fn axis_non_style_with_unknown_fields(
        generated: tsch::generated::ChartAxisNonStyleArchive,
    ) -> Vec<u8> {
        let mut extension = generated.encode_to_vec();
        append_varint_field(&mut extension, UNMAPPED_GENERATED_FIELD, UNMAPPED_VALUE).unwrap();
        let base = tsch::ChartAxisNonStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        };
        let mut data = base.encode_to_vec();
        append_length_delimited_field(
            &mut data,
            GENERATED_CHART_AXIS_NON_STYLE_EXTENSION_FIELD,
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
                generated_axis_non_style_extension(patched)
                    .unwrap()
                    .unwrap(),
                UNMAPPED_GENERATED_FIELD
            ),
            raw_field(
                generated_axis_non_style_extension(original)
                    .unwrap()
                    .unwrap(),
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
