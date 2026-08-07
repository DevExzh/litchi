//! Lossless native chart-axis non-style storage and mutation.
//!
//! iWork stores title, label, and scale properties in the generated extension
//! of a chart's `TSCH.ChartAxisNonStyleArchive`. This module owns the common
//! native object lookup and lossless wire-level access used by focused
//! axis-property modules.

use litchi_iwa_common::chart::axis::Axis;
use prost::Message;

use crate::archive::RawMessage;
use crate::charts::IWorkChartArchive;
use crate::charts::source::{AXIS_NON_STYLE_MESSAGE_TYPE, CHART_MESSAGE_TYPE};
use crate::charts::unique_chart_object_archive_name;
use crate::protobuf::tsch;
use crate::wire::{parse_wire_fields, patch_length_delimited_field, patch_varint_field};
use crate::{Error, IWorkPackage, Result};

/// Proto2 extension holding the generated chart-axis non-style properties.
pub(crate) const GENERATED_CHART_AXIS_NON_STYLE_EXTENSION_FIELD: u32 = 10_000;
/// `tschchartaxiscategoryshowtitle` in `TSCH.Generated.ChartAxisNonStyleArchive`.
const CATEGORY_AXIS_TITLE_VISIBLE_FIELD: u32 = 13;
/// `tschchartaxisvalueshowtitle` in `TSCH.Generated.ChartAxisNonStyleArchive`.
const VALUE_AXIS_TITLE_VISIBLE_FIELD: u32 = 14;
/// `tschchartaxiscategorytitle` in `TSCH.Generated.ChartAxisNonStyleArchive`.
const CATEGORY_AXIS_TITLE_TEXT_FIELD: u32 = 15;
/// `tschchartaxisvaluetitle` in `TSCH.Generated.ChartAxisNonStyleArchive`.
const VALUE_AXIS_TITLE_TEXT_FIELD: u32 = 16;
/// `tschchartaxiscategoryshowlabels` in `TSCH.Generated.ChartAxisNonStyleArchive`.
const CATEGORY_AXIS_LABELS_VISIBLE_FIELD: u32 = 9;
/// `tschchartaxisvalueshowlabels` in `TSCH.Generated.ChartAxisNonStyleArchive`.
const VALUE_AXIS_LABELS_VISIBLE_FIELD: u32 = 11;
/// `tschchartaxiscategoryshowserieslabels` in
/// `TSCH.Generated.ChartAxisNonStyleArchive`.
const CATEGORY_AXIS_SERIES_NAMES_VISIBLE_FIELD: u32 = 12;

#[derive(Debug, Clone, Copy)]
struct AxisTitleFields {
    visible: u32,
    text: u32,
}

const fn title_fields(axis: Axis) -> AxisTitleFields {
    match axis {
        Axis::Category => AxisTitleFields {
            visible: CATEGORY_AXIS_TITLE_VISIBLE_FIELD,
            text: CATEGORY_AXIS_TITLE_TEXT_FIELD,
        },
        Axis::Value => AxisTitleFields {
            visible: VALUE_AXIS_TITLE_VISIBLE_FIELD,
            text: VALUE_AXIS_TITLE_TEXT_FIELD,
        },
    }
}

const fn labels_visible_field(axis: Axis) -> u32 {
    match axis {
        Axis::Category => CATEGORY_AXIS_LABELS_VISIBLE_FIELD,
        Axis::Value => VALUE_AXIS_LABELS_VISIBLE_FIELD,
    }
}

fn primary_non_style_identifier(axis: Axis, chart: &tsch::ChartArchive) -> Option<u64> {
    let references = match axis {
        Axis::Category => &chart.category_axis_nonstyles,
        Axis::Value => &chart.value_axis_nonstyles,
    };
    references
        .first()
        .map(|reference| reference.identifier)
        .filter(|identifier| *identifier != 0)
}

pub(crate) fn primary_style_identifier(axis: Axis, chart: &tsch::ChartArchive) -> Option<u64> {
    let references = match axis {
        Axis::Category => &chart.category_axis_styles,
        Axis::Value => &chart.value_axis_styles,
    };
    references
        .first()
        .map(|reference| reference.identifier)
        .filter(|identifier| *identifier != 0)
}

/// The single mutable native axis non-style payload for one chart axis.
#[derive(Debug)]
pub(crate) struct AxisNonStyleSlot {
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
    axis: Axis,
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
    axis: Axis,
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
            axis.as_str()
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
    axis: Axis,
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
            axis.as_str()
        )));
    }
    Ok(true)
}

/// Read whether iWork shows labels for one native chart axis.
pub(crate) fn chart_axis_labels_visible(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    axis: Axis,
) -> Result<bool> {
    axis_non_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        axis,
    )?
    .read(package, |data| read_axis_labels_visible(data, axis))
}

/// Set whether iWork shows labels for one native chart axis.
pub(crate) fn set_chart_axis_labels_visible(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    axis: Axis,
    visible: bool,
) -> Result<()> {
    let slot = axis_non_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        axis,
    )?;
    if slot.read(package, |data| read_axis_labels_visible(data, axis))? == visible {
        return Ok(());
    }
    slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
    slot.update(package, |data| {
        patch_axis_labels_visibility(data, axis, visible)
    })?;
    if slot.read(package, |data| read_axis_labels_visible(data, axis))? != visible {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} {}-axis labels update failed validation",
            axis.as_str()
        )));
    }
    Ok(())
}

/// Read whether iWork shows series names on a native chart category axis.
pub(crate) fn chart_category_axis_series_names_visible(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<bool> {
    axis_non_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        Axis::Category,
    )?
    .read(package, read_category_axis_series_names_visible)
}

/// Set whether iWork shows series names on a native chart category axis.
pub(crate) fn set_chart_category_axis_series_names_visible(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    visible: bool,
) -> Result<()> {
    let slot = axis_non_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        Axis::Category,
    )?;
    if slot.read(package, read_category_axis_series_names_visible)? == visible {
        return Ok(());
    }
    slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
    slot.update(package, |data| {
        patch_category_axis_series_names_visibility(data, visible)
    })?;
    if slot.read(package, read_category_axis_series_names_visible)? != visible {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} category-axis series-names update failed validation"
        )));
    }
    Ok(())
}

/// Decode a `TSCH.ChartAxisNonStyleArchive` and return one visible title.
fn read_axis_non_style_title(data: &[u8], axis: Axis) -> Result<Option<String>> {
    let Some(extension) = generated_axis_non_style_extension(data)? else {
        return Ok(None);
    };
    let generated = tsch::generated::ChartAxisNonStyleArchive::decode(extension)?;
    let (visible, title) = match axis {
        Axis::Category => (
            generated.tschchartaxiscategoryshowtitle,
            generated.tschchartaxiscategorytitle,
        ),
        Axis::Value => (
            generated.tschchartaxisvalueshowtitle,
            generated.tschchartaxisvaluetitle,
        ),
    };
    if visible != Some(true) {
        return Ok(None);
    }
    Ok(Some(title.unwrap_or_default()))
}

/// Decode a `TSCH.ChartAxisNonStyleArchive` label-visibility switch.
///
/// iWork shows axis labels by default when the native switch is absent.
fn read_axis_labels_visible(data: &[u8], axis: Axis) -> Result<bool> {
    let Some(extension) = generated_axis_non_style_extension(data)? else {
        return Ok(true);
    };
    let generated = tsch::generated::ChartAxisNonStyleArchive::decode(extension)?;
    let visible = match axis {
        Axis::Category => generated.tschchartaxiscategoryshowlabels,
        Axis::Value => generated.tschchartaxisvalueshowlabels,
    };
    Ok(visible.unwrap_or(true))
}

/// Decode a `TSCH.ChartAxisNonStyleArchive` category-axis series-names switch.
fn read_category_axis_series_names_visible(data: &[u8]) -> Result<bool> {
    let Some(extension) = generated_axis_non_style_extension(data)? else {
        return Ok(false);
    };
    let generated = tsch::generated::ChartAxisNonStyleArchive::decode(extension)?;
    Ok(generated
        .tschchartaxiscategoryshowserieslabels
        .unwrap_or(false))
}

pub(crate) fn axis_non_style_slot(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    axis: Axis,
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
    let non_style_id = primary_non_style_identifier(axis, payload).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} has no primary {}-axis non-style",
            axis.as_str()
        ))
    })?;
    let archive_name =
        unique_chart_object_archive_name(package, non_style_id, "chart axis non-style object")?;
    let archive = package.archive(&archive_name)?;
    let non_style_object = archive.object(non_style_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "{drawable_label} chart {}-axis non-style {non_style_id} is missing",
            axis.as_str()
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
            axis.as_str()
        )));
    };
    if messages.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {}-axis non-style {non_style_id} must have exactly one axis non-style payload",
            axis.as_str()
        )));
    }
    Ok(AxisNonStyleSlot {
        archive_name,
        object_id: non_style_id,
        message_index,
    })
}

impl AxisNonStyleSlot {
    pub(crate) fn read<T>(
        &self,
        package: &IWorkPackage,
        read: impl FnOnce(&[u8]) -> Result<T>,
    ) -> Result<T> {
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

    pub(crate) fn ensure_exclusive(
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

    pub(crate) fn update(
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

fn patch_axis_non_style_title(data: &[u8], axis: Axis, title: Option<&str>) -> Result<Vec<u8>> {
    let fields = title_fields(axis);
    let Some(extension) = generated_axis_non_style_extension(data)? else {
        let Some(title) = title else {
            return Ok(data.to_vec());
        };
        let mut generated = tsch::generated::ChartAxisNonStyleArchive::default();
        match axis {
            Axis::Category => {
                generated.tschchartaxiscategoryshowtitle = Some(true);
                generated.tschchartaxiscategorytitle = Some(title.to_owned());
            },
            Axis::Value => {
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
        Axis::Category => (
            generated.tschchartaxiscategoryshowtitle.is_some(),
            generated.tschchartaxiscategorytitle.is_some(),
        ),
        Axis::Value => (
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

fn patch_axis_labels_visibility(data: &[u8], axis: Axis, visible: bool) -> Result<Vec<u8>> {
    let Some(extension) = generated_axis_non_style_extension(data)? else {
        if visible {
            return Ok(data.to_vec());
        }
        let mut generated = tsch::generated::ChartAxisNonStyleArchive::default();
        match axis {
            Axis::Category => generated.tschchartaxiscategoryshowlabels = Some(false),
            Axis::Value => generated.tschchartaxisvalueshowlabels = Some(false),
        }
        let patched = patch_length_delimited_field(
            data,
            GENERATED_CHART_AXIS_NON_STYLE_EXTENSION_FIELD,
            false,
            Some(generated.encode_to_vec().as_slice()),
        )?;
        validate_patched_axis_labels_visibility(&patched, axis, visible)?;
        return Ok(patched);
    };

    let generated = tsch::generated::ChartAxisNonStyleArchive::decode(extension)?;
    let visible_present = match axis {
        Axis::Category => generated.tschchartaxiscategoryshowlabels.is_some(),
        Axis::Value => generated.tschchartaxisvalueshowlabels.is_some(),
    };
    let extension = patch_varint_field(
        extension,
        labels_visible_field(axis),
        visible_present,
        Some(u64::from(visible)),
    )?;
    let patched = patch_length_delimited_field(
        data,
        GENERATED_CHART_AXIS_NON_STYLE_EXTENSION_FIELD,
        true,
        Some(extension.as_slice()),
    )?;
    validate_patched_axis_labels_visibility(&patched, axis, visible)?;
    Ok(patched)
}

fn patch_category_axis_series_names_visibility(data: &[u8], visible: bool) -> Result<Vec<u8>> {
    let Some(extension) = generated_axis_non_style_extension(data)? else {
        let generated = tsch::generated::ChartAxisNonStyleArchive {
            tschchartaxiscategoryshowserieslabels: Some(visible),
            ..Default::default()
        };
        let patched = patch_length_delimited_field(
            data,
            GENERATED_CHART_AXIS_NON_STYLE_EXTENSION_FIELD,
            false,
            Some(generated.encode_to_vec().as_slice()),
        )?;
        validate_patched_category_axis_series_names_visibility(&patched, visible)?;
        return Ok(patched);
    };

    let generated = tsch::generated::ChartAxisNonStyleArchive::decode(extension)?;
    let extension = patch_varint_field(
        extension,
        CATEGORY_AXIS_SERIES_NAMES_VISIBLE_FIELD,
        generated.tschchartaxiscategoryshowserieslabels.is_some(),
        Some(u64::from(visible)),
    )?;
    let patched = patch_length_delimited_field(
        data,
        GENERATED_CHART_AXIS_NON_STYLE_EXTENSION_FIELD,
        true,
        Some(extension.as_slice()),
    )?;
    validate_patched_category_axis_series_names_visibility(&patched, visible)?;
    Ok(patched)
}

fn validate_patched_axis_title(data: &[u8], axis: Axis, expected: Option<&str>) -> Result<()> {
    if read_axis_non_style_title(data, axis)?.as_deref() != expected {
        return Err(Error::InvalidFormat(format!(
            "{}-axis title wire patch failed validation",
            axis.as_str()
        )));
    }
    Ok(())
}

fn validate_patched_axis_labels_visibility(data: &[u8], axis: Axis, expected: bool) -> Result<()> {
    if read_axis_labels_visible(data, axis)? != expected {
        return Err(Error::InvalidFormat(format!(
            "{}-axis labels wire patch failed validation",
            axis.as_str()
        )));
    }
    Ok(())
}

fn validate_patched_category_axis_series_names_visibility(
    data: &[u8],
    expected: bool,
) -> Result<()> {
    if read_category_axis_series_names_visible(data)? != expected {
        return Err(Error::InvalidFormat(
            "category-axis series-names wire patch failed validation".to_owned(),
        ));
    }
    Ok(())
}

pub(crate) fn generated_axis_non_style_extension(data: &[u8]) -> Result<Option<&[u8]>> {
    tsch::ChartAxisNonStyleArchive::decode(data)?;
    let fields = parse_wire_fields(data)?;
    let mut extensions = fields
        .iter()
        .filter(|field| field.number() == GENERATED_CHART_AXIS_NON_STYLE_EXTENSION_FIELD);
    let Some(extension) = extensions.next() else {
        return Ok(None);
    };
    if extensions.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "chart axis non-style extension {GENERATED_CHART_AXIS_NON_STYLE_EXTENSION_FIELD} occurs more than once"
        )));
    }
    if extension.wire_type() != 2 {
        return Err(Error::InvalidFormat(format!(
            "chart axis non-style extension {GENERATED_CHART_AXIS_NON_STYLE_EXTENSION_FIELD} is not length-delimited"
        )));
    }
    Ok(Some(&data[extension.payload_start()..extension.end()]))
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

        let titled = patch_axis_non_style_title(&original, Axis::Category, Some("Month")).unwrap();
        assert_eq!(
            read_axis_non_style_title(&titled, Axis::Category).unwrap(),
            Some("Month".to_owned())
        );
        assert_eq!(
            read_axis_non_style_title(&titled, Axis::Value).unwrap(),
            Some("Revenue".to_owned())
        );
        assert_unknown_fields_retained(&original, &titled);

        let removed = patch_axis_non_style_title(&titled, Axis::Category, None).unwrap();
        assert_eq!(
            read_axis_non_style_title(&removed, Axis::Category).unwrap(),
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

        let titled = patch_axis_non_style_title(&original, Axis::Value, Some("Revenue")).unwrap();
        assert_eq!(
            read_axis_non_style_title(&titled, Axis::Category).unwrap(),
            Some("Month".to_owned())
        );
        assert_eq!(
            read_axis_non_style_title(&titled, Axis::Value).unwrap(),
            Some("Revenue".to_owned())
        );
        assert_unknown_fields_retained(&original, &titled);

        let removed = patch_axis_non_style_title(&titled, Axis::Value, None).unwrap();
        assert_eq!(
            read_axis_non_style_title(&removed, Axis::Value).unwrap(),
            None
        );
        assert_eq!(removed, original);
    }

    #[test]
    fn category_axis_series_names_patch_retains_titles_and_unmapped_fields() {
        let generated = tsch::generated::ChartAxisNonStyleArchive {
            tschchartaxiscategoryshowtitle: Some(true),
            tschchartaxiscategorytitle: Some("Month".to_owned()),
            tschchartaxisvalueshowtitle: Some(true),
            tschchartaxisvaluetitle: Some("Revenue".to_owned()),
            tschchartaxiscategoryshowserieslabels: Some(false),
            ..Default::default()
        };
        let original = axis_non_style_with_unknown_fields(generated);

        let visible = patch_category_axis_series_names_visibility(&original, true).unwrap();
        assert!(read_category_axis_series_names_visible(&visible).unwrap());
        assert_eq!(
            read_axis_non_style_title(&visible, Axis::Category).unwrap(),
            Some("Month".to_owned())
        );
        assert_eq!(
            read_axis_non_style_title(&visible, Axis::Value).unwrap(),
            Some("Revenue".to_owned())
        );
        assert_unknown_fields_retained(&original, &visible);

        let restored = patch_category_axis_series_names_visibility(&visible, false).unwrap();
        assert_eq!(restored, original);
    }

    #[test]
    fn category_axis_series_names_defaults_hidden_and_creates_an_extension() {
        let original = tsch::ChartAxisNonStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();
        assert!(!read_category_axis_series_names_visible(&original).unwrap());

        let visible = patch_category_axis_series_names_visibility(&original, true).unwrap();
        assert!(read_category_axis_series_names_visible(&visible).unwrap());
        assert!(
            generated_axis_non_style_extension(&visible)
                .unwrap()
                .is_some()
        );
    }

    #[test]
    fn axis_labels_patch_retains_titles_and_unmapped_fields() {
        let generated = tsch::generated::ChartAxisNonStyleArchive {
            tschchartaxiscategoryshowlabels: Some(true),
            tschchartaxisvalueshowlabels: Some(true),
            tschchartaxiscategoryshowtitle: Some(true),
            tschchartaxiscategorytitle: Some("Month".to_owned()),
            tschchartaxisvalueshowtitle: Some(true),
            tschchartaxisvaluetitle: Some("Revenue".to_owned()),
            tschchartaxiscategoryshowserieslabels: Some(true),
            ..Default::default()
        };
        let original = axis_non_style_with_unknown_fields(generated);

        let category_hidden =
            patch_axis_labels_visibility(&original, Axis::Category, false).unwrap();
        assert!(!read_axis_labels_visible(&category_hidden, Axis::Category).unwrap());
        assert!(read_axis_labels_visible(&category_hidden, Axis::Value).unwrap());
        assert!(read_category_axis_series_names_visible(&category_hidden).unwrap());
        assert_eq!(
            read_axis_non_style_title(&category_hidden, Axis::Category).unwrap(),
            Some("Month".to_owned())
        );
        assert_unknown_fields_retained(&original, &category_hidden);

        let value_hidden =
            patch_axis_labels_visibility(&category_hidden, Axis::Value, false).unwrap();
        assert!(!read_axis_labels_visible(&value_hidden, Axis::Category).unwrap());
        assert!(!read_axis_labels_visible(&value_hidden, Axis::Value).unwrap());
        assert_eq!(
            read_axis_non_style_title(&value_hidden, Axis::Value).unwrap(),
            Some("Revenue".to_owned())
        );
        assert_unknown_fields_retained(&original, &value_hidden);
    }

    #[test]
    fn axis_labels_default_visible_and_hidden_setting_creates_an_extension() {
        let original = tsch::ChartAxisNonStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();
        for axis in [Axis::Category, Axis::Value] {
            assert!(read_axis_labels_visible(&original, axis).unwrap());
            assert_eq!(
                patch_axis_labels_visibility(&original, axis, true).unwrap(),
                original
            );
            let hidden = patch_axis_labels_visibility(&original, axis, false).unwrap();
            assert!(!read_axis_labels_visible(&hidden, axis).unwrap());
            assert!(
                generated_axis_non_style_extension(&hidden)
                    .unwrap()
                    .is_some()
            );
        }
    }

    #[test]
    fn title_patch_creates_an_axis_extension_when_missing() {
        let original = tsch::ChartAxisNonStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();

        let titled = patch_axis_non_style_title(&original, Axis::Value, Some("Revenue")).unwrap();
        assert_eq!(
            read_axis_non_style_title(&titled, Axis::Value).unwrap(),
            Some("Revenue".to_owned())
        );
        assert_eq!(
            read_axis_non_style_title(&titled, Axis::Category).unwrap(),
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
            .filter(|field| field.number() == number)
            .map(|field| data[field.start()..field.end()].to_vec())
            .collect()
    }
}
