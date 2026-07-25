//! Lossless native chart-axis style storage and mutation.
//!
//! iWork stores axis-line, gridline, tick-mark visibility and location, and
//! value-axis minimum-label switches in the generated extension of a chart's
//! `TSCH.ChartAxisStyleArchive`. This module identifies the primary native
//! axis-style object, preserves both protobuf layers losslessly, and changes
//! only the requested style switch.

use prost::Message;

use crate::archive::RawMessage;
use crate::charts::ChartAxis;
use crate::charts::IWorkChartArchive;
use crate::charts::source::{AXIS_STYLE_MESSAGE_TYPE, CHART_MESSAGE_TYPE};
use crate::charts::unique_chart_object_archive_name;
use crate::protobuf::tsch;
use crate::wire::{parse_wire_fields, patch_length_delimited_field, patch_varint_field};
use crate::{Error, IWorkPackage, Result};

/// Where iWork draws the major tick marks for one chart axis.
///
/// The known values correspond to the Tick Marks pop-up in Pages, Numbers,
/// and Keynote. `Unsupported` preserves a future native integer without
/// collapsing it into a different location.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
#[non_exhaustive]
pub enum ChartAxisTickMarkLocation {
    /// Draw no major tick marks.
    None,
    /// Draw tick marks toward the chart's plot area.
    Inside,
    /// Draw tick marks centered on the axis line.
    #[default]
    Centered,
    /// Draw tick marks away from the chart's plot area.
    Outside,
    /// Preserve an unrecognized native iWork value.
    Unsupported(i32),
}

impl ChartAxisTickMarkLocation {
    /// Decode the integer stored by the iWork protobuf schema.
    pub const fn from_raw(value: i32) -> Self {
        match value {
            0 => Self::None,
            1 => Self::Inside,
            2 => Self::Centered,
            3 => Self::Outside,
            value => Self::Unsupported(value),
        }
    }

    /// Return the integer used by the iWork protobuf schema.
    pub const fn into_raw(self) -> i32 {
        match self {
            Self::None => 0,
            Self::Inside => 1,
            Self::Centered => 2,
            Self::Outside => 3,
            Self::Unsupported(value) => value,
        }
    }
}

/// Proto2 extension holding the generated chart-axis style properties.
pub(crate) const GENERATED_CHART_AXIS_STYLE_EXTENSION_FIELD: u32 = 10_000;
/// `tschchartaxiscategoryshowaxis` in `TSCH.Generated.ChartAxisStyleArchive`.
const CATEGORY_AXIS_LINE_VISIBLE_FIELD: u32 = 24;
/// `tschchartaxisvalueshowaxis` in `TSCH.Generated.ChartAxisStyleArchive`.
const VALUE_AXIS_LINE_VISIBLE_FIELD: u32 = 25;
/// `tschchartaxiscategoryshowmajorgridlines` in
/// `TSCH.Generated.ChartAxisStyleArchive`.
const CATEGORY_MAJOR_GRIDLINES_VISIBLE_FIELD: u32 = 27;
/// `tschchartaxisvalueshowmajorgridlines` in
/// `TSCH.Generated.ChartAxisStyleArchive`.
const VALUE_MAJOR_GRIDLINES_VISIBLE_FIELD: u32 = 28;
/// `tschchartaxiscategoryshowminorgridlines` in
/// `TSCH.Generated.ChartAxisStyleArchive`.
const CATEGORY_MINOR_GRIDLINES_VISIBLE_FIELD: u32 = 32;
/// `tschchartaxisvalueshowminorgridlines` in
/// `TSCH.Generated.ChartAxisStyleArchive`.
const VALUE_MINOR_GRIDLINES_VISIBLE_FIELD: u32 = 33;
/// `tschchartaxiscategoryshowminortickmarks` in
/// `TSCH.Generated.ChartAxisStyleArchive`.
const CATEGORY_MINOR_TICK_MARKS_VISIBLE_FIELD: u32 = 34;
/// `tschchartaxisvalueshowminortickmarks` in
/// `TSCH.Generated.ChartAxisStyleArchive`.
const VALUE_MINOR_TICK_MARKS_VISIBLE_FIELD: u32 = 35;
/// `tschchartaxiscategorytickmarklocation` in
/// `TSCH.Generated.ChartAxisStyleArchive`.
const CATEGORY_TICK_MARK_LOCATION_FIELD: u32 = 36;
/// `tschchartaxisvaluetickmarklocation` in
/// `TSCH.Generated.ChartAxisStyleArchive`.
const VALUE_TICK_MARK_LOCATION_FIELD: u32 = 37;
/// `tschchartaxisvalueshowminimumlabel` in
/// `TSCH.Generated.ChartAxisStyleArchive`.
const VALUE_AXIS_MINIMUM_LABEL_VISIBLE_FIELD: u32 = 31;

/// One boolean chart-axis style switch with an explicit native field mapping.
#[derive(Debug, Clone, Copy)]
enum AxisStyleSwitch {
    Line,
    MajorGridlines,
    MinorGridlines,
    MinorTickMarks,
    ValueMinimumLabel,
}

impl AxisStyleSwitch {
    fn validate_axis(self, axis: ChartAxis) -> Result<()> {
        if matches!((self, axis), (Self::ValueMinimumLabel, ChartAxis::Category)) {
            return Err(Error::InvalidFormat(
                "chart minimum-value labels are only available on the value axis".to_owned(),
            ));
        }
        Ok(())
    }

    const fn field(self, axis: ChartAxis) -> u32 {
        match (self, axis) {
            (Self::Line, ChartAxis::Category) => CATEGORY_AXIS_LINE_VISIBLE_FIELD,
            (Self::Line, ChartAxis::Value) => VALUE_AXIS_LINE_VISIBLE_FIELD,
            (Self::MajorGridlines, ChartAxis::Category) => CATEGORY_MAJOR_GRIDLINES_VISIBLE_FIELD,
            (Self::MajorGridlines, ChartAxis::Value) => VALUE_MAJOR_GRIDLINES_VISIBLE_FIELD,
            (Self::MinorGridlines, ChartAxis::Category) => CATEGORY_MINOR_GRIDLINES_VISIBLE_FIELD,
            (Self::MinorGridlines, ChartAxis::Value) => VALUE_MINOR_GRIDLINES_VISIBLE_FIELD,
            (Self::MinorTickMarks, ChartAxis::Category) => CATEGORY_MINOR_TICK_MARKS_VISIBLE_FIELD,
            (Self::MinorTickMarks, ChartAxis::Value) => VALUE_MINOR_TICK_MARKS_VISIBLE_FIELD,
            (Self::ValueMinimumLabel, _) => VALUE_AXIS_MINIMUM_LABEL_VISIBLE_FIELD,
        }
    }

    const fn label(self) -> &'static str {
        match self {
            Self::Line => "line",
            Self::MajorGridlines => "major gridlines",
            Self::MinorGridlines => "minor gridlines",
            Self::MinorTickMarks => "minor tick marks",
            Self::ValueMinimumLabel => "minimum value label",
        }
    }

    const fn default_visible(self) -> bool {
        match self {
            Self::ValueMinimumLabel => true,
            Self::MinorTickMarks => true,
            Self::Line | Self::MajorGridlines | Self::MinorGridlines => false,
        }
    }

    fn visible(
        self,
        generated: &tsch::generated::ChartAxisStyleArchive,
        axis: ChartAxis,
    ) -> Option<bool> {
        match (self, axis) {
            (Self::Line, ChartAxis::Category) => generated.tschchartaxiscategoryshowaxis,
            (Self::Line, ChartAxis::Value) => generated.tschchartaxisvalueshowaxis,
            (Self::MajorGridlines, ChartAxis::Category) => {
                generated.tschchartaxiscategoryshowmajorgridlines
            },
            (Self::MajorGridlines, ChartAxis::Value) => {
                generated.tschchartaxisvalueshowmajorgridlines
            },
            (Self::MinorGridlines, ChartAxis::Category) => {
                generated.tschchartaxiscategoryshowminorgridlines
            },
            (Self::MinorGridlines, ChartAxis::Value) => {
                generated.tschchartaxisvalueshowminorgridlines
            },
            (Self::MinorTickMarks, ChartAxis::Category) => {
                generated.tschchartaxiscategoryshowminortickmarks
            },
            (Self::MinorTickMarks, ChartAxis::Value) => {
                generated.tschchartaxisvalueshowminortickmarks
            },
            (Self::ValueMinimumLabel, _) => generated.tschchartaxisvalueshowminimumlabel,
        }
    }

    fn set(
        self,
        generated: &mut tsch::generated::ChartAxisStyleArchive,
        axis: ChartAxis,
        visible: bool,
    ) {
        match (self, axis) {
            (Self::Line, ChartAxis::Category) => {
                generated.tschchartaxiscategoryshowaxis = Some(visible)
            },
            (Self::Line, ChartAxis::Value) => generated.tschchartaxisvalueshowaxis = Some(visible),
            (Self::MajorGridlines, ChartAxis::Category) => {
                generated.tschchartaxiscategoryshowmajorgridlines = Some(visible)
            },
            (Self::MajorGridlines, ChartAxis::Value) => {
                generated.tschchartaxisvalueshowmajorgridlines = Some(visible)
            },
            (Self::MinorGridlines, ChartAxis::Category) => {
                generated.tschchartaxiscategoryshowminorgridlines = Some(visible)
            },
            (Self::MinorGridlines, ChartAxis::Value) => {
                generated.tschchartaxisvalueshowminorgridlines = Some(visible)
            },
            (Self::MinorTickMarks, ChartAxis::Category) => {
                generated.tschchartaxiscategoryshowminortickmarks = Some(visible)
            },
            (Self::MinorTickMarks, ChartAxis::Value) => {
                generated.tschchartaxisvalueshowminortickmarks = Some(visible)
            },
            (Self::ValueMinimumLabel, _) => {
                generated.tschchartaxisvalueshowminimumlabel = Some(visible)
            },
        }
    }
}

/// The single mutable native axis-style payload for one chart axis.
#[derive(Debug)]
pub(crate) struct AxisStyleSlot {
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
    chart_axis_style_switch_visible(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        axis,
        AxisStyleSwitch::Line,
    )
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
    set_chart_axis_style_switch_visible(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        axis,
        AxisStyleSwitch::Line,
        visible,
    )
}

/// Read whether iWork shows major gridlines for one native chart axis.
pub(crate) fn chart_axis_major_gridlines_visible(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    axis: ChartAxis,
) -> Result<bool> {
    chart_axis_style_switch_visible(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        axis,
        AxisStyleSwitch::MajorGridlines,
    )
}

/// Set whether iWork shows major gridlines for one native chart axis.
pub(crate) fn set_chart_axis_major_gridlines_visible(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    axis: ChartAxis,
    visible: bool,
) -> Result<()> {
    set_chart_axis_style_switch_visible(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        axis,
        AxisStyleSwitch::MajorGridlines,
        visible,
    )
}

/// Read whether iWork shows minor gridlines for one native chart axis.
pub(crate) fn chart_axis_minor_gridlines_visible(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    axis: ChartAxis,
) -> Result<bool> {
    chart_axis_style_switch_visible(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        axis,
        AxisStyleSwitch::MinorGridlines,
    )
}

/// Set whether iWork shows minor gridlines for one native chart axis.
pub(crate) fn set_chart_axis_minor_gridlines_visible(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    axis: ChartAxis,
    visible: bool,
) -> Result<()> {
    set_chart_axis_style_switch_visible(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        axis,
        AxisStyleSwitch::MinorGridlines,
        visible,
    )
}

/// Read whether iWork shows minor tick marks for one native chart axis.
pub(crate) fn chart_axis_minor_tick_marks_visible(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    axis: ChartAxis,
) -> Result<bool> {
    chart_axis_style_switch_visible(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        axis,
        AxisStyleSwitch::MinorTickMarks,
    )
}

/// Set whether iWork shows minor tick marks for one native chart axis.
pub(crate) fn set_chart_axis_minor_tick_marks_visible(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    axis: ChartAxis,
    visible: bool,
) -> Result<()> {
    set_chart_axis_style_switch_visible(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        axis,
        AxisStyleSwitch::MinorTickMarks,
        visible,
    )
}

/// Read where iWork draws major tick marks for one native chart axis.
pub(crate) fn chart_axis_tick_mark_location(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    axis: ChartAxis,
) -> Result<ChartAxisTickMarkLocation> {
    axis_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        axis,
    )?
    .read(package, |data| read_axis_tick_mark_location(data, axis))
}

/// Set where iWork draws major tick marks for one native chart axis.
pub(crate) fn set_chart_axis_tick_mark_location(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    axis: ChartAxis,
    location: ChartAxisTickMarkLocation,
) -> Result<()> {
    let slot = axis_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        axis,
    )?;
    if slot.read(package, |data| read_axis_tick_mark_location(data, axis))? == location {
        return Ok(());
    }
    slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
    slot.update(package, |data| {
        patch_axis_tick_mark_location(data, axis, location)
    })?;
    if slot.read(package, |data| read_axis_tick_mark_location(data, axis))? != location {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} {}-axis tick-mark-location update failed validation",
            axis.label(),
        )));
    }
    Ok(())
}

/// Read whether iWork shows the minimum label on a native chart value axis.
pub(crate) fn chart_value_axis_minimum_label_visible(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<bool> {
    chart_axis_style_switch_visible(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        ChartAxis::Value,
        AxisStyleSwitch::ValueMinimumLabel,
    )
}

/// Set whether iWork shows the minimum label on a native chart value axis.
pub(crate) fn set_chart_value_axis_minimum_label_visible(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    visible: bool,
) -> Result<()> {
    set_chart_axis_style_switch_visible(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        ChartAxis::Value,
        AxisStyleSwitch::ValueMinimumLabel,
        visible,
    )
}

fn chart_axis_style_switch_visible(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    axis: ChartAxis,
    style_switch: AxisStyleSwitch,
) -> Result<bool> {
    style_switch.validate_axis(axis)?;
    axis_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        axis,
    )?
    .read(package, |data| {
        read_axis_style_switch_visibility(data, axis, style_switch)
    })
}

fn set_chart_axis_style_switch_visible(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    axis: ChartAxis,
    style_switch: AxisStyleSwitch,
    visible: bool,
) -> Result<()> {
    style_switch.validate_axis(axis)?;
    let slot = axis_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        axis,
    )?;
    if slot.read(package, |data| {
        read_axis_style_switch_visibility(data, axis, style_switch)
    })? == visible
    {
        return Ok(());
    }
    slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
    slot.update(package, |data| {
        patch_axis_style_switch_visibility(data, axis, style_switch, visible)
    })?;
    if slot.read(package, |data| {
        read_axis_style_switch_visibility(data, axis, style_switch)
    })? != visible
    {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} {}-axis {} update failed validation",
            axis.label(),
            style_switch.label()
        )));
    }
    Ok(())
}

/// Decode a `TSCH.ChartAxisStyleArchive` and return one native style switch.
fn read_axis_style_switch_visibility(
    data: &[u8],
    axis: ChartAxis,
    style_switch: AxisStyleSwitch,
) -> Result<bool> {
    style_switch.validate_axis(axis)?;
    let Some(extension) = generated_axis_style_extension(data)? else {
        return Ok(style_switch.default_visible());
    };
    let generated = tsch::generated::ChartAxisStyleArchive::decode(extension)?;
    Ok(style_switch
        .visible(&generated, axis)
        .unwrap_or_else(|| style_switch.default_visible()))
}

/// Decode the native major-tick-mark location for one chart axis.
fn read_axis_tick_mark_location(data: &[u8], axis: ChartAxis) -> Result<ChartAxisTickMarkLocation> {
    let Some(extension) = generated_axis_style_extension(data)? else {
        return Ok(ChartAxisTickMarkLocation::default());
    };
    let generated = tsch::generated::ChartAxisStyleArchive::decode(extension)?;
    Ok(axis_tick_mark_location(&generated, axis)
        .map(ChartAxisTickMarkLocation::from_raw)
        .unwrap_or_default())
}

const fn tick_mark_location_field(axis: ChartAxis) -> u32 {
    match axis {
        ChartAxis::Category => CATEGORY_TICK_MARK_LOCATION_FIELD,
        ChartAxis::Value => VALUE_TICK_MARK_LOCATION_FIELD,
    }
}

fn axis_tick_mark_location(
    generated: &tsch::generated::ChartAxisStyleArchive,
    axis: ChartAxis,
) -> Option<i32> {
    match axis {
        ChartAxis::Category => generated.tschchartaxiscategorytickmarklocation,
        ChartAxis::Value => generated.tschchartaxisvaluetickmarklocation,
    }
}

pub(crate) fn axis_style_slot(
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
    pub(crate) fn read<T>(
        &self,
        package: &IWorkPackage,
        read: impl FnOnce(&[u8]) -> Result<T>,
    ) -> Result<T> {
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

    pub(crate) fn update(
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

fn patch_axis_style_switch_visibility(
    data: &[u8],
    axis: ChartAxis,
    style_switch: AxisStyleSwitch,
    visible: bool,
) -> Result<Vec<u8>> {
    style_switch.validate_axis(axis)?;
    let Some(extension) = generated_axis_style_extension(data)? else {
        let mut generated = tsch::generated::ChartAxisStyleArchive::default();
        style_switch.set(&mut generated, axis, visible);
        let encoded = generated.encode_to_vec();
        let patched = patch_length_delimited_field(
            data,
            GENERATED_CHART_AXIS_STYLE_EXTENSION_FIELD,
            false,
            Some(encoded.as_slice()),
        )?;
        validate_patched_axis_style_switch(&patched, axis, style_switch, visible)?;
        return Ok(patched);
    };

    let generated = tsch::generated::ChartAxisStyleArchive::decode(extension)?;
    let visible_present = style_switch.visible(&generated, axis).is_some();
    let extension = patch_varint_field(
        extension,
        style_switch.field(axis),
        visible_present,
        Some(u64::from(visible)),
    )?;
    let patched = patch_length_delimited_field(
        data,
        GENERATED_CHART_AXIS_STYLE_EXTENSION_FIELD,
        true,
        Some(extension.as_slice()),
    )?;
    validate_patched_axis_style_switch(&patched, axis, style_switch, visible)?;
    Ok(patched)
}

fn patch_axis_tick_mark_location(
    data: &[u8],
    axis: ChartAxis,
    location: ChartAxisTickMarkLocation,
) -> Result<Vec<u8>> {
    let Some(extension) = generated_axis_style_extension(data)? else {
        let mut generated = tsch::generated::ChartAxisStyleArchive::default();
        match axis {
            ChartAxis::Category => {
                generated.tschchartaxiscategorytickmarklocation = Some(location.into_raw())
            },
            ChartAxis::Value => {
                generated.tschchartaxisvaluetickmarklocation = Some(location.into_raw())
            },
        }
        let encoded = generated.encode_to_vec();
        let patched = patch_length_delimited_field(
            data,
            GENERATED_CHART_AXIS_STYLE_EXTENSION_FIELD,
            false,
            Some(encoded.as_slice()),
        )?;
        validate_patched_axis_tick_mark_location(&patched, axis, location)?;
        return Ok(patched);
    };

    let generated = tsch::generated::ChartAxisStyleArchive::decode(extension)?;
    let location_present = axis_tick_mark_location(&generated, axis).is_some();
    let extension = patch_varint_field(
        extension,
        tick_mark_location_field(axis),
        location_present,
        Some(location.into_raw() as u64),
    )?;
    let patched = patch_length_delimited_field(
        data,
        GENERATED_CHART_AXIS_STYLE_EXTENSION_FIELD,
        true,
        Some(extension.as_slice()),
    )?;
    validate_patched_axis_tick_mark_location(&patched, axis, location)?;
    Ok(patched)
}

fn validate_patched_axis_style_switch(
    data: &[u8],
    axis: ChartAxis,
    style_switch: AxisStyleSwitch,
    expected: bool,
) -> Result<()> {
    if read_axis_style_switch_visibility(data, axis, style_switch)? != expected {
        return Err(Error::InvalidFormat(format!(
            "{}-axis {} wire patch failed validation",
            axis.label(),
            style_switch.label()
        )));
    }
    Ok(())
}

fn validate_patched_axis_tick_mark_location(
    data: &[u8],
    axis: ChartAxis,
    expected: ChartAxisTickMarkLocation,
) -> Result<()> {
    if read_axis_tick_mark_location(data, axis)? != expected {
        return Err(Error::InvalidFormat(format!(
            "{}-axis tick-mark-location wire patch failed validation",
            axis.label(),
        )));
    }
    Ok(())
}

pub(crate) fn generated_axis_style_extension(data: &[u8]) -> Result<Option<&[u8]>> {
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

        let visible = patch_axis_style_switch_visibility(
            &original,
            ChartAxis::Category,
            AxisStyleSwitch::Line,
            true,
        )
        .unwrap();
        assert!(
            read_axis_style_switch_visibility(&visible, ChartAxis::Category, AxisStyleSwitch::Line)
                .unwrap()
        );
        assert!(
            read_axis_style_switch_visibility(&visible, ChartAxis::Value, AxisStyleSwitch::Line)
                .unwrap()
        );
        assert_unknown_fields_retained(&original, &visible);

        let restored = patch_axis_style_switch_visibility(
            &visible,
            ChartAxis::Category,
            AxisStyleSwitch::Line,
            false,
        )
        .unwrap();
        assert_eq!(restored, original);
    }

    #[test]
    fn major_gridline_patch_retains_axis_lines_and_unmapped_fields() {
        let generated = tsch::generated::ChartAxisStyleArchive {
            tschchartaxiscategoryshowaxis: Some(true),
            tschchartaxisvalueshowaxis: Some(false),
            tschchartaxiscategoryshowmajorgridlines: Some(false),
            tschchartaxisvalueshowmajorgridlines: Some(true),
            ..Default::default()
        };
        let original = axis_style_with_unknown_fields(generated);

        let visible = patch_axis_style_switch_visibility(
            &original,
            ChartAxis::Category,
            AxisStyleSwitch::MajorGridlines,
            true,
        )
        .unwrap();
        assert!(
            read_axis_style_switch_visibility(
                &visible,
                ChartAxis::Category,
                AxisStyleSwitch::MajorGridlines,
            )
            .unwrap()
        );
        assert!(
            read_axis_style_switch_visibility(&visible, ChartAxis::Category, AxisStyleSwitch::Line)
                .unwrap()
        );
        assert!(
            !read_axis_style_switch_visibility(&visible, ChartAxis::Value, AxisStyleSwitch::Line)
                .unwrap()
        );
        assert!(
            read_axis_style_switch_visibility(
                &visible,
                ChartAxis::Value,
                AxisStyleSwitch::MajorGridlines,
            )
            .unwrap()
        );
        assert_unknown_fields_retained(&original, &visible);

        let restored = patch_axis_style_switch_visibility(
            &visible,
            ChartAxis::Category,
            AxisStyleSwitch::MajorGridlines,
            false,
        )
        .unwrap();
        assert_eq!(restored, original);
    }

    #[test]
    fn minor_gridline_patch_retains_other_style_switches_and_unmapped_fields() {
        let generated = tsch::generated::ChartAxisStyleArchive {
            tschchartaxiscategoryshowaxis: Some(true),
            tschchartaxisvalueshowaxis: Some(false),
            tschchartaxiscategoryshowmajorgridlines: Some(false),
            tschchartaxisvalueshowmajorgridlines: Some(true),
            tschchartaxiscategoryshowminorgridlines: Some(false),
            tschchartaxisvalueshowminorgridlines: Some(true),
            ..Default::default()
        };
        let original = axis_style_with_unknown_fields(generated);

        let visible = patch_axis_style_switch_visibility(
            &original,
            ChartAxis::Category,
            AxisStyleSwitch::MinorGridlines,
            true,
        )
        .unwrap();
        assert!(
            read_axis_style_switch_visibility(
                &visible,
                ChartAxis::Category,
                AxisStyleSwitch::MinorGridlines,
            )
            .unwrap()
        );
        assert!(
            read_axis_style_switch_visibility(&visible, ChartAxis::Category, AxisStyleSwitch::Line)
                .unwrap()
        );
        assert!(
            !read_axis_style_switch_visibility(&visible, ChartAxis::Value, AxisStyleSwitch::Line)
                .unwrap()
        );
        assert!(
            !read_axis_style_switch_visibility(
                &visible,
                ChartAxis::Category,
                AxisStyleSwitch::MajorGridlines,
            )
            .unwrap()
        );
        assert!(
            read_axis_style_switch_visibility(
                &visible,
                ChartAxis::Value,
                AxisStyleSwitch::MajorGridlines,
            )
            .unwrap()
        );
        assert!(
            read_axis_style_switch_visibility(
                &visible,
                ChartAxis::Value,
                AxisStyleSwitch::MinorGridlines,
            )
            .unwrap()
        );
        assert_unknown_fields_retained(&original, &visible);

        let restored = patch_axis_style_switch_visibility(
            &visible,
            ChartAxis::Category,
            AxisStyleSwitch::MinorGridlines,
            false,
        )
        .unwrap();
        assert_eq!(restored, original);
    }

    #[test]
    fn minor_tick_mark_patch_retains_other_style_switches_and_unmapped_fields() {
        let generated = tsch::generated::ChartAxisStyleArchive {
            tschchartaxiscategoryshowaxis: Some(true),
            tschchartaxisvalueshowaxis: Some(false),
            tschchartaxiscategoryshowmajorgridlines: Some(false),
            tschchartaxisvalueshowmajorgridlines: Some(true),
            tschchartaxiscategoryshowminortickmarks: Some(true),
            tschchartaxisvalueshowminortickmarks: Some(true),
            ..Default::default()
        };
        let original = axis_style_with_unknown_fields(generated);

        let hidden = patch_axis_style_switch_visibility(
            &original,
            ChartAxis::Category,
            AxisStyleSwitch::MinorTickMarks,
            false,
        )
        .unwrap();
        assert!(
            !read_axis_style_switch_visibility(
                &hidden,
                ChartAxis::Category,
                AxisStyleSwitch::MinorTickMarks,
            )
            .unwrap()
        );
        assert!(
            read_axis_style_switch_visibility(
                &hidden,
                ChartAxis::Value,
                AxisStyleSwitch::MinorTickMarks,
            )
            .unwrap()
        );
        assert!(
            read_axis_style_switch_visibility(&hidden, ChartAxis::Category, AxisStyleSwitch::Line)
                .unwrap()
        );
        assert!(
            read_axis_style_switch_visibility(
                &hidden,
                ChartAxis::Value,
                AxisStyleSwitch::MajorGridlines,
            )
            .unwrap()
        );
        assert_unknown_fields_retained(&original, &hidden);

        let restored = patch_axis_style_switch_visibility(
            &hidden,
            ChartAxis::Category,
            AxisStyleSwitch::MinorTickMarks,
            true,
        )
        .unwrap();
        assert_eq!(restored, original);
    }

    #[test]
    fn minor_tick_marks_default_visible_when_missing_from_an_extension() {
        let original = axis_style_with_unknown_fields(tsch::generated::ChartAxisStyleArchive {
            tschchartaxisvalueshowaxis: Some(true),
            ..Default::default()
        });
        for axis in [ChartAxis::Category, ChartAxis::Value] {
            assert!(
                read_axis_style_switch_visibility(&original, axis, AxisStyleSwitch::MinorTickMarks)
                    .unwrap()
            );
            let hidden = patch_axis_style_switch_visibility(
                &original,
                axis,
                AxisStyleSwitch::MinorTickMarks,
                false,
            )
            .unwrap();
            assert!(
                !read_axis_style_switch_visibility(&hidden, axis, AxisStyleSwitch::MinorTickMarks,)
                    .unwrap()
            );
            assert_unknown_fields_retained(&original, &hidden);
        }
    }

    #[test]
    fn minor_tick_marks_default_visible_and_hidden_setting_creates_a_style_extension() {
        let original = tsch::ChartAxisStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();

        for axis in [ChartAxis::Category, ChartAxis::Value] {
            assert!(
                read_axis_style_switch_visibility(&original, axis, AxisStyleSwitch::MinorTickMarks)
                    .unwrap()
            );
            let hidden = patch_axis_style_switch_visibility(
                &original,
                axis,
                AxisStyleSwitch::MinorTickMarks,
                false,
            )
            .unwrap();
            assert!(
                !read_axis_style_switch_visibility(&hidden, axis, AxisStyleSwitch::MinorTickMarks,)
                    .unwrap()
            );
            assert!(generated_axis_style_extension(&hidden).unwrap().is_some());
        }
    }

    #[test]
    fn tick_mark_locations_round_trip_known_and_future_native_values() {
        let known = [
            (0, ChartAxisTickMarkLocation::None),
            (1, ChartAxisTickMarkLocation::Inside),
            (2, ChartAxisTickMarkLocation::Centered),
            (3, ChartAxisTickMarkLocation::Outside),
        ];
        for (raw, location) in known {
            assert_eq!(ChartAxisTickMarkLocation::from_raw(raw), location);
            assert_eq!(location.into_raw(), raw);
        }
        assert_eq!(
            ChartAxisTickMarkLocation::default(),
            ChartAxisTickMarkLocation::Centered
        );

        const FUTURE_LOCATION: i32 = 9_001;
        assert_eq!(
            ChartAxisTickMarkLocation::from_raw(FUTURE_LOCATION),
            ChartAxisTickMarkLocation::Unsupported(FUTURE_LOCATION)
        );
        assert_eq!(
            ChartAxisTickMarkLocation::Unsupported(FUTURE_LOCATION).into_raw(),
            FUTURE_LOCATION
        );
    }

    #[test]
    fn tick_mark_location_patch_retains_other_axis_and_unmapped_fields() {
        let generated = tsch::generated::ChartAxisStyleArchive {
            tschchartaxiscategorytickmarklocation: Some(ChartAxisTickMarkLocation::None.into_raw()),
            tschchartaxisvaluetickmarklocation: Some(ChartAxisTickMarkLocation::Inside.into_raw()),
            tschchartaxisvalueshowminorgridlines: Some(true),
            ..Default::default()
        };
        let original = axis_style_with_unknown_fields(generated);

        let outside = patch_axis_tick_mark_location(
            &original,
            ChartAxis::Category,
            ChartAxisTickMarkLocation::Outside,
        )
        .unwrap();
        assert_eq!(
            read_axis_tick_mark_location(&outside, ChartAxis::Category).unwrap(),
            ChartAxisTickMarkLocation::Outside
        );
        assert_eq!(
            read_axis_tick_mark_location(&outside, ChartAxis::Value).unwrap(),
            ChartAxisTickMarkLocation::Inside
        );
        assert!(
            read_axis_style_switch_visibility(
                &outside,
                ChartAxis::Value,
                AxisStyleSwitch::MinorGridlines,
            )
            .unwrap()
        );
        assert_unknown_fields_retained(&original, &outside);

        let restored = patch_axis_tick_mark_location(
            &outside,
            ChartAxis::Category,
            ChartAxisTickMarkLocation::None,
        )
        .unwrap();
        assert_eq!(restored, original);
    }

    #[test]
    fn tick_mark_location_defaults_centered_and_can_create_a_style_extension() {
        let original = tsch::ChartAxisStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();

        for axis in [ChartAxis::Category, ChartAxis::Value] {
            assert_eq!(
                read_axis_tick_mark_location(&original, axis).unwrap(),
                ChartAxisTickMarkLocation::Centered
            );
            let none =
                patch_axis_tick_mark_location(&original, axis, ChartAxisTickMarkLocation::None)
                    .unwrap();
            assert_eq!(
                read_axis_tick_mark_location(&none, axis).unwrap(),
                ChartAxisTickMarkLocation::None
            );
            assert!(generated_axis_style_extension(&none).unwrap().is_some());
        }
    }

    #[test]
    fn tick_mark_location_preserves_future_native_values() {
        const FUTURE_LOCATION: i32 = 9_001;
        let original = axis_style_with_unknown_fields(tsch::generated::ChartAxisStyleArchive {
            tschchartaxisvaluetickmarklocation: Some(FUTURE_LOCATION),
            ..Default::default()
        });
        assert_eq!(
            read_axis_tick_mark_location(&original, ChartAxis::Value).unwrap(),
            ChartAxisTickMarkLocation::Unsupported(FUTURE_LOCATION)
        );

        let patched = patch_axis_tick_mark_location(
            &original,
            ChartAxis::Value,
            ChartAxisTickMarkLocation::Outside,
        )
        .unwrap();
        assert_eq!(
            read_axis_tick_mark_location(&patched, ChartAxis::Value).unwrap(),
            ChartAxisTickMarkLocation::Outside
        );
        assert_unknown_fields_retained(&original, &patched);
    }

    #[test]
    fn value_axis_minimum_label_patch_retains_other_style_switches_and_unmapped_fields() {
        let generated = tsch::generated::ChartAxisStyleArchive {
            tschchartaxisvalueshowaxis: Some(true),
            tschchartaxisvalueshowmajorgridlines: Some(true),
            tschchartaxisvalueshowminimumlabel: Some(true),
            ..Default::default()
        };
        let original = axis_style_with_unknown_fields(generated);

        let hidden = patch_axis_style_switch_visibility(
            &original,
            ChartAxis::Value,
            AxisStyleSwitch::ValueMinimumLabel,
            false,
        )
        .unwrap();
        assert!(
            !read_axis_style_switch_visibility(
                &hidden,
                ChartAxis::Value,
                AxisStyleSwitch::ValueMinimumLabel,
            )
            .unwrap()
        );
        assert!(
            read_axis_style_switch_visibility(&hidden, ChartAxis::Value, AxisStyleSwitch::Line)
                .unwrap()
        );
        assert!(
            read_axis_style_switch_visibility(
                &hidden,
                ChartAxis::Value,
                AxisStyleSwitch::MajorGridlines,
            )
            .unwrap()
        );
        assert_unknown_fields_retained(&original, &hidden);

        let restored = patch_axis_style_switch_visibility(
            &hidden,
            ChartAxis::Value,
            AxisStyleSwitch::ValueMinimumLabel,
            true,
        )
        .unwrap();
        assert_eq!(restored, original);
    }

    #[test]
    fn value_axis_minimum_label_defaults_visible_when_missing_from_an_extension() {
        let original = axis_style_with_unknown_fields(tsch::generated::ChartAxisStyleArchive {
            tschchartaxisvalueshowaxis: Some(true),
            ..Default::default()
        });
        assert!(
            read_axis_style_switch_visibility(
                &original,
                ChartAxis::Value,
                AxisStyleSwitch::ValueMinimumLabel,
            )
            .unwrap()
        );

        let hidden = patch_axis_style_switch_visibility(
            &original,
            ChartAxis::Value,
            AxisStyleSwitch::ValueMinimumLabel,
            false,
        )
        .unwrap();
        assert!(
            !read_axis_style_switch_visibility(
                &hidden,
                ChartAxis::Value,
                AxisStyleSwitch::ValueMinimumLabel,
            )
            .unwrap()
        );
        assert!(
            read_axis_style_switch_visibility(&hidden, ChartAxis::Value, AxisStyleSwitch::Line)
                .unwrap()
        );
        assert_unknown_fields_retained(&original, &hidden);
    }

    #[test]
    fn style_switch_patch_creates_a_style_extension_when_missing() {
        let original = tsch::ChartAxisStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();

        let visible = patch_axis_style_switch_visibility(
            &original,
            ChartAxis::Value,
            AxisStyleSwitch::MajorGridlines,
            true,
        )
        .unwrap();
        assert!(
            read_axis_style_switch_visibility(
                &visible,
                ChartAxis::Value,
                AxisStyleSwitch::MajorGridlines,
            )
            .unwrap()
        );
        assert!(
            !read_axis_style_switch_visibility(&visible, ChartAxis::Value, AxisStyleSwitch::Line)
                .unwrap()
        );
        assert!(
            read_axis_style_switch_visibility(
                &original,
                ChartAxis::Value,
                AxisStyleSwitch::ValueMinimumLabel,
            )
            .unwrap()
        );
        assert!(
            read_axis_style_switch_visibility(
                &original,
                ChartAxis::Category,
                AxisStyleSwitch::ValueMinimumLabel,
            )
            .is_err()
        );
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
