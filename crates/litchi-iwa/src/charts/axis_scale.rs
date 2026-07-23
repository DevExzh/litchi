//! Lossless native value-axis scale storage and mutation.
//!
//! iWork stores the scale mode for a chart's primary value axis in the
//! generated extension of `TSCH.ChartAxisNonStyleArchive`. This module keeps
//! the native integer forward-compatible and patches only that field.

use prost::Message;

use crate::charts::ChartAxis;
use crate::charts::axis::{
    GENERATED_CHART_AXIS_NON_STYLE_EXTENSION_FIELD, axis_non_style_slot,
    generated_axis_non_style_extension,
};
use crate::protobuf::tsch;
use crate::wire::{patch_length_delimited_field, patch_varint_field};
use crate::{Error, IWorkPackage, Result};

/// `tschchartaxisvaluescale` in `TSCH.Generated.ChartAxisNonStyleArchive`.
const VALUE_AXIS_SCALE_FIELD: u32 = 8;

/// The numeric scale used for a native chart's primary value axis.
///
/// The known values correspond to the `Axis Scale` pop-up in Pages, Numbers,
/// and Keynote. `Unsupported` preserves a future native value without
/// changing it during a read-modify-write cycle.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
#[non_exhaustive]
pub enum ChartValueAxisScale {
    /// Plot values along a linear scale.
    #[default]
    Linear,
    /// Plot values along a logarithmic scale.
    Logarithmic,
    /// Preserve an unrecognized native iWork value.
    Unsupported(i32),
}

impl ChartValueAxisScale {
    /// Decode the integer stored by the iWork protobuf schema.
    pub const fn from_raw(value: i32) -> Self {
        match value {
            1 => Self::Linear,
            2 => Self::Logarithmic,
            value => Self::Unsupported(value),
        }
    }

    /// Return the integer used by the iWork protobuf schema.
    pub const fn into_raw(self) -> i32 {
        match self {
            Self::Linear => 1,
            Self::Logarithmic => 2,
            Self::Unsupported(value) => value,
        }
    }
}

/// Read the scale mode of one native chart's primary value axis.
pub(crate) fn chart_value_axis_scale(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<ChartValueAxisScale> {
    axis_non_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        ChartAxis::Value,
    )?
    .read(package, read_value_axis_scale)
}

/// Set the scale mode of one native chart's primary value axis.
pub(crate) fn set_chart_value_axis_scale(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    scale: ChartValueAxisScale,
) -> Result<()> {
    let slot = axis_non_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        ChartAxis::Value,
    )?;
    if slot.read(package, read_value_axis_scale)? == scale {
        return Ok(());
    }
    slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
    slot.update(package, |data| patch_value_axis_scale(data, scale))?;
    if slot.read(package, read_value_axis_scale)? != scale {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} value-axis scale update failed validation"
        )));
    }
    Ok(())
}

fn read_value_axis_scale(data: &[u8]) -> Result<ChartValueAxisScale> {
    let Some(extension) = generated_axis_non_style_extension(data)? else {
        return Ok(ChartValueAxisScale::default());
    };
    let generated = tsch::generated::ChartAxisNonStyleArchive::decode(extension)?;
    Ok(generated
        .tschchartaxisvaluescale
        .map(ChartValueAxisScale::from_raw)
        .unwrap_or_default())
}

fn patch_value_axis_scale(data: &[u8], scale: ChartValueAxisScale) -> Result<Vec<u8>> {
    let Some(extension) = generated_axis_non_style_extension(data)? else {
        if scale == ChartValueAxisScale::default() {
            return Ok(data.to_vec());
        }
        let generated = tsch::generated::ChartAxisNonStyleArchive {
            tschchartaxisvaluescale: Some(scale.into_raw()),
            ..Default::default()
        };
        let patched = patch_length_delimited_field(
            data,
            GENERATED_CHART_AXIS_NON_STYLE_EXTENSION_FIELD,
            false,
            Some(generated.encode_to_vec().as_slice()),
        )?;
        validate_patched_value_axis_scale(&patched, scale)?;
        return Ok(patched);
    };

    let generated = tsch::generated::ChartAxisNonStyleArchive::decode(extension)?;
    let scale_present = generated.tschchartaxisvaluescale.is_some();
    let value = (scale_present || scale != ChartValueAxisScale::default())
        .then_some(scale.into_raw() as u64);
    let extension = patch_varint_field(extension, VALUE_AXIS_SCALE_FIELD, scale_present, value)?;
    let patched = patch_length_delimited_field(
        data,
        GENERATED_CHART_AXIS_NON_STYLE_EXTENSION_FIELD,
        true,
        Some(extension.as_slice()),
    )?;
    validate_patched_value_axis_scale(&patched, scale)?;
    Ok(patched)
}

fn validate_patched_value_axis_scale(data: &[u8], expected: ChartValueAxisScale) -> Result<()> {
    if read_value_axis_scale(data)? != expected {
        return Err(Error::InvalidFormat(
            "chart value-axis scale wire patch failed validation".to_owned(),
        ));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::protobuf::tss;
    use crate::wire::{append_length_delimited_field, append_varint_field, parse_wire_fields};

    const UNMAPPED_OUTER_FIELD: u32 = 4_096;
    const UNMAPPED_GENERATED_FIELD: u32 = 4_097;
    const UNMAPPED_VALUE: u64 = 42;

    #[test]
    fn value_axis_scales_round_trip_known_and_future_native_values() {
        let known = [
            (1, ChartValueAxisScale::Linear),
            (2, ChartValueAxisScale::Logarithmic),
        ];
        for (raw, scale) in known {
            assert_eq!(ChartValueAxisScale::from_raw(raw), scale);
            assert_eq!(scale.into_raw(), raw);
        }
        assert_eq!(ChartValueAxisScale::default(), ChartValueAxisScale::Linear);
        assert_eq!(
            ChartValueAxisScale::from_raw(0),
            ChartValueAxisScale::Unsupported(0)
        );

        const FUTURE_SCALE: i32 = 9_001;
        assert_eq!(
            ChartValueAxisScale::from_raw(FUTURE_SCALE),
            ChartValueAxisScale::Unsupported(FUTURE_SCALE)
        );
        assert_eq!(
            ChartValueAxisScale::Unsupported(FUTURE_SCALE).into_raw(),
            FUTURE_SCALE
        );
    }

    #[test]
    fn value_axis_scale_patch_retains_other_fields_and_unmapped_data() {
        let original =
            axis_non_style_with_unknown_fields(tsch::generated::ChartAxisNonStyleArchive {
                tschchartaxisvaluescale: Some(ChartValueAxisScale::Linear.into_raw()),
                tschchartaxisvaluenumberofmajorgridlines: Some(5),
                tschchartaxiscategorytitle: Some("Month".to_owned()),
                ..Default::default()
            });

        let logarithmic =
            patch_value_axis_scale(&original, ChartValueAxisScale::Logarithmic).unwrap();
        assert_eq!(
            read_value_axis_scale(&logarithmic).unwrap(),
            ChartValueAxisScale::Logarithmic
        );
        let generated = tsch::generated::ChartAxisNonStyleArchive::decode(
            generated_axis_non_style_extension(&logarithmic)
                .unwrap()
                .unwrap(),
        )
        .unwrap();
        assert_eq!(generated.tschchartaxisvaluenumberofmajorgridlines, Some(5));
        assert_eq!(
            generated.tschchartaxiscategorytitle.as_deref(),
            Some("Month")
        );
        assert_unknown_fields_retained(&original, &logarithmic);

        let restored = patch_value_axis_scale(&logarithmic, ChartValueAxisScale::Linear).unwrap();
        assert_eq!(restored, original);
    }

    #[test]
    fn value_axis_scale_defaults_linear_and_creates_an_extension_when_needed() {
        let original = tsch::ChartAxisNonStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();
        assert_eq!(
            read_value_axis_scale(&original).unwrap(),
            ChartValueAxisScale::Linear
        );
        assert_eq!(
            patch_value_axis_scale(&original, ChartValueAxisScale::Linear).unwrap(),
            original
        );

        let logarithmic =
            patch_value_axis_scale(&original, ChartValueAxisScale::Logarithmic).unwrap();
        assert_eq!(
            read_value_axis_scale(&logarithmic).unwrap(),
            ChartValueAxisScale::Logarithmic
        );
        assert!(
            generated_axis_non_style_extension(&logarithmic)
                .unwrap()
                .is_some()
        );
    }

    #[test]
    fn value_axis_scale_preserves_future_native_values() {
        const FUTURE_SCALE: i32 = 9_001;
        let original =
            axis_non_style_with_unknown_fields(tsch::generated::ChartAxisNonStyleArchive {
                tschchartaxisvaluescale: Some(FUTURE_SCALE),
                ..Default::default()
            });
        assert_eq!(
            read_value_axis_scale(&original).unwrap(),
            ChartValueAxisScale::Unsupported(FUTURE_SCALE)
        );

        let patched = patch_value_axis_scale(&original, ChartValueAxisScale::Logarithmic).unwrap();
        assert_eq!(
            read_value_axis_scale(&patched).unwrap(),
            ChartValueAxisScale::Logarithmic
        );
        assert_unknown_fields_retained(&original, &patched);
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
                UNMAPPED_GENERATED_FIELD,
            ),
            raw_field(
                generated_axis_non_style_extension(original)
                    .unwrap()
                    .unwrap(),
                UNMAPPED_GENERATED_FIELD,
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
