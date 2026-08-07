//! Lossless native value-axis scale-step storage and mutation.
//!
//! iWork stores the major and minor step counts for a chart's primary value
//! axis in the generated extension of `TSCH.ChartAxisNonStyleArchive`. This
//! module represents automatic versus explicit step counts with distinct types,
//! preserves both protobuf layers losslessly, and patches only those fields.

use prost::Message;

use litchi_iwa_common::chart::axis::{
    Axis,
    steps::{MajorStepCount, MinorStepCount, Steps},
};

use crate::charts::axis::{
    GENERATED_CHART_AXIS_NON_STYLE_EXTENSION_FIELD, axis_non_style_slot,
    generated_axis_non_style_extension,
};
use crate::protobuf::tsch;
use crate::wire::patch_varint_field;
use crate::{Error, IWorkPackage, Result};

/// `tschchartaxisvaluenumberofmajorgridlines` in
/// `TSCH.Generated.ChartAxisNonStyleArchive`.
const VALUE_AXIS_MAJOR_STEPS_FIELD: u32 = 5;
/// `tschchartaxisvaluenumberofminorgridlines` in
/// `TSCH.Generated.ChartAxisNonStyleArchive`.
const VALUE_AXIS_MINOR_STEPS_FIELD: u32 = 6;
/// Read the scale steps of one native chart's primary value axis.
pub(crate) fn chart_value_axis_steps(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<Steps> {
    axis_non_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        Axis::Value,
    )?
    .read(package, read_value_axis_steps)
}

/// Set the scale steps of one native chart's primary value axis.
pub(crate) fn set_chart_value_axis_steps(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    steps: Steps,
) -> Result<()> {
    let slot = axis_non_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        Axis::Value,
    )?;
    if slot.read(package, read_value_axis_steps)? == steps {
        return Ok(());
    }
    slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
    slot.update(package, |data| patch_value_axis_steps(data, steps))?;
    if slot.read(package, read_value_axis_steps)? != steps {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} value-axis steps update failed validation"
        )));
    }
    Ok(())
}

fn read_value_axis_steps(data: &[u8]) -> Result<Steps> {
    let Some(extension) = generated_axis_non_style_extension(data)? else {
        return Ok(Steps::automatic());
    };
    let generated = tsch::generated::ChartAxisNonStyleArchive::decode(extension)?;
    Ok(Steps::new(
        decode_major_steps(generated.tschchartaxisvaluenumberofmajorgridlines)?,
        decode_minor_steps(generated.tschchartaxisvaluenumberofminorgridlines)?,
    ))
}

fn decode_major_steps(value: Option<i32>) -> Result<Option<MajorStepCount>> {
    let Some(value) = value else {
        return Ok(None);
    };
    let value = native_step_count(value, "major")?;
    Ok(Some(MajorStepCount::new(value)?))
}

fn decode_minor_steps(value: Option<i32>) -> Result<Option<MinorStepCount>> {
    let Some(value) = value else {
        return Ok(None);
    };
    let value = native_step_count(value, "minor")?;
    Ok(Some(MinorStepCount::new(value)?))
}

fn native_step_count(value: i32, label: &str) -> Result<u32> {
    u32::try_from(value).map_err(|_| {
        Error::InvalidFormat(format!(
            "chart value-axis {label} step count must not be negative"
        ))
    })
}

fn patch_value_axis_steps(data: &[u8], steps: Steps) -> Result<Vec<u8>> {
    let Some(extension) = generated_axis_non_style_extension(data)? else {
        if steps == Steps::automatic() {
            return Ok(data.to_vec());
        }
        let generated = tsch::generated::ChartAxisNonStyleArchive {
            tschchartaxisvaluenumberofmajorgridlines: steps
                .major()
                .map(|value| value.value() as i32),
            tschchartaxisvaluenumberofminorgridlines: steps
                .minor()
                .map(|value| value.value() as i32),
            ..Default::default()
        };
        let patched = crate::wire::patch_length_delimited_field(
            data,
            GENERATED_CHART_AXIS_NON_STYLE_EXTENSION_FIELD,
            false,
            Some(generated.encode_to_vec().as_slice()),
        )?;
        validate_patched_value_axis_steps(&patched, steps)?;
        return Ok(patched);
    };

    let generated = tsch::generated::ChartAxisNonStyleArchive::decode(extension)?;
    let extension = patch_varint_field(
        extension,
        VALUE_AXIS_MAJOR_STEPS_FIELD,
        generated.tschchartaxisvaluenumberofmajorgridlines.is_some(),
        steps.major().map(|value| u64::from(value.value())),
    )?;
    let extension = patch_varint_field(
        &extension,
        VALUE_AXIS_MINOR_STEPS_FIELD,
        generated.tschchartaxisvaluenumberofminorgridlines.is_some(),
        steps.minor().map(|value| u64::from(value.value())),
    )?;
    let patched = crate::wire::patch_length_delimited_field(
        data,
        GENERATED_CHART_AXIS_NON_STYLE_EXTENSION_FIELD,
        true,
        Some(extension.as_slice()),
    )?;
    validate_patched_value_axis_steps(&patched, steps)?;
    Ok(patched)
}

fn validate_patched_value_axis_steps(data: &[u8], expected: Steps) -> Result<()> {
    if read_value_axis_steps(data)? != expected {
        return Err(Error::InvalidFormat(
            "chart value-axis steps wire patch failed validation".to_owned(),
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
    const MAX_NATIVE_AXIS_STEP_COUNT: u32 = i32::MAX as u32;

    #[test]
    fn step_counts_reject_values_outside_native_ranges() {
        assert!(MajorStepCount::new(0).is_err());
        assert_eq!(MajorStepCount::ONE.value(), 1);
        assert!(MajorStepCount::new(MAX_NATIVE_AXIS_STEP_COUNT + 1).is_err());
        assert_eq!(MinorStepCount::new(0).unwrap().value(), 0);
        assert!(MinorStepCount::new(MAX_NATIVE_AXIS_STEP_COUNT + 1).is_err());
    }

    #[test]
    fn negative_native_step_counts_are_rejected() {
        let data = axis_non_style_with_unknown_fields(tsch::generated::ChartAxisNonStyleArchive {
            tschchartaxisvaluenumberofmajorgridlines: Some(-1),
            ..Default::default()
        });
        assert!(read_value_axis_steps(&data).is_err());

        let data = axis_non_style_with_unknown_fields(tsch::generated::ChartAxisNonStyleArchive {
            tschchartaxisvaluenumberofminorgridlines: Some(-1),
            ..Default::default()
        });
        assert!(read_value_axis_steps(&data).is_err());
    }

    #[test]
    fn value_axis_steps_patch_retains_other_fields_and_unmapped_data() {
        let original =
            axis_non_style_with_unknown_fields(tsch::generated::ChartAxisNonStyleArchive {
                tschchartaxiscategoryshowtitle: Some(true),
                tschchartaxiscategorytitle: Some("Month".to_owned()),
                tschchartaxisvaluenumberofmajorgridlines: Some(5),
                tschchartaxisvaluenumberofminorgridlines: Some(1),
                ..Default::default()
            });
        let replacement = Steps::fixed(
            MajorStepCount::new(6).unwrap(),
            MinorStepCount::new(2).unwrap(),
        );

        let patched = patch_value_axis_steps(&original, replacement).unwrap();
        assert_eq!(read_value_axis_steps(&patched).unwrap(), replacement);
        let generated = tsch::generated::ChartAxisNonStyleArchive::decode(
            generated_axis_non_style_extension(&patched)
                .unwrap()
                .unwrap(),
        )
        .unwrap();
        assert_eq!(
            generated.tschchartaxiscategorytitle.as_deref(),
            Some("Month")
        );
        assert_unknown_fields_retained(&original, &patched);

        let restored = patch_value_axis_steps(
            &patched,
            Steps::fixed(
                MajorStepCount::new(5).unwrap(),
                MinorStepCount::new(1).unwrap(),
            ),
        )
        .unwrap();
        assert_eq!(restored, original);
    }

    #[test]
    fn automatic_steps_remove_only_the_requested_fields() {
        let original =
            axis_non_style_with_unknown_fields(tsch::generated::ChartAxisNonStyleArchive {
                tschchartaxisvaluenumberofmajorgridlines: Some(5),
                tschchartaxisvaluenumberofminorgridlines: Some(1),
                tschchartaxisvalueshowtitle: Some(true),
                tschchartaxisvaluetitle: Some("Revenue".to_owned()),
                ..Default::default()
            });

        let automatic = patch_value_axis_steps(&original, Steps::automatic()).unwrap();
        assert_eq!(
            read_value_axis_steps(&automatic).unwrap(),
            Steps::automatic()
        );
        let generated = tsch::generated::ChartAxisNonStyleArchive::decode(
            generated_axis_non_style_extension(&automatic)
                .unwrap()
                .unwrap(),
        )
        .unwrap();
        assert_eq!(
            generated.tschchartaxisvaluetitle.as_deref(),
            Some("Revenue")
        );
        assert_unknown_fields_retained(&original, &automatic);
    }

    #[test]
    fn value_axis_steps_patch_creates_an_extension_when_missing() {
        let original = tsch::ChartAxisNonStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();
        let steps = Steps::new(Some(MajorStepCount::new(4).unwrap()), None);

        let patched = patch_value_axis_steps(&original, steps).unwrap();
        assert_eq!(read_value_axis_steps(&patched).unwrap(), steps);

        let minor_only = Steps::new(None, Some(MinorStepCount::new(0).unwrap()));
        let patched = patch_value_axis_steps(&original, minor_only).unwrap();
        assert_eq!(read_value_axis_steps(&patched).unwrap(), minor_only);

        assert_eq!(
            patch_value_axis_steps(&original, Steps::automatic()).unwrap(),
            original
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
            .filter(|field| field.number() == number)
            .map(|field| data[field.start()..field.end()].to_vec())
            .collect()
    }
}
