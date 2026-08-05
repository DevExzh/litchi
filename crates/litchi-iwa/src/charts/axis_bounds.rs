//! Lossless native value-axis scale-bound storage and mutation.
//!
//! iWork stores the manual minimum and maximum of a chart's primary value
//! axis in the generated extension of `TSCH.ChartAxisNonStyleArchive`. This
//! module represents automatic versus fixed bounds explicitly, preserves both
//! protobuf layers losslessly, and patches only those two nested messages.

use prost::Message;

use litchi_iwa_common::chart::axis::{
    Axis,
    bounds::{Bound, Bounds},
};

use crate::charts::axis::{
    GENERATED_CHART_AXIS_NON_STYLE_EXTENSION_FIELD, axis_non_style_slot,
    generated_axis_non_style_extension,
};
use crate::protobuf::tsch;
use crate::wire::patch_length_delimited_field;
use crate::{Error, IWorkPackage, Result};

/// `tschchartaxisdefaultusermax` in `TSCH.Generated.ChartAxisNonStyleArchive`.
const VALUE_AXIS_MAXIMUM_FIELD: u32 = 17;
/// `tschchartaxisdefaultusermin` in `TSCH.Generated.ChartAxisNonStyleArchive`.
const VALUE_AXIS_MINIMUM_FIELD: u32 = 18;

/// Read the manual bounds of one native chart's primary value axis.
pub(crate) fn chart_value_axis_bounds(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<Bounds> {
    axis_non_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        Axis::Value,
    )?
    .read(package, read_value_axis_bounds)
}

/// Set the manual bounds of one native chart's primary value axis.
pub(crate) fn set_chart_value_axis_bounds(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    bounds: Bounds,
) -> Result<()> {
    let slot = axis_non_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        Axis::Value,
    )?;
    if slot.read(package, read_value_axis_bounds)? == bounds {
        return Ok(());
    }
    slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
    slot.update(package, |data| patch_value_axis_bounds(data, bounds))?;
    if slot.read(package, read_value_axis_bounds)? != bounds {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} value-axis bounds update failed validation"
        )));
    }
    Ok(())
}

fn read_value_axis_bounds(data: &[u8]) -> Result<Bounds> {
    let Some(extension) = generated_axis_non_style_extension(data)? else {
        return Ok(Bounds::automatic());
    };
    let generated = tsch::generated::ChartAxisNonStyleArchive::decode(extension)?;
    Ok(Bounds::new(
        decode_bound(generated.tschchartaxisdefaultusermin.as_ref(), "minimum")?,
        decode_bound(generated.tschchartaxisdefaultusermax.as_ref(), "maximum")?,
    )?)
}

fn decode_bound(
    archive: Option<&tsch::ChartsNsNumberDoubleArchive>,
    label: &str,
) -> Result<Option<Bound>> {
    let Some(archive) = archive else {
        return Ok(None);
    };
    let value = archive.number_archive.ok_or_else(|| {
        Error::InvalidFormat(format!(
            "chart value-axis {label} is missing its numeric value"
        ))
    })?;
    Bound::new(value)
        .map(Some)
        .map_err(|error| Error::InvalidFormat(error.to_string()))
}

fn encoded_bound(bound: Option<Bound>) -> Option<Vec<u8>> {
    bound.map(|bound| {
        tsch::ChartsNsNumberDoubleArchive {
            number_archive: Some(bound.value()),
        }
        .encode_to_vec()
    })
}

fn patch_value_axis_bounds(data: &[u8], bounds: Bounds) -> Result<Vec<u8>> {
    let Some(extension) = generated_axis_non_style_extension(data)? else {
        if bounds == Bounds::automatic() {
            return Ok(data.to_vec());
        }
        let generated = tsch::generated::ChartAxisNonStyleArchive {
            tschchartaxisdefaultusermin: bounds.minimum().map(number_archive),
            tschchartaxisdefaultusermax: bounds.maximum().map(number_archive),
            ..Default::default()
        };
        let patched = patch_length_delimited_field(
            data,
            GENERATED_CHART_AXIS_NON_STYLE_EXTENSION_FIELD,
            false,
            Some(generated.encode_to_vec().as_slice()),
        )?;
        validate_patched_value_axis_bounds(&patched, bounds)?;
        return Ok(patched);
    };

    let generated = tsch::generated::ChartAxisNonStyleArchive::decode(extension)?;
    let minimum = encoded_bound(bounds.minimum());
    let maximum = encoded_bound(bounds.maximum());
    let extension = patch_length_delimited_field(
        extension,
        VALUE_AXIS_MINIMUM_FIELD,
        generated.tschchartaxisdefaultusermin.is_some(),
        minimum.as_deref(),
    )?;
    let extension = patch_length_delimited_field(
        &extension,
        VALUE_AXIS_MAXIMUM_FIELD,
        generated.tschchartaxisdefaultusermax.is_some(),
        maximum.as_deref(),
    )?;
    let patched = patch_length_delimited_field(
        data,
        GENERATED_CHART_AXIS_NON_STYLE_EXTENSION_FIELD,
        true,
        Some(extension.as_slice()),
    )?;
    validate_patched_value_axis_bounds(&patched, bounds)?;
    Ok(patched)
}

fn number_archive(bound: Bound) -> tsch::ChartsNsNumberDoubleArchive {
    tsch::ChartsNsNumberDoubleArchive {
        number_archive: Some(bound.value()),
    }
}

fn validate_patched_value_axis_bounds(data: &[u8], expected: Bounds) -> Result<()> {
    if read_value_axis_bounds(data)? != expected {
        return Err(Error::InvalidFormat(
            "chart value-axis bounds wire patch failed validation".to_owned(),
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
    fn bounds_reject_nonfinite_and_inverted_ranges() {
        assert!(Bound::new(f64::NAN).is_err());
        assert!(Bound::new(f64::INFINITY).is_err());

        let low = Bound::new(-1.0).unwrap();
        let high = Bound::new(1.0).unwrap();
        assert!(Bounds::fixed(high, low).is_err());
        assert_eq!(Bounds::fixed(low, high).unwrap().minimum(), Some(low));
    }

    #[test]
    fn value_axis_bounds_patch_retains_other_fields_and_unmapped_data() {
        let original =
            axis_non_style_with_unknown_fields(tsch::generated::ChartAxisNonStyleArchive {
                tschchartaxiscategoryshowtitle: Some(true),
                tschchartaxiscategorytitle: Some("Month".to_owned()),
                tschchartaxisdefaultusermin: Some(number_archive(Bound::new(-5.0).unwrap())),
                tschchartaxisdefaultusermax: Some(number_archive(Bound::new(50.0).unwrap())),
                ..Default::default()
            });
        let replacement =
            Bounds::fixed(Bound::new(0.0).unwrap(), Bound::new(30.0).unwrap()).unwrap();

        let patched = patch_value_axis_bounds(&original, replacement).unwrap();
        assert_eq!(read_value_axis_bounds(&patched).unwrap(), replacement);
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

        let restored = patch_value_axis_bounds(
            &patched,
            Bounds::fixed(Bound::new(-5.0).unwrap(), Bound::new(50.0).unwrap()).unwrap(),
        )
        .unwrap();
        assert_eq!(restored, original);
    }

    #[test]
    fn automatic_bounds_remove_only_the_requested_fields() {
        let original =
            axis_non_style_with_unknown_fields(tsch::generated::ChartAxisNonStyleArchive {
                tschchartaxisdefaultusermin: Some(number_archive(Bound::new(0.0).unwrap())),
                tschchartaxisdefaultusermax: Some(number_archive(Bound::new(30.0).unwrap())),
                tschchartaxisvalueshowtitle: Some(true),
                tschchartaxisvaluetitle: Some("Revenue".to_owned()),
                ..Default::default()
            });

        let automatic = patch_value_axis_bounds(&original, Bounds::automatic()).unwrap();
        assert_eq!(
            read_value_axis_bounds(&automatic).unwrap(),
            Bounds::automatic()
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
    fn value_axis_bounds_patch_creates_an_extension_when_missing() {
        let original = tsch::ChartAxisNonStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();
        let bounds = Bounds::new(Some(Bound::new(-2.0).unwrap()), None).unwrap();

        let patched = patch_value_axis_bounds(&original, bounds).unwrap();
        assert_eq!(read_value_axis_bounds(&patched).unwrap(), bounds);
        assert_eq!(
            patch_value_axis_bounds(&original, Bounds::automatic()).unwrap(),
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
