//! Lossless label-angle CRUD for native chart axes.
//!
//! Pages, Numbers, and Keynote persist a normalized degree value as a
//! protobuf float in the selected axis style's generated extension.

use prost::Message;

use litchi_iwa_common::chart::axis::{Axis, label_angle::LabelAngle};

use crate::charts::axis_style::{
    GENERATED_CHART_AXIS_STYLE_EXTENSION_FIELD, axis_style_slot, generated_axis_style_extension,
};
use crate::protobuf::tsch;
use crate::wire::{parse_wire_fields, patch_fixed32_field, patch_length_delimited_field};
use crate::{Error, IWorkPackage, Result};

const CATEGORY_LABEL_ANGLE_FIELD: u32 = 9;
const VALUE_LABEL_ANGLE_FIELD: u32 = 11;
/// Read the label angle for one native chart axis.
pub(crate) fn chart_axis_label_angle(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    axis: Axis,
) -> Result<LabelAngle> {
    axis_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        axis,
    )?
    .read(package, |data| read_axis_label_angle(data, axis))
}

/// Set or reset the label angle for one native chart axis.
pub(crate) fn set_chart_axis_label_angle(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    axis: Axis,
    angle: LabelAngle,
) -> Result<()> {
    let slot = axis_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        axis,
    )?;
    if slot.read(package, |data| read_axis_label_angle(data, axis))? == angle {
        return Ok(());
    }
    slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
    slot.update(package, |data| patch_axis_label_angle(data, axis, angle))?;
    if slot.read(package, |data| read_axis_label_angle(data, axis))? != angle {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} {}-axis label-angle update failed validation",
            axis.as_str()
        )));
    }
    Ok(())
}

fn read_axis_label_angle(data: &[u8], axis: Axis) -> Result<LabelAngle> {
    let Some(extension) = generated_axis_style_extension(data)? else {
        return Ok(LabelAngle::HORIZONTAL);
    };
    tsch::generated::ChartAxisStyleArchive::decode(extension)?;
    let Some(bits) = strict_optional_fixed32(extension, label_angle_field(axis))? else {
        return Ok(LabelAngle::HORIZONTAL);
    };
    LabelAngle::new(f32::from_bits(bits)).map_err(|error| {
        Error::InvalidFormat(format!(
            "native {}-axis label angle is invalid: {error}",
            axis.as_str()
        ))
    })
}

fn patch_axis_label_angle(data: &[u8], axis: Axis, expected: LabelAngle) -> Result<Vec<u8>> {
    let existing_extension = generated_axis_style_extension(data)?;
    if existing_extension.is_none() && expected == LabelAngle::HORIZONTAL {
        return Ok(data.to_vec());
    }
    let extension = existing_extension.unwrap_or_default();
    tsch::generated::ChartAxisStyleArchive::decode(extension)?;
    let field = label_angle_field(axis);
    let present = strict_optional_fixed32(extension, field)?.is_some();
    let replacement = (expected != LabelAngle::HORIZONTAL).then_some(expected.degrees().to_bits());
    let patched_extension = patch_fixed32_field(extension, field, present, replacement)?;
    let patched = patch_length_delimited_field(
        data,
        GENERATED_CHART_AXIS_STYLE_EXTENSION_FIELD,
        existing_extension.is_some(),
        (!patched_extension.is_empty()).then_some(patched_extension.as_slice()),
    )?;
    if read_axis_label_angle(&patched, axis)? != expected {
        return Err(Error::InvalidFormat(format!(
            "{}-axis label-angle wire patch failed validation",
            axis.as_str()
        )));
    }
    Ok(patched)
}

const fn label_angle_field(axis: Axis) -> u32 {
    match axis {
        Axis::Category => CATEGORY_LABEL_ANGLE_FIELD,
        Axis::Value => VALUE_LABEL_ANGLE_FIELD,
    }
}

fn strict_optional_fixed32(data: &[u8], field_number: u32) -> Result<Option<u32>> {
    let fields = parse_wire_fields(data)?;
    let mut matches = fields.iter().filter(|field| field.number() == field_number);
    let Some(field) = matches.next() else {
        return Ok(None);
    };
    if matches.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "singular chart axis label-angle field {field_number} occurs more than once"
        )));
    }
    if field.wire_type() != 5 {
        return Err(Error::InvalidFormat(format!(
            "chart axis label-angle field {field_number} is not fixed32"
        )));
    }
    let bytes: [u8; 4] = data[field.payload_start()..field.end()]
        .try_into()
        .map_err(|_| Error::InvalidFormat("truncated chart axis label angle".to_owned()))?;
    Ok(Some(u32::from_le_bytes(bytes)))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::protobuf::tss;
    use crate::wire::{append_length_delimited_field, append_varint_field};

    const UNKNOWN_OUTER_FIELD: u32 = 4_096;
    const UNKNOWN_GENERATED_FIELD: u32 = 4_097;

    #[test]
    fn named_and_custom_angles_are_strict_and_normalized() {
        for angle in [
            LabelAngle::HORIZONTAL,
            LabelAngle::LEFT_DIAGONAL,
            LabelAngle::LEFT_VERTICAL,
            LabelAngle::RIGHT_VERTICAL,
            LabelAngle::RIGHT_DIAGONAL,
            LabelAngle::new(12.5).unwrap(),
        ] {
            assert_eq!(LabelAngle::new(angle.degrees()).unwrap(), angle);
        }
        assert_eq!(
            LabelAngle::new(-0.0).unwrap().degrees().to_bits(),
            0.0_f32.to_bits()
        );
        for invalid in [f32::NAN, f32::INFINITY, f32::NEG_INFINITY, -0.1, 360.0] {
            assert!(LabelAngle::new(invalid).is_err());
        }
    }

    #[test]
    fn axis_angles_patch_independently_and_restore_exactly() {
        let original = axis_style_with_unknown_fields();
        let category =
            patch_axis_label_angle(&original, Axis::Category, LabelAngle::LEFT_DIAGONAL).unwrap();
        let both =
            patch_axis_label_angle(&category, Axis::Value, LabelAngle::RIGHT_DIAGONAL).unwrap();
        assert_eq!(
            read_axis_label_angle(&both, Axis::Category).unwrap(),
            LabelAngle::LEFT_DIAGONAL
        );
        assert_eq!(
            read_axis_label_angle(&both, Axis::Value).unwrap(),
            LabelAngle::RIGHT_DIAGONAL
        );
        let category_reset =
            patch_axis_label_angle(&both, Axis::Category, LabelAngle::HORIZONTAL).unwrap();
        let reset =
            patch_axis_label_angle(&category_reset, Axis::Value, LabelAngle::HORIZONTAL).unwrap();
        assert_eq!(reset, original);
    }

    #[test]
    fn malformed_native_angles_are_rejected() {
        let original = axis_style_with_unknown_fields();
        let customized =
            patch_axis_label_angle(&original, Axis::Value, LabelAngle::LEFT_VERTICAL).unwrap();
        let extension = generated_axis_style_extension(&customized)
            .unwrap()
            .unwrap();
        let duplicate_field = patch_fixed32_field(
            &[],
            VALUE_LABEL_ANGLE_FIELD,
            false,
            Some(LabelAngle::RIGHT_VERTICAL.degrees().to_bits()),
        )
        .unwrap();
        let mut duplicate_extension = extension.to_vec();
        duplicate_extension.extend_from_slice(&duplicate_field);
        let duplicate = patch_length_delimited_field(
            &customized,
            GENERATED_CHART_AXIS_STYLE_EXTENSION_FIELD,
            true,
            Some(duplicate_extension.as_slice()),
        )
        .unwrap();
        assert!(read_axis_label_angle(&duplicate, Axis::Value).is_err());

        let mut invalid_extension = extension.to_vec();
        let present = strict_optional_fixed32(&invalid_extension, VALUE_LABEL_ANGLE_FIELD)
            .unwrap()
            .is_some();
        invalid_extension = patch_fixed32_field(
            &invalid_extension,
            VALUE_LABEL_ANGLE_FIELD,
            present,
            Some(f32::NAN.to_bits()),
        )
        .unwrap();
        let invalid = patch_length_delimited_field(
            &customized,
            GENERATED_CHART_AXIS_STYLE_EXTENSION_FIELD,
            true,
            Some(invalid_extension.as_slice()),
        )
        .unwrap();
        assert!(read_axis_label_angle(&invalid, Axis::Value).is_err());
    }

    fn axis_style_with_unknown_fields() -> Vec<u8> {
        let mut generated = tsch::generated::ChartAxisStyleArchive {
            tschchartaxisvalueshowaxis: Some(true),
            ..Default::default()
        }
        .encode_to_vec();
        append_varint_field(&mut generated, UNKNOWN_GENERATED_FIELD, 73).unwrap();
        let mut outer = tsch::ChartAxisStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();
        append_length_delimited_field(
            &mut outer,
            GENERATED_CHART_AXIS_STYLE_EXTENSION_FIELD,
            &generated,
        )
        .unwrap();
        append_varint_field(&mut outer, UNKNOWN_OUTER_FIELD, 91).unwrap();
        outer
    }
}
