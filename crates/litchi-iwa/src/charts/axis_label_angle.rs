//! Lossless label-angle CRUD for native chart axes.
//!
//! Pages, Numbers, and Keynote persist a normalized degree value as a
//! protobuf float in the selected axis style's generated extension.

use prost::Message;

use crate::charts::Axis;
use crate::charts::axis_style::{
    GENERATED_CHART_AXIS_STYLE_EXTENSION_FIELD, axis_style_slot, generated_axis_style_extension,
};
use crate::protobuf::tsch;
use crate::wire::{parse_wire_fields, patch_fixed32_field, patch_length_delimited_field};
use crate::{Error, IWorkPackage, Result};

const CATEGORY_LABEL_ANGLE_FIELD: u32 = 9;
const VALUE_LABEL_ANGLE_FIELD: u32 = 11;
const MINIMUM_DEGREES: f32 = 0.0;
const MAXIMUM_DEGREES_EXCLUSIVE: f32 = 360.0;

/// A normalized chart-axis label angle in degrees.
///
/// The associated constants correspond to the five named choices in the
/// iWork Label Angle pop-up. Other finite values in `[0, 360)` represent the
/// inspector's Custom choice.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ChartAxisLabelAngle(f32);

impl ChartAxisLabelAngle {
    /// Horizontal labels.
    pub const HORIZONTAL: Self = Self(MINIMUM_DEGREES);
    /// Labels rising diagonally toward the left.
    pub const LEFT_DIAGONAL: Self = Self(45.0);
    /// Labels written vertically toward the left.
    pub const LEFT_VERTICAL: Self = Self(90.0);
    /// Labels written vertically toward the right.
    pub const RIGHT_VERTICAL: Self = Self(270.0);
    /// Labels rising diagonally toward the right.
    pub const RIGHT_DIAGONAL: Self = Self(315.0);

    /// Construct a normalized native label angle.
    pub fn new(degrees: f32) -> Result<Self> {
        if !degrees.is_finite() || !(MINIMUM_DEGREES..MAXIMUM_DEGREES_EXCLUSIVE).contains(&degrees)
        {
            return Err(Error::InvalidFormat(format!(
                "chart axis label angle must be finite and in [{MINIMUM_DEGREES}, {MAXIMUM_DEGREES_EXCLUSIVE}) degrees"
            )));
        }
        // Canonicalize negative zero so equality and wire output are stable.
        Ok(Self(if degrees == MINIMUM_DEGREES {
            MINIMUM_DEGREES
        } else {
            degrees
        }))
    }

    /// Return the normalized angle in degrees.
    pub const fn degrees(self) -> f32 {
        self.0
    }
}

impl Default for ChartAxisLabelAngle {
    fn default() -> Self {
        Self::HORIZONTAL
    }
}

impl TryFrom<f32> for ChartAxisLabelAngle {
    type Error = Error;

    fn try_from(value: f32) -> Result<Self> {
        Self::new(value)
    }
}

/// Read the label angle for one native chart axis.
pub(crate) fn chart_axis_label_angle(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    axis: Axis,
) -> Result<ChartAxisLabelAngle> {
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
    angle: ChartAxisLabelAngle,
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

fn read_axis_label_angle(data: &[u8], axis: Axis) -> Result<ChartAxisLabelAngle> {
    let Some(extension) = generated_axis_style_extension(data)? else {
        return Ok(ChartAxisLabelAngle::HORIZONTAL);
    };
    tsch::generated::ChartAxisStyleArchive::decode(extension)?;
    let Some(bits) = strict_optional_fixed32(extension, label_angle_field(axis))? else {
        return Ok(ChartAxisLabelAngle::HORIZONTAL);
    };
    ChartAxisLabelAngle::new(f32::from_bits(bits)).map_err(|error| {
        Error::InvalidFormat(format!(
            "native {}-axis label angle is invalid: {error}",
            axis.as_str()
        ))
    })
}

fn patch_axis_label_angle(
    data: &[u8],
    axis: Axis,
    expected: ChartAxisLabelAngle,
) -> Result<Vec<u8>> {
    let existing_extension = generated_axis_style_extension(data)?;
    if existing_extension.is_none() && expected == ChartAxisLabelAngle::HORIZONTAL {
        return Ok(data.to_vec());
    }
    let extension = existing_extension.unwrap_or_default();
    tsch::generated::ChartAxisStyleArchive::decode(extension)?;
    let field = label_angle_field(axis);
    let present = strict_optional_fixed32(extension, field)?.is_some();
    let replacement =
        (expected != ChartAxisLabelAngle::HORIZONTAL).then_some(expected.degrees().to_bits());
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
            ChartAxisLabelAngle::HORIZONTAL,
            ChartAxisLabelAngle::LEFT_DIAGONAL,
            ChartAxisLabelAngle::LEFT_VERTICAL,
            ChartAxisLabelAngle::RIGHT_VERTICAL,
            ChartAxisLabelAngle::RIGHT_DIAGONAL,
            ChartAxisLabelAngle::new(12.5).unwrap(),
        ] {
            assert_eq!(ChartAxisLabelAngle::new(angle.degrees()).unwrap(), angle);
        }
        assert_eq!(
            ChartAxisLabelAngle::new(-0.0).unwrap().degrees().to_bits(),
            0.0_f32.to_bits()
        );
        for invalid in [f32::NAN, f32::INFINITY, f32::NEG_INFINITY, -0.1, 360.0] {
            assert!(ChartAxisLabelAngle::new(invalid).is_err());
        }
    }

    #[test]
    fn axis_angles_patch_independently_and_restore_exactly() {
        let original = axis_style_with_unknown_fields();
        let category = patch_axis_label_angle(
            &original,
            Axis::Category,
            ChartAxisLabelAngle::LEFT_DIAGONAL,
        )
        .unwrap();
        let both =
            patch_axis_label_angle(&category, Axis::Value, ChartAxisLabelAngle::RIGHT_DIAGONAL)
                .unwrap();
        assert_eq!(
            read_axis_label_angle(&both, Axis::Category).unwrap(),
            ChartAxisLabelAngle::LEFT_DIAGONAL
        );
        assert_eq!(
            read_axis_label_angle(&both, Axis::Value).unwrap(),
            ChartAxisLabelAngle::RIGHT_DIAGONAL
        );
        let category_reset =
            patch_axis_label_angle(&both, Axis::Category, ChartAxisLabelAngle::HORIZONTAL).unwrap();
        let reset = patch_axis_label_angle(
            &category_reset,
            Axis::Value,
            ChartAxisLabelAngle::HORIZONTAL,
        )
        .unwrap();
        assert_eq!(reset, original);
    }

    #[test]
    fn malformed_native_angles_are_rejected() {
        let original = axis_style_with_unknown_fields();
        let customized =
            patch_axis_label_angle(&original, Axis::Value, ChartAxisLabelAngle::LEFT_VERTICAL)
                .unwrap();
        let extension = generated_axis_style_extension(&customized)
            .unwrap()
            .unwrap();
        let duplicate_field = patch_fixed32_field(
            &[],
            VALUE_LABEL_ANGLE_FIELD,
            false,
            Some(ChartAxisLabelAngle::RIGHT_VERTICAL.degrees().to_bits()),
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
