//! Lossless native 3D chart-scene rotation storage and mutation.
//!
//! The Chart inspector exposes independent X/Y rotation controls for 3D
//! charts. iWork stores both angles in the generated chart non-style
//! extension as a four-component `TSCH.Chart3DVectorArchive`; Z/W are retained
//! byte-for-byte when an existing vector is updated.

use prost::Message;

use crate::charts::non_style::{
    GENERATED_CHART_NON_STYLE_EXTENSION_FIELD, chart_non_style_slot,
    generated_chart_non_style_extension,
};
use crate::protobuf::tsch;
use crate::wire::{
    patch_fixed32_field, patch_length_delimited_field, transform_length_delimited_field,
};
use crate::{Error, IWorkPackage, Result};

/// `tschchartinfodefault3drotation` in the generated chart non-style archive.
const CHART_3D_ROTATION_FIELD: u32 = 4;
const VECTOR_X_FIELD: u32 = 1;
const VECTOR_Y_FIELD: u32 = 2;
const MINIMUM_ROTATION_DEGREES: f32 = -40.0;
const MAXIMUM_ROTATION_DEGREES: f32 = 40.0;
const DEFAULT_X_DEGREES: f32 = 18.5;
const DEFAULT_Y_DEGREES: f32 = -18.125;

/// X/Y orientation of a native 3D chart scene, in inspector degrees.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Chart3dRotation {
    x_degrees: f32,
    y_degrees: f32,
}

impl Chart3dRotation {
    /// Orientation used by iWork when no direct rotation vector is stored.
    pub const DEFAULT: Self = Self {
        x_degrees: DEFAULT_X_DEGREES,
        y_degrees: DEFAULT_Y_DEGREES,
    };

    /// Construct a 3D scene orientation in inspector degrees.
    ///
    /// Both axes must be finite and within the native inspector's inclusive
    /// `-40°..=40°` range.
    pub fn from_degrees(x_degrees: f32, y_degrees: f32) -> Result<Self> {
        if !valid_rotation(x_degrees) || !valid_rotation(y_degrees) {
            return Err(Error::InvalidFormat(format!(
                "chart 3D rotation angles must be finite and within {MINIMUM_ROTATION_DEGREES}°..={MAXIMUM_ROTATION_DEGREES}°"
            )));
        }
        Ok(Self {
            x_degrees,
            y_degrees,
        })
    }

    /// Return the rotation about the scene's X axis, in degrees.
    pub const fn x_degrees(self) -> f32 {
        self.x_degrees
    }

    /// Return the rotation about the scene's Y axis, in degrees.
    pub const fn y_degrees(self) -> f32 {
        self.y_degrees
    }
}

fn valid_rotation(degrees: f32) -> bool {
    degrees.is_finite() && (MINIMUM_ROTATION_DEGREES..=MAXIMUM_ROTATION_DEGREES).contains(&degrees)
}

impl Default for Chart3dRotation {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// Read one chart's effective native 3D scene rotation.
pub(crate) fn chart_3d_rotation(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<Chart3dRotation> {
    chart_non_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?
    .read(package, read_chart_3d_rotation)
}

/// Set one chart's native 3D scene rotation.
pub(crate) fn set_chart_3d_rotation(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    rotation: Chart3dRotation,
) -> Result<()> {
    let slot = chart_non_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    if slot.read(package, read_chart_3d_rotation)? == rotation {
        return Ok(());
    }
    slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
    slot.update(package, |data| patch_chart_3d_rotation(data, rotation))?;
    if slot.read(package, read_chart_3d_rotation)? != rotation {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} 3D rotation update failed validation"
        )));
    }
    Ok(())
}

fn read_chart_3d_rotation(data: &[u8]) -> Result<Chart3dRotation> {
    let Some(extension) = generated_chart_non_style_extension(data)? else {
        return Ok(Chart3dRotation::DEFAULT);
    };
    let generated = tsch::generated::ChartNonStyleArchive::decode(extension)?;
    let Some(rotation) = generated.tschchartinfodefault3drotation else {
        return Ok(Chart3dRotation::DEFAULT);
    };
    if !rotation.z.is_finite() || !rotation.w.is_finite() {
        return Err(Error::InvalidFormat(
            "native chart 3D rotation Z/W components must be finite".to_owned(),
        ));
    }
    Chart3dRotation::from_degrees(rotation.x, rotation.y)
}

fn patch_chart_3d_rotation(data: &[u8], rotation: Chart3dRotation) -> Result<Vec<u8>> {
    let Some(extension) = generated_chart_non_style_extension(data)? else {
        if rotation == Chart3dRotation::DEFAULT {
            return Ok(data.to_vec());
        }
        let generated = tsch::generated::ChartNonStyleArchive {
            tschchartinfodefault3drotation: Some(native_rotation(rotation)),
            ..Default::default()
        };
        let patched = patch_length_delimited_field(
            data,
            GENERATED_CHART_NON_STYLE_EXTENSION_FIELD,
            false,
            Some(generated.encode_to_vec().as_slice()),
        )?;
        validate_patched_chart_3d_rotation(&patched, rotation)?;
        return Ok(patched);
    };

    let generated = tsch::generated::ChartNonStyleArchive::decode(extension)?;
    let present = generated.tschchartinfodefault3drotation.is_some();
    let extension = match (present, rotation == Chart3dRotation::DEFAULT) {
        (false, true) => extension.to_vec(),
        (false, false) => patch_length_delimited_field(
            extension,
            CHART_3D_ROTATION_FIELD,
            false,
            Some(native_rotation(rotation).encode_to_vec().as_slice()),
        )?,
        (true, true) => {
            patch_length_delimited_field(extension, CHART_3D_ROTATION_FIELD, true, None)?
        },
        (true, false) => {
            transform_length_delimited_field(extension, CHART_3D_ROTATION_FIELD, |vector| {
                let vector = patch_fixed32_field(
                    vector,
                    VECTOR_X_FIELD,
                    true,
                    Some(rotation.x_degrees.to_bits()),
                )?;
                patch_fixed32_field(
                    &vector,
                    VECTOR_Y_FIELD,
                    true,
                    Some(rotation.y_degrees.to_bits()),
                )
            })?
        },
    };
    let patched = patch_length_delimited_field(
        data,
        GENERATED_CHART_NON_STYLE_EXTENSION_FIELD,
        true,
        Some(extension.as_slice()),
    )?;
    validate_patched_chart_3d_rotation(&patched, rotation)?;
    Ok(patched)
}

const fn native_rotation(rotation: Chart3dRotation) -> tsch::Chart3DVectorArchive {
    tsch::Chart3DVectorArchive {
        x: rotation.x_degrees,
        y: rotation.y_degrees,
        z: 0.0,
        w: 0.0,
    }
}

fn validate_patched_chart_3d_rotation(data: &[u8], expected: Chart3dRotation) -> Result<()> {
    if read_chart_3d_rotation(data)? != expected {
        return Err(Error::InvalidFormat(
            "chart 3D rotation wire patch failed validation".to_owned(),
        ));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::protobuf::tss;
    use crate::wire::{append_length_delimited_field, append_varint_field, parse_wire_fields};

    const UNKNOWN_OUTER_FIELD: u32 = 4_096;
    const UNKNOWN_EXTENSION_FIELD: u32 = 4_097;
    const UNKNOWN_VECTOR_FIELD: u32 = 4_098;

    #[test]
    fn rotations_are_finite_and_have_the_native_default() {
        assert_eq!(Chart3dRotation::default(), Chart3dRotation::DEFAULT);
        assert_eq!(Chart3dRotation::DEFAULT.x_degrees(), 18.5);
        assert_eq!(Chart3dRotation::DEFAULT.y_degrees(), -18.125);
        assert!(Chart3dRotation::from_degrees(f32::NAN, 0.0).is_err());
        assert!(Chart3dRotation::from_degrees(0.0, f32::INFINITY).is_err());
        assert!(Chart3dRotation::from_degrees(-40.001, 0.0).is_err());
        assert!(Chart3dRotation::from_degrees(0.0, 40.001).is_err());
        assert!(Chart3dRotation::from_degrees(-40.0, 40.0).is_ok());
    }

    #[test]
    fn rotation_patch_preserves_unknown_outer_extension_and_vector_fields() {
        let mut vector = native_rotation(Chart3dRotation::DEFAULT).encode_to_vec();
        append_varint_field(&mut vector, UNKNOWN_VECTOR_FIELD, 42).unwrap();
        let mut extension = tsch::generated::ChartNonStyleArchive {
            tschchartinfodefault3drotation: Some(native_rotation(Chart3dRotation::DEFAULT)),
            tschchartinfodefaultshowlegend: Some(true),
            ..Default::default()
        }
        .encode_to_vec();
        extension =
            patch_length_delimited_field(&extension, CHART_3D_ROTATION_FIELD, true, Some(&vector))
                .unwrap();
        append_varint_field(&mut extension, UNKNOWN_EXTENSION_FIELD, 43).unwrap();
        let mut original = tsch::ChartNonStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();
        append_length_delimited_field(
            &mut original,
            GENERATED_CHART_NON_STYLE_EXTENSION_FIELD,
            &extension,
        )
        .unwrap();
        append_varint_field(&mut original, UNKNOWN_OUTER_FIELD, 44).unwrap();

        let replacement = Chart3dRotation::from_degrees(30.0, -40.0).unwrap();
        let patched = patch_chart_3d_rotation(&original, replacement).unwrap();
        assert_eq!(read_chart_3d_rotation(&patched).unwrap(), replacement);
        assert!(has_field(&patched, UNKNOWN_OUTER_FIELD));
        let extension = generated_chart_non_style_extension(&patched)
            .unwrap()
            .unwrap();
        assert!(has_field(extension, UNKNOWN_EXTENSION_FIELD));
        let vector = parse_wire_fields(extension)
            .unwrap()
            .into_iter()
            .find(|field| field.number == CHART_3D_ROTATION_FIELD)
            .unwrap();
        assert!(has_field(
            &extension[vector.payload_start..vector.end],
            UNKNOWN_VECTOR_FIELD
        ));
        assert_eq!(
            patch_chart_3d_rotation(&patched, Chart3dRotation::DEFAULT).unwrap(),
            {
                let extension = patch_length_delimited_field(
                    generated_chart_non_style_extension(&original)
                        .unwrap()
                        .unwrap(),
                    CHART_3D_ROTATION_FIELD,
                    true,
                    None,
                )
                .unwrap();
                patch_length_delimited_field(
                    &original,
                    GENERATED_CHART_NON_STYLE_EXTENSION_FIELD,
                    true,
                    Some(&extension),
                )
                .unwrap()
            }
        );
    }

    #[test]
    fn malformed_rotation_vectors_are_rejected() {
        let mut vector = native_rotation(Chart3dRotation::DEFAULT).encode_to_vec();
        vector =
            patch_fixed32_field(&vector, VECTOR_X_FIELD, true, Some(f32::NAN.to_bits())).unwrap();
        let generated = tsch::generated::ChartNonStyleArchive {
            tschchartinfodefault3drotation: Some(native_rotation(Chart3dRotation::DEFAULT)),
            ..Default::default()
        }
        .encode_to_vec();
        let generated =
            patch_length_delimited_field(&generated, CHART_3D_ROTATION_FIELD, true, Some(&vector))
                .unwrap();
        let mut outer = tsch::ChartNonStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();
        append_length_delimited_field(
            &mut outer,
            GENERATED_CHART_NON_STYLE_EXTENSION_FIELD,
            &generated,
        )
        .unwrap();
        assert!(read_chart_3d_rotation(&outer).is_err());
    }

    fn has_field(data: &[u8], field_number: u32) -> bool {
        parse_wire_fields(data)
            .unwrap()
            .iter()
            .any(|field| field.number == field_number)
    }
}
