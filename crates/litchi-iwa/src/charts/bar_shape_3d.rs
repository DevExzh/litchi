//! Lossless native 3D bar-shape storage and mutation.
//!
//! Column and bar charts expose `Rectangle` and `Cylinder` in the 3D Scene
//! inspector. iWork stores that selection as a generated chart non-style
//! integer while this module exposes a closed enum and preserves unrelated
//! protobuf fields byte-for-byte.

use prost::Message;

use crate::charts::non_style::{
    GENERATED_CHART_NON_STYLE_EXTENSION_FIELD, chart_non_style_slot,
    generated_chart_non_style_extension,
};
use crate::protobuf::tsch;
use crate::wire::{patch_length_delimited_field, patch_varint_field};
use crate::{Error, IWorkPackage, Result};

/// `tschchartinfodefault3dbarshape` in the generated chart non-style archive.
const CHART_3D_BAR_SHAPE_FIELD: u32 = 1;
const RECTANGLE_NATIVE_VALUE: i32 = 0;
const CYLINDER_NATIVE_VALUE: i32 = 1;

/// Geometry used for bars or columns in a native 3D chart.
#[derive(Debug, Default, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Chart3dBarShape {
    /// Rectangular prisms, the native default.
    #[default]
    Rectangle,
    /// Circular cylinders.
    Cylinder,
}

impl Chart3dBarShape {
    const fn native(self) -> i32 {
        match self {
            Self::Rectangle => RECTANGLE_NATIVE_VALUE,
            Self::Cylinder => CYLINDER_NATIVE_VALUE,
        }
    }

    fn from_native(value: i32) -> Result<Self> {
        match value {
            RECTANGLE_NATIVE_VALUE => Ok(Self::Rectangle),
            CYLINDER_NATIVE_VALUE => Ok(Self::Cylinder),
            _ => Err(Error::InvalidFormat(format!(
                "unsupported native chart 3D bar shape {value}"
            ))),
        }
    }
}

/// Read one chart's effective native 3D bar shape.
pub(crate) fn chart_3d_bar_shape(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<Chart3dBarShape> {
    chart_non_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?
    .read(package, read_chart_3d_bar_shape)
}

/// Set one chart's native 3D bar shape.
pub(crate) fn set_chart_3d_bar_shape(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    shape: Chart3dBarShape,
) -> Result<()> {
    let slot = chart_non_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    if slot.read(package, read_chart_3d_bar_shape)? == shape {
        return Ok(());
    }
    slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
    slot.update(package, |data| patch_chart_3d_bar_shape(data, shape))?;
    if slot.read(package, read_chart_3d_bar_shape)? != shape {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} 3D bar-shape update failed validation"
        )));
    }
    Ok(())
}

fn read_chart_3d_bar_shape(data: &[u8]) -> Result<Chart3dBarShape> {
    let Some(extension) = generated_chart_non_style_extension(data)? else {
        return Ok(Chart3dBarShape::Rectangle);
    };
    let generated = tsch::generated::ChartNonStyleArchive::decode(extension)?;
    generated
        .tschchartinfodefault3dbarshape
        .map(Chart3dBarShape::from_native)
        .transpose()
        .map(|shape| shape.unwrap_or_default())
}

fn patch_chart_3d_bar_shape(data: &[u8], shape: Chart3dBarShape) -> Result<Vec<u8>> {
    let Some(extension) = generated_chart_non_style_extension(data)? else {
        if shape == Chart3dBarShape::Rectangle {
            return Ok(data.to_vec());
        }
        let generated = tsch::generated::ChartNonStyleArchive {
            tschchartinfodefault3dbarshape: Some(shape.native()),
            ..Default::default()
        };
        let encoded = generated.encode_to_vec();
        let patched = patch_length_delimited_field(
            data,
            GENERATED_CHART_NON_STYLE_EXTENSION_FIELD,
            false,
            Some(encoded.as_slice()),
        )?;
        validate_patched_chart_3d_bar_shape(&patched, shape)?;
        return Ok(patched);
    };

    let generated = tsch::generated::ChartNonStyleArchive::decode(extension)?;
    let present = generated.tschchartinfodefault3dbarshape.is_some();
    let native = (shape != Chart3dBarShape::Rectangle)
        .then_some(u64::from(shape == Chart3dBarShape::Cylinder));
    let extension = patch_varint_field(extension, CHART_3D_BAR_SHAPE_FIELD, present, native)?;
    let patched = patch_length_delimited_field(
        data,
        GENERATED_CHART_NON_STYLE_EXTENSION_FIELD,
        true,
        Some(extension.as_slice()),
    )?;
    validate_patched_chart_3d_bar_shape(&patched, shape)?;
    Ok(patched)
}

fn validate_patched_chart_3d_bar_shape(data: &[u8], expected: Chart3dBarShape) -> Result<()> {
    if read_chart_3d_bar_shape(data)? != expected {
        return Err(Error::InvalidFormat(
            "chart 3D bar-shape wire patch failed validation".to_owned(),
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
    const UNKNOWN_VALUE: u64 = 42;

    #[test]
    fn bar_shapes_are_closed_and_rectangle_is_the_native_default() {
        assert_eq!(Chart3dBarShape::default(), Chart3dBarShape::Rectangle);
        assert_eq!(
            Chart3dBarShape::from_native(RECTANGLE_NATIVE_VALUE).unwrap(),
            Chart3dBarShape::Rectangle
        );
        assert_eq!(
            Chart3dBarShape::from_native(CYLINDER_NATIVE_VALUE).unwrap(),
            Chart3dBarShape::Cylinder
        );
        assert!(Chart3dBarShape::from_native(-1).is_err());
        assert!(Chart3dBarShape::from_native(2).is_err());
    }

    #[test]
    fn bar_shape_patch_preserves_other_fields_and_unknown_data() {
        let mut extension = tsch::generated::ChartNonStyleArchive {
            tschchartinfodefault3dbarshape: Some(RECTANGLE_NATIVE_VALUE),
            tschchartinfodefault3dbeveledges: Some(true),
            tschchartinfodefaultshowlegend: Some(true),
            ..Default::default()
        }
        .encode_to_vec();
        append_varint_field(&mut extension, UNKNOWN_EXTENSION_FIELD, UNKNOWN_VALUE).unwrap();
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
        append_varint_field(&mut original, UNKNOWN_OUTER_FIELD, UNKNOWN_VALUE).unwrap();

        let cylinder = patch_chart_3d_bar_shape(&original, Chart3dBarShape::Cylinder).unwrap();
        assert_eq!(
            read_chart_3d_bar_shape(&cylinder).unwrap(),
            Chart3dBarShape::Cylinder
        );
        assert!(has_field(&cylinder, UNKNOWN_OUTER_FIELD));
        let extension = generated_chart_non_style_extension(&cylinder)
            .unwrap()
            .unwrap();
        assert!(has_field(extension, UNKNOWN_EXTENSION_FIELD));
        let decoded = tsch::generated::ChartNonStyleArchive::decode(extension).unwrap();
        assert_eq!(decoded.tschchartinfodefault3dbeveledges, Some(true));
        assert_eq!(decoded.tschchartinfodefaultshowlegend, Some(true));

        let rectangle = patch_chart_3d_bar_shape(&cylinder, Chart3dBarShape::Rectangle).unwrap();
        assert_eq!(
            read_chart_3d_bar_shape(&rectangle).unwrap(),
            Chart3dBarShape::Rectangle
        );
        assert_eq!(
            tsch::generated::ChartNonStyleArchive::decode(
                generated_chart_non_style_extension(&rectangle)
                    .unwrap()
                    .unwrap()
            )
            .unwrap()
            .tschchartinfodefault3dbarshape,
            None
        );
    }

    #[test]
    fn malformed_native_bar_shape_is_rejected() {
        let generated = tsch::generated::ChartNonStyleArchive {
            tschchartinfodefault3dbarshape: Some(2),
            ..Default::default()
        };
        let mut outer = tsch::ChartNonStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();
        append_length_delimited_field(
            &mut outer,
            GENERATED_CHART_NON_STYLE_EXTENSION_FIELD,
            &generated.encode_to_vec(),
        )
        .unwrap();
        assert!(read_chart_3d_bar_shape(&outer).is_err());
    }

    fn has_field(data: &[u8], field_number: u32) -> bool {
        parse_wire_fields(data)
            .unwrap()
            .iter()
            .any(|field| field.number() == field_number)
    }
}
