//! Native 3D bar-shape storage and mutation.
//!
//! Column and bar charts expose `Rectangle` and `Cylinder` in the 3D Scene
//! inspector. iWork stores that selection as a generated chart non-style
//! integer while this module preserves unrelated protobuf fields byte-for-byte.

use litchi_iwa_common::chart3d::BarShape;
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

/// Read one chart's effective native 3D bar shape.
pub(crate) fn chart_3d_bar_shape(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<BarShape> {
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
    shape: BarShape,
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

fn read_chart_3d_bar_shape(data: &[u8]) -> Result<BarShape> {
    let Some(extension) = generated_chart_non_style_extension(data)? else {
        return Ok(BarShape::Rectangle);
    };
    let generated = tsch::generated::ChartNonStyleArchive::decode(extension)?;
    Ok(generated
        .tschchartinfodefault3dbarshape
        .map(BarShape::from_native)
        .unwrap_or_default())
}

fn patch_chart_3d_bar_shape(data: &[u8], shape: BarShape) -> Result<Vec<u8>> {
    let Some(extension) = generated_chart_non_style_extension(data)? else {
        if shape == BarShape::Rectangle {
            return Ok(data.to_vec());
        }
        let generated = tsch::generated::ChartNonStyleArchive {
            tschchartinfodefault3dbarshape: Some(shape.native_value()),
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
    let native =
        (shape != BarShape::Rectangle).then_some(i64::from(shape.native_value()).cast_unsigned());
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

fn validate_patched_chart_3d_bar_shape(data: &[u8], expected: BarShape) -> Result<()> {
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
    fn bar_shapes_use_known_values_and_preserve_unknown_values() {
        assert_eq!(BarShape::default(), BarShape::Rectangle);
        assert_eq!(
            BarShape::from_native(BarShape::Rectangle.native_value()),
            BarShape::Rectangle
        );
        assert_eq!(
            BarShape::from_native(BarShape::Cylinder.native_value()),
            BarShape::Cylinder
        );
        assert!(BarShape::from_native(-1).is_unsupported());
        assert!(BarShape::from_native(2).is_unsupported());
    }

    #[test]
    fn bar_shape_patch_preserves_other_fields_and_unknown_data() {
        let mut extension = tsch::generated::ChartNonStyleArchive {
            tschchartinfodefault3dbarshape: Some(BarShape::Rectangle.native_value()),
            tschchartinfodefault3dbeveledges: Some(true),
            tschchartinfodefaultshowlegend: Some(true),
            ..Default::default()
        }
        .encode_to_vec();
        append_varint_field(&mut extension, UNKNOWN_EXTENSION_FIELD, UNKNOWN_VALUE).unwrap();
        let original_extension = extension.clone();
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

        let unknown = BarShape::from_native(9_001);
        let patched = patch_chart_3d_bar_shape(&original, unknown).unwrap();
        assert_eq!(read_chart_3d_bar_shape(&patched).unwrap(), unknown);
        assert_eq!(
            field_raw(&patched, UNKNOWN_OUTER_FIELD),
            field_raw(&original, UNKNOWN_OUTER_FIELD)
        );
        let extension = generated_chart_non_style_extension(&patched)
            .unwrap()
            .unwrap();
        assert_eq!(
            field_raw(extension, UNKNOWN_EXTENSION_FIELD),
            field_raw(&original_extension, UNKNOWN_EXTENSION_FIELD)
        );
        let decoded = tsch::generated::ChartNonStyleArchive::decode(extension).unwrap();
        assert_eq!(decoded.tschchartinfodefault3dbeveledges, Some(true));
        assert_eq!(decoded.tschchartinfodefaultshowlegend, Some(true));

        let rectangle = patch_chart_3d_bar_shape(&patched, BarShape::Rectangle).unwrap();
        assert_eq!(
            read_chart_3d_bar_shape(&rectangle).unwrap(),
            BarShape::Rectangle
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
    fn unknown_native_bar_shape_is_preserved() {
        let generated = tsch::generated::ChartNonStyleArchive {
            tschchartinfodefault3dbarshape: Some(-7),
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
        assert_eq!(
            read_chart_3d_bar_shape(&outer).unwrap(),
            BarShape::from_native(-7)
        );
        let patched = patch_chart_3d_bar_shape(&outer, BarShape::from_native(-7)).unwrap();
        assert_eq!(
            read_chart_3d_bar_shape(&patched).unwrap(),
            BarShape::from_native(-7)
        );
    }

    fn field_raw(data: &[u8], field_number: u32) -> Vec<u8> {
        parse_wire_fields(data)
            .unwrap()
            .into_iter()
            .find(|field| field.number() == field_number)
            .unwrap()
            .raw(data)
            .unwrap()
            .to_vec()
    }
}
