//! Lossless native chart-border stroke storage and mutation.
//!
//! iWork stores the chart-area border's color, width, and pattern in the
//! generated extension of `TSCH.ChartStyleArchive`. The border visibility
//! switch is separate. This module preserves both protobuf layers and every
//! unrelated field byte-for-byte while changing only the native stroke.

use prost::Message;

use crate::charts::style::{
    GENERATED_CHART_STYLE_EXTENSION_FIELD, chart_style_slot, generated_chart_style_extension,
};
use crate::protobuf::tsch;
use crate::shapes::{
    RgbaColor, ShapeStroke, StrokePattern, StrokeWidth, empty_stroke_archive, stroke_from_native,
    stroke_to_native,
};
use crate::wire::patch_length_delimited_field;
use crate::{Error, IWorkPackage, Result};

/// `tschchartinfodefaultborderstroke` in `TSCH.Generated.ChartStyleArchive`.
const CHART_BORDER_STROKE_FIELD: u32 = 12;

/// The solid, black, one-point stroke used by a newly inserted native chart.
pub(crate) fn native_default_chart_border_stroke() -> ShapeStroke {
    ShapeStroke::new(RgbaColor::black(), StrokeWidth::ONE, StrokePattern::Solid)
}

/// Read the chart-area border stroke, or `None` when the native style is empty.
pub(crate) fn chart_border_stroke(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<Option<ShapeStroke>> {
    chart_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?
    .read(package, read_chart_border_stroke)
}

/// Set the chart-area border stroke independently of border visibility.
pub(crate) fn set_chart_border_stroke(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    stroke: Option<ShapeStroke>,
) -> Result<()> {
    let slot = chart_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    if slot.read(package, read_chart_border_stroke)? == stroke {
        return Ok(());
    }
    slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
    slot.update(package, |data| patch_chart_border_stroke(data, stroke))?;
    if slot.read(package, read_chart_border_stroke)? != stroke {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} border stroke update failed validation"
        )));
    }
    Ok(())
}

fn read_chart_border_stroke(data: &[u8]) -> Result<Option<ShapeStroke>> {
    let Some(extension) = generated_chart_style_extension(data)? else {
        return Ok(Some(native_default_chart_border_stroke()));
    };
    let generated = tsch::generated::ChartStyleArchive::decode(extension)?;
    generated
        .tschchartinfodefaultborderstroke
        .as_ref()
        .map(stroke_from_native)
        .transpose()
        .map(|stroke| stroke.unwrap_or_else(|| Some(native_default_chart_border_stroke())))
}

fn patch_chart_border_stroke(data: &[u8], stroke: Option<ShapeStroke>) -> Result<Vec<u8>> {
    let Some(extension) = generated_chart_style_extension(data)? else {
        if stroke == Some(native_default_chart_border_stroke()) {
            return Ok(data.to_vec());
        }
        let generated = tsch::generated::ChartStyleArchive {
            tschchartinfodefaultborderstroke: Some(native_stroke(stroke)),
            ..Default::default()
        };
        let encoded = generated.encode_to_vec();
        let patched = patch_length_delimited_field(
            data,
            GENERATED_CHART_STYLE_EXTENSION_FIELD,
            false,
            Some(encoded.as_slice()),
        )?;
        validate_patched_chart_border_stroke(&patched, stroke)?;
        return Ok(patched);
    };

    let generated = tsch::generated::ChartStyleArchive::decode(extension)?;
    let stroke_present = generated.tschchartinfodefaultborderstroke.is_some();
    let native = (stroke != Some(native_default_chart_border_stroke()))
        .then(|| native_stroke(stroke).encode_to_vec());
    let extension = patch_length_delimited_field(
        extension,
        CHART_BORDER_STROKE_FIELD,
        stroke_present,
        native.as_deref(),
    )?;
    let patched = patch_length_delimited_field(
        data,
        GENERATED_CHART_STYLE_EXTENSION_FIELD,
        true,
        Some(extension.as_slice()),
    )?;
    validate_patched_chart_border_stroke(&patched, stroke)?;
    Ok(patched)
}

fn native_stroke(stroke: Option<ShapeStroke>) -> crate::protobuf::tsd::StrokeArchive {
    stroke.map_or_else(empty_stroke_archive, stroke_to_native)
}

fn validate_patched_chart_border_stroke(data: &[u8], expected: Option<ShapeStroke>) -> Result<()> {
    if read_chart_border_stroke(data)? != expected {
        return Err(Error::InvalidFormat(
            "chart border stroke wire patch failed validation".to_owned(),
        ));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::protobuf::tss;
    use crate::shapes::RgbColorSpace;
    use crate::wire::{append_length_delimited_field, append_varint_field, parse_wire_fields};

    const UNMAPPED_OUTER_FIELD: u32 = 4_096;
    const UNMAPPED_GENERATED_FIELD: u32 = 4_097;
    const UNMAPPED_VALUE: u64 = 42;

    #[test]
    fn border_stroke_defaults_natively_and_creates_an_extension_when_needed() {
        let original = tsch::ChartStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();
        let replacement = test_stroke(StrokePattern::MediumDash, 3.0);

        assert_eq!(
            read_chart_border_stroke(&original).unwrap(),
            Some(native_default_chart_border_stroke())
        );
        assert_eq!(
            patch_chart_border_stroke(&original, Some(native_default_chart_border_stroke()))
                .unwrap(),
            original
        );

        let patched = patch_chart_border_stroke(&original, Some(replacement)).unwrap();
        assert_eq!(
            read_chart_border_stroke(&patched).unwrap(),
            Some(replacement)
        );
        assert!(generated_chart_style_extension(&patched).unwrap().is_some());

        let empty = patch_chart_border_stroke(&patched, None).unwrap();
        assert_eq!(read_chart_border_stroke(&empty).unwrap(), None);
    }

    #[test]
    fn border_stroke_patch_retains_other_style_fields_and_unmapped_data() {
        let original_stroke = test_stroke(StrokePattern::Solid, 2.0);
        let replacement = test_stroke(StrokePattern::RoundedDash, 4.0);
        let original = style_with_unknown_fields(tsch::generated::ChartStyleArchive {
            tschchartinfodefaultborderstroke: Some(stroke_to_native(original_stroke)),
            tschchartinfodefaultshowborder: Some(true),
            tschchartinfodefaultinterbargap: Some(25.0),
            tschchartinfodefaultroundedcornerradius: Some(0.2),
            ..Default::default()
        });

        let patched = patch_chart_border_stroke(&original, Some(replacement)).unwrap();
        assert_eq!(
            read_chart_border_stroke(&patched).unwrap(),
            Some(replacement)
        );
        let generated = tsch::generated::ChartStyleArchive::decode(
            generated_chart_style_extension(&patched).unwrap().unwrap(),
        )
        .unwrap();
        assert_eq!(generated.tschchartinfodefaultshowborder, Some(true));
        assert_eq!(generated.tschchartinfodefaultinterbargap, Some(25.0));
        assert_eq!(generated.tschchartinfodefaultroundedcornerradius, Some(0.2));
        assert_unknown_fields_retained(&original, &patched);

        let restored = patch_chart_border_stroke(&patched, Some(original_stroke)).unwrap();
        assert_eq!(restored, original);
    }

    #[test]
    fn resetting_border_stroke_retains_other_style_fields() {
        let original = style_with_unknown_fields(tsch::generated::ChartStyleArchive {
            tschchartinfodefaultborderstroke: Some(stroke_to_native(test_stroke(
                StrokePattern::MediumDash,
                3.0,
            ))),
            tschchartinfodefaultshowborder: Some(true),
            tschchartinfodefaultintersetgap: Some(70.0),
            ..Default::default()
        });

        let reset =
            patch_chart_border_stroke(&original, Some(native_default_chart_border_stroke()))
                .unwrap();
        assert_eq!(
            read_chart_border_stroke(&reset).unwrap(),
            Some(native_default_chart_border_stroke())
        );
        let generated = tsch::generated::ChartStyleArchive::decode(
            generated_chart_style_extension(&reset).unwrap().unwrap(),
        )
        .unwrap();
        assert_eq!(generated.tschchartinfodefaultborderstroke, None);
        assert_eq!(generated.tschchartinfodefaultshowborder, Some(true));
        assert_eq!(generated.tschchartinfodefaultintersetgap, Some(70.0));
        assert_unknown_fields_retained(&original, &reset);
    }

    fn test_stroke(pattern: StrokePattern, width: f32) -> ShapeStroke {
        ShapeStroke::new(
            RgbaColor::new(0.1, 0.3, 0.8, 1.0, RgbColorSpace::Srgb).unwrap(),
            StrokeWidth::new(width).unwrap(),
            pattern,
        )
    }

    fn style_with_unknown_fields(generated: tsch::generated::ChartStyleArchive) -> Vec<u8> {
        let mut extension = generated.encode_to_vec();
        append_varint_field(&mut extension, UNMAPPED_GENERATED_FIELD, UNMAPPED_VALUE).unwrap();
        let mut data = tsch::ChartStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();
        append_length_delimited_field(&mut data, GENERATED_CHART_STYLE_EXTENSION_FIELD, &extension)
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
                generated_chart_style_extension(patched).unwrap().unwrap(),
                UNMAPPED_GENERATED_FIELD,
            ),
            raw_field(
                generated_chart_style_extension(original).unwrap().unwrap(),
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
