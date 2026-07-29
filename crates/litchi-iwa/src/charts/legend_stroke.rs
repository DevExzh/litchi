//! Lossless native chart-legend stroke storage and mutation.
//!
//! Legend stroke is an independently inherited `TSD.StrokeArchive` in the
//! native legend-style extension. The exact override state is retained so
//! callers can distinguish inheritance from an explicit “None”.

use prost::Message;

use crate::charts::legend_style::{
    GENERATED_LEGEND_STYLE_EXTENSION_FIELD, generated_legend_style_extension, legend_style_slot,
};
use crate::protobuf::tsch;
use crate::shapes::{ShapeStroke, empty_stroke_archive, stroke_from_native, stroke_to_native};
use crate::wire::patch_length_delimited_field;
use crate::{Error, IWorkPackage, Result};

/// `tschlegendmodeldefaultstroke` in `TSCH.Generated.LegendStyleArchive`.
const LEGEND_STROKE_FIELD: u32 = 5;

/// Exact direct stroke state for a native chart legend.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub enum ChartLegendStroke {
    /// No direct override; iWork resolves the legend style's parent chain.
    #[default]
    Inherited,
    /// A direct empty stroke (the applications display “None”).
    NoStroke,
    /// A direct typed line stroke.
    Stroke(ShapeStroke),
}

/// Read the exact direct legend-stroke state of one native chart.
pub(crate) fn chart_legend_stroke(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<ChartLegendStroke> {
    legend_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?
    .read(package, read_legend_stroke)
}

/// Set or remove the direct legend-stroke override of one native chart.
pub(crate) fn set_chart_legend_stroke(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    stroke: ChartLegendStroke,
) -> Result<()> {
    let slot = legend_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    if slot.read(package, read_legend_stroke)? == stroke {
        return Ok(());
    }
    slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
    slot.update(package, |data| patch_legend_stroke(data, stroke))?;
    if slot.read(package, read_legend_stroke)? != stroke {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} legend stroke update failed validation"
        )));
    }
    Ok(())
}

fn read_legend_stroke(data: &[u8]) -> Result<ChartLegendStroke> {
    let Some(extension) = generated_legend_style_extension(data)? else {
        return Ok(ChartLegendStroke::Inherited);
    };
    let generated = tsch::generated::LegendStyleArchive::decode(extension)?;
    let Some(native) = generated.tschlegendmodeldefaultstroke.as_ref() else {
        return Ok(ChartLegendStroke::Inherited);
    };
    Ok(match stroke_from_native(native)? {
        Some(stroke) => ChartLegendStroke::Stroke(stroke),
        None => ChartLegendStroke::NoStroke,
    })
}

fn patch_legend_stroke(data: &[u8], stroke: ChartLegendStroke) -> Result<Vec<u8>> {
    let Some(extension) = generated_legend_style_extension(data)? else {
        let native = match stroke {
            ChartLegendStroke::Inherited => return Ok(data.to_vec()),
            ChartLegendStroke::NoStroke => empty_stroke_archive(),
            ChartLegendStroke::Stroke(stroke) => stroke_to_native(stroke),
        };
        let generated = tsch::generated::LegendStyleArchive {
            tschlegendmodeldefaultstroke: Some(native),
            ..Default::default()
        };
        let encoded = generated.encode_to_vec();
        let patched = patch_length_delimited_field(
            data,
            GENERATED_LEGEND_STYLE_EXTENSION_FIELD,
            false,
            Some(encoded.as_slice()),
        )?;
        validate_patched_legend_stroke(&patched, stroke)?;
        return Ok(patched);
    };

    let generated = tsch::generated::LegendStyleArchive::decode(extension)?;
    let stroke_present = generated.tschlegendmodeldefaultstroke.is_some();
    let native = match stroke {
        ChartLegendStroke::Inherited => None,
        ChartLegendStroke::NoStroke => Some(empty_stroke_archive().encode_to_vec()),
        ChartLegendStroke::Stroke(stroke) => Some(stroke_to_native(stroke).encode_to_vec()),
    };
    let extension = patch_length_delimited_field(
        extension,
        LEGEND_STROKE_FIELD,
        stroke_present,
        native.as_deref(),
    )?;
    let patched = patch_length_delimited_field(
        data,
        GENERATED_LEGEND_STYLE_EXTENSION_FIELD,
        true,
        Some(extension.as_slice()),
    )?;
    validate_patched_legend_stroke(&patched, stroke)?;
    Ok(patched)
}

fn validate_patched_legend_stroke(data: &[u8], expected: ChartLegendStroke) -> Result<()> {
    if read_legend_stroke(data)? != expected {
        return Err(Error::InvalidFormat(
            "Chart legend stroke wire patch failed validation".to_owned(),
        ));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::protobuf::tss;
    use crate::shapes::{
        RgbColorSpace, RgbaColor, ShapeFill, StrokePattern, StrokeWidth, fill_to_native,
    };
    use crate::wire::{append_length_delimited_field, append_varint_field, parse_wire_fields};

    const UNMAPPED_OUTER_FIELD: u32 = 4_096;
    const UNMAPPED_GENERATED_FIELD: u32 = 4_097;
    const UNMAPPED_VALUE: u64 = 42;

    #[test]
    fn legend_stroke_is_exact_and_preserves_fill_and_unknown_fields() {
        let fill =
            ShapeFill::Solid(RgbaColor::new(0.9, 0.95, 1.0, 1.0, RgbColorSpace::Srgb).unwrap());
        let mut generated = tsch::generated::LegendStyleArchive {
            tschlegendmodeldefaultfill: Some(fill_to_native(&fill)),
            tschlegendmodeldefaultopacity: Some(0.8),
            ..Default::default()
        }
        .encode_to_vec();
        append_varint_field(&mut generated, UNMAPPED_GENERATED_FIELD, UNMAPPED_VALUE).unwrap();
        let mut original = tsch::LegendStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();
        append_length_delimited_field(
            &mut original,
            GENERATED_LEGEND_STYLE_EXTENSION_FIELD,
            &generated,
        )
        .unwrap();
        append_varint_field(&mut original, UNMAPPED_OUTER_FIELD, UNMAPPED_VALUE).unwrap();

        assert_eq!(
            read_legend_stroke(&original).unwrap(),
            ChartLegendStroke::Inherited
        );
        let stroke = ChartLegendStroke::Stroke(ShapeStroke::new(
            RgbaColor::new(0.1, 0.3, 0.8, 1.0, RgbColorSpace::Srgb).unwrap(),
            StrokeWidth::new(2.5).unwrap(),
            StrokePattern::MediumDash,
        ));
        let stroked = patch_legend_stroke(&original, stroke).unwrap();
        assert_eq!(read_legend_stroke(&stroked).unwrap(), stroke);
        let generated_after = generated_legend_style_extension(&stroked).unwrap().unwrap();
        let decoded = tsch::generated::LegendStyleArchive::decode(generated_after).unwrap();
        assert_eq!(
            decoded.tschlegendmodeldefaultfill,
            Some(fill_to_native(&fill))
        );
        assert_eq!(decoded.tschlegendmodeldefaultopacity, Some(0.8));
        assert!(
            parse_wire_fields(&stroked)
                .unwrap()
                .iter()
                .any(|field| field.number == UNMAPPED_OUTER_FIELD && field.wire_type == 0)
        );
        assert!(
            parse_wire_fields(generated_after)
                .unwrap()
                .iter()
                .any(|field| field.number == UNMAPPED_GENERATED_FIELD && field.wire_type == 0)
        );

        let empty = patch_legend_stroke(&stroked, ChartLegendStroke::NoStroke).unwrap();
        assert_eq!(
            read_legend_stroke(&empty).unwrap(),
            ChartLegendStroke::NoStroke
        );
        let inherited = patch_legend_stroke(&empty, ChartLegendStroke::Inherited).unwrap();
        assert_eq!(inherited, original);
    }
}
