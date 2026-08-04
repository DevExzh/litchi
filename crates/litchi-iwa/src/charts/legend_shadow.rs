//! Lossless native chart-legend shadow storage and mutation.
//!
//! Legend shadow is an independently inherited `TSD.ShadowArchive` in the
//! native legend-style extension. The legend inspector supports drop shadows
//! only, so contact and curved shadows are rejected rather than normalized.

use prost::Message;

use crate::charts::legend_style::{
    GENERATED_LEGEND_STYLE_EXTENSION_FIELD, generated_legend_style_extension, legend_style_slot,
};
use crate::protobuf::tsch;
use crate::shapes::{ShapeDropShadow, ShapeShadow, shadow_from_native, shadow_to_native};
use crate::wire::patch_length_delimited_field;
use crate::{Error, IWorkPackage, Result};

/// `tschlegendmodeldefaultshadow` in `TSCH.Generated.LegendStyleArchive`.
const LEGEND_SHADOW_FIELD: u32 = 4;

/// Exact direct shadow state for a native chart legend.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub enum ChartLegendShadow {
    /// No direct override; iWork resolves the legend style's parent chain.
    #[default]
    Inherited,
    /// A direct disabled shadow (the inspector checkbox is off).
    NoShadow,
    /// A direct typed drop shadow.
    Shadow(ShapeDropShadow),
}

/// Read the exact direct legend-shadow state of one native chart.
pub(crate) fn chart_legend_shadow(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<ChartLegendShadow> {
    legend_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?
    .read(package, read_legend_shadow)
}

/// Set or remove the direct legend-shadow override of one native chart.
pub(crate) fn set_chart_legend_shadow(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    shadow: ChartLegendShadow,
) -> Result<()> {
    let mut slot = legend_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    if slot.read(package, read_legend_shadow)? == shadow {
        return Ok(());
    }
    slot.ensure_exclusive(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    slot.update(package, |data| patch_legend_shadow(data, shadow))?;
    slot.collapse_if_equivalent(package, chart_archive_name, drawable_object_id)?;
    if chart_legend_shadow(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )? != shadow
    {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} legend shadow update failed validation"
        )));
    }
    Ok(())
}

fn read_legend_shadow(data: &[u8]) -> Result<ChartLegendShadow> {
    let Some(extension) = generated_legend_style_extension(data)? else {
        return Ok(ChartLegendShadow::Inherited);
    };
    let generated = tsch::generated::LegendStyleArchive::decode(extension)?;
    let Some(native) = generated.tschlegendmodeldefaultshadow.as_ref() else {
        return Ok(ChartLegendShadow::Inherited);
    };
    match shadow_from_native(native)? {
        ShapeShadow::Disabled => Ok(ChartLegendShadow::NoShadow),
        ShapeShadow::Drop(shadow) => Ok(ChartLegendShadow::Shadow(shadow)),
        ShapeShadow::Contact(_) | ShapeShadow::Curved(_) => Err(Error::InvalidFormat(
            "native chart legend uses a non-drop shadow".to_owned(),
        )),
    }
}

fn patch_legend_shadow(data: &[u8], shadow: ChartLegendShadow) -> Result<Vec<u8>> {
    let Some(extension) = generated_legend_style_extension(data)? else {
        let native = match shadow {
            ChartLegendShadow::Inherited => return Ok(data.to_vec()),
            ChartLegendShadow::NoShadow => shadow_to_native(ShapeShadow::Disabled),
            ChartLegendShadow::Shadow(shadow) => shadow_to_native(ShapeShadow::Drop(shadow)),
        };
        let generated = tsch::generated::LegendStyleArchive {
            tschlegendmodeldefaultshadow: Some(native),
            ..Default::default()
        };
        let encoded = generated.encode_to_vec();
        let patched = patch_length_delimited_field(
            data,
            GENERATED_LEGEND_STYLE_EXTENSION_FIELD,
            false,
            Some(encoded.as_slice()),
        )?;
        validate_patched_legend_shadow(&patched, shadow)?;
        return Ok(patched);
    };

    let generated = tsch::generated::LegendStyleArchive::decode(extension)?;
    let shadow_present = generated.tschlegendmodeldefaultshadow.is_some();
    let native = match shadow {
        ChartLegendShadow::Inherited => None,
        ChartLegendShadow::NoShadow => {
            Some(shadow_to_native(ShapeShadow::Disabled).encode_to_vec())
        },
        ChartLegendShadow::Shadow(shadow) => {
            Some(shadow_to_native(ShapeShadow::Drop(shadow)).encode_to_vec())
        },
    };
    let extension = patch_length_delimited_field(
        extension,
        LEGEND_SHADOW_FIELD,
        shadow_present,
        native.as_deref(),
    )?;
    let patched = patch_length_delimited_field(
        data,
        GENERATED_LEGEND_STYLE_EXTENSION_FIELD,
        true,
        Some(extension.as_slice()),
    )?;
    validate_patched_legend_shadow(&patched, shadow)?;
    Ok(patched)
}

fn validate_patched_legend_shadow(data: &[u8], expected: ChartLegendShadow) -> Result<()> {
    if read_legend_shadow(data)? != expected {
        return Err(Error::InvalidFormat(
            "Chart legend shadow wire patch failed validation".to_owned(),
        ));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::protobuf::tss;
    use crate::shapes::{
        RgbColorSpace, RgbaColor, ShapeFill, ShapeShadowAngle, ShapeShadowAppearance,
        ShapeShadowBlurRadius, ShapeShadowOffset, ShapeShadowOpacity, ShapeStroke, StrokePattern,
        StrokeWidth, fill_to_native, stroke_to_native,
    };
    use crate::wire::{append_length_delimited_field, append_varint_field, parse_wire_fields};

    const UNMAPPED_OUTER_FIELD: u32 = 4_096;
    const UNMAPPED_GENERATED_FIELD: u32 = 4_097;
    const UNMAPPED_VALUE: u64 = 42;

    #[test]
    fn legend_shadow_is_exact_and_preserves_fill_stroke_and_unknown_fields() {
        let fill =
            ShapeFill::Solid(RgbaColor::new(0.9, 0.95, 1.0, 1.0, RgbColorSpace::Srgb).unwrap());
        let stroke = ShapeStroke::new(
            RgbaColor::new(0.1, 0.3, 0.8, 1.0, RgbColorSpace::Srgb).unwrap(),
            StrokeWidth::new(2.5).unwrap(),
            StrokePattern::MediumDash,
        );
        let mut generated = tsch::generated::LegendStyleArchive {
            tschlegendmodeldefaultfill: Some(fill_to_native(&fill)),
            tschlegendmodeldefaultopacity: Some(0.8),
            tschlegendmodeldefaultstroke: Some(stroke_to_native(stroke)),
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
            read_legend_shadow(&original).unwrap(),
            ChartLegendShadow::Inherited
        );
        let shadow = ChartLegendShadow::Shadow(ShapeDropShadow::new(
            ShapeShadowAppearance::new(
                RgbaColor::black(),
                ShapeShadowBlurRadius::from_points(12).unwrap(),
                ShapeShadowOffset::from_points(8.0).unwrap(),
                ShapeShadowOpacity::new(0.6).unwrap(),
            ),
            ShapeShadowAngle::from_degrees(30.0).unwrap(),
        ));
        let shadowed = patch_legend_shadow(&original, shadow).unwrap();
        assert_eq!(read_legend_shadow(&shadowed).unwrap(), shadow);
        let generated_after = generated_legend_style_extension(&shadowed)
            .unwrap()
            .unwrap();
        let decoded = tsch::generated::LegendStyleArchive::decode(generated_after).unwrap();
        assert_eq!(
            decoded.tschlegendmodeldefaultfill,
            Some(fill_to_native(&fill))
        );
        assert_eq!(
            decoded.tschlegendmodeldefaultstroke,
            Some(stroke_to_native(stroke))
        );
        assert_eq!(decoded.tschlegendmodeldefaultopacity, Some(0.8));
        assert!(
            parse_wire_fields(&shadowed)
                .unwrap()
                .iter()
                .any(|field| field.number() == UNMAPPED_OUTER_FIELD && field.wire_type() == 0)
        );
        assert!(
            parse_wire_fields(generated_after)
                .unwrap()
                .iter()
                .any(|field| field.number() == UNMAPPED_GENERATED_FIELD && field.wire_type() == 0)
        );

        let disabled = patch_legend_shadow(&shadowed, ChartLegendShadow::NoShadow).unwrap();
        assert_eq!(
            read_legend_shadow(&disabled).unwrap(),
            ChartLegendShadow::NoShadow
        );
        let inherited = patch_legend_shadow(&disabled, ChartLegendShadow::Inherited).unwrap();
        assert_eq!(inherited, original);
    }
}
