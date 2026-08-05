//! Native chart-wide Radar series-style CRUD.
//!
//! The Radar `Style` control applies one coordinated representation to every
//! series: translucent fills, colored strokes, or strokes with derived fills.
//! iWork stores that choice across three inherited series-style properties, so
//! exposing the control as one enum prevents callers from creating partial or
//! contradictory combinations.

use prost::Message;

use crate::charts::Kind;
use crate::charts::series_style::{
    ChartSeriesStyleSlot, GENERATED_CHART_SERIES_STYLE_EXTENSION_FIELD,
    effective_chart_series_style_slots, generated_chart_series_style_extension,
};
use crate::protobuf::tsch;
use crate::shapes::{
    RgbaColor, ShapeFill, ShapeStroke, StrokePattern, StrokeWidth, fill_from_native,
    fill_to_native, stroke_from_native, stroke_to_native,
};
use crate::wire::{
    parse_wire_fields, patch_fixed32_field, patch_length_delimited_field, patch_varint_field,
};
use crate::{Error, IWorkPackage, Result};

const RADAR_FILL_FIELD: u32 = 165;
const RADAR_STROKE_FIELD: u32 = 172;
const RADAR_FILL_USES_STROKE_FIELD: u32 = 188;
const RADAR_FILL_STROKE_ALPHA_MULTIPLIER_FIELD: u32 = 189;
const NATIVE_FILL_ONLY_ALPHA: f32 = 0.5;
const NATIVE_FILL_AND_STROKE_ALPHA_MULTIPLIER: f32 = 0.15;
const NATIVE_RADAR_STROKE_WIDTH_POINTS: f32 = 4.0;

/// Appearance selected by the chart-wide Radar `Style` control.
#[derive(Debug, Default, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ChartRadarSeriesStyle {
    /// Translucent series-color fills with no outlines.
    Fill,
    /// Series-color outlines with no area fill.
    Stroke,
    /// Series-color outlines with a derived translucent area fill.
    #[default]
    FillAndStroke,
}

/// Read the uniform native Radar style applied to every data series.
pub(crate) fn chart_radar_series_style(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    kind: Kind,
    series_count: usize,
) -> Result<ChartRadarSeriesStyle> {
    require_radar_chart(kind)?;
    let slots = effective_chart_series_style_slots(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        series_count,
    )?;
    let mut styles = slots
        .iter()
        .map(|slot| effective_radar_style(slot, package));
    let first = styles.next().transpose()?.ok_or_else(|| {
        Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} has no Radar series"
        ))
    })?;
    for style in styles {
        let style = style?;
        if style != first {
            return Err(Error::InvalidFormat(format!(
                "{drawable_label} chart {drawable_object_id} has mixed Radar series styles"
            )));
        }
    }
    Ok(first)
}

/// Apply one native Radar style to every data series.
pub(crate) fn set_chart_radar_series_style(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    kind: Kind,
    series_count: usize,
    style: ChartRadarSeriesStyle,
) -> Result<()> {
    require_radar_chart(kind)?;
    let slots = effective_chart_series_style_slots(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        series_count,
    )?;
    if slots.is_empty() {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} has no Radar series"
        )));
    }
    let colors = slots
        .iter()
        .map(|slot| effective_radar_color(slot, package))
        .collect::<Result<Vec<_>>>()?;
    for (slot, color) in slots.iter().zip(colors) {
        if effective_radar_style(slot, package).is_ok_and(|current| current == style) {
            continue;
        }
        slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
        slot.update(package, |data| patch_local_radar_style(data, color, style))?;
    }
    if chart_radar_series_style(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        kind,
        series_count,
    )? != style
    {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} Radar style update failed validation"
        )));
    }
    Ok(())
}

fn require_radar_chart(kind: Kind) -> Result<()> {
    if !kind.supports_radar_series_style() {
        return Err(Error::InvalidFormat(format!(
            "chart kind {kind:?} does not expose Radar series styles"
        )));
    }
    Ok(())
}

fn effective_radar_style(
    slot: &ChartSeriesStyleSlot,
    package: &IWorkPackage,
) -> Result<ChartRadarSeriesStyle> {
    let fill = slot
        .read_inherited(package, read_local_radar_fill)?
        .unwrap_or_default();
    let stroke = slot
        .read_inherited(package, read_local_radar_stroke)?
        .flatten();
    let uses_stroke = slot
        .read_inherited(package, read_local_fill_uses_stroke)?
        .unwrap_or(false);
    if uses_stroke {
        validate_effective_alpha_multiplier(slot, package)?;
    }
    classify_radar_style(&fill, stroke.as_ref(), uses_stroke)
}

fn classify_radar_style(
    fill: &ShapeFill,
    stroke: Option<&ShapeStroke>,
    uses_stroke: bool,
) -> Result<ChartRadarSeriesStyle> {
    let has_fill = !matches!(fill, ShapeFill::None);
    let has_stroke = stroke.is_some();
    match (uses_stroke, has_fill, has_stroke) {
        (false, true, false) => Ok(ChartRadarSeriesStyle::Fill),
        (false, false, true) => Ok(ChartRadarSeriesStyle::Stroke),
        (true, _, true) => Ok(ChartRadarSeriesStyle::FillAndStroke),
        state => Err(Error::InvalidFormat(format!(
            "native Radar series style has inconsistent fill/stroke state {state:?}"
        ))),
    }
}

fn validate_effective_alpha_multiplier(
    slot: &ChartSeriesStyleSlot,
    package: &IWorkPackage,
) -> Result<()> {
    let multiplier = slot
        .read_inherited(package, read_local_alpha_multiplier)?
        .unwrap_or(NATIVE_FILL_AND_STROKE_ALPHA_MULTIPLIER);
    if !multiplier.is_finite() || !(0.0..=1.0).contains(&multiplier) {
        return Err(Error::InvalidFormat(format!(
            "native Radar fill/stroke alpha multiplier {multiplier} must be finite and between zero and one"
        )));
    }
    Ok(())
}

fn effective_radar_color(slot: &ChartSeriesStyleSlot, package: &IWorkPackage) -> Result<RgbaColor> {
    if let Some(stroke) = slot
        .read_inherited(package, read_local_radar_stroke)?
        .flatten()
    {
        return opaque(stroke.color);
    }
    let fill = slot
        .read_inherited(package, read_local_radar_fill)?
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "chart series style {} has no inherited Radar color",
                slot.object_id()
            ))
        })?;
    let ShapeFill::Solid(color) = fill else {
        return Err(Error::InvalidFormat(format!(
            "chart series style {} needs a solid Radar palette color",
            slot.object_id()
        )));
    };
    opaque(color)
}

fn opaque(color: RgbaColor) -> Result<RgbaColor> {
    Ok(RgbaColor::new(
        color.red(),
        color.green(),
        color.blue(),
        1.0,
        color.color_space(),
    )?)
}

fn with_alpha(color: RgbaColor, alpha: f32) -> Result<RgbaColor> {
    Ok(RgbaColor::new(
        color.red(),
        color.green(),
        color.blue(),
        alpha,
        color.color_space(),
    )?)
}

fn read_local_radar_fill(data: &[u8]) -> Result<Option<ShapeFill>> {
    let Some(extension) = generated_chart_series_style_extension(data)? else {
        return Ok(None);
    };
    require_singular_field(extension, RADAR_FILL_FIELD, 2)?;
    let generated = tsch::generated::ChartSeriesStyleArchive::decode(extension)?;
    generated
        .tschchartseriesradarareafill
        .as_ref()
        .map(fill_from_native)
        .transpose()
}

fn read_local_radar_stroke(data: &[u8]) -> Result<Option<Option<ShapeStroke>>> {
    let Some(extension) = generated_chart_series_style_extension(data)? else {
        return Ok(None);
    };
    require_singular_field(extension, RADAR_STROKE_FIELD, 2)?;
    let generated = tsch::generated::ChartSeriesStyleArchive::decode(extension)?;
    generated
        .tschchartseriesradarareastroke
        .as_ref()
        .map(stroke_from_native)
        .transpose()
}

fn read_local_fill_uses_stroke(data: &[u8]) -> Result<Option<bool>> {
    let Some(extension) = generated_chart_series_style_extension(data)? else {
        return Ok(None);
    };
    let Some(field) = singular_field(extension, RADAR_FILL_USES_STROKE_FIELD, 0)? else {
        return Ok(None);
    };
    let (value, length) = litchi_iwa_common::varint::decode_varint_from_bytes(
        &extension[field.key_end()..field.end()],
    )
    .map_err(|error| Error::InvalidFormat(format!("invalid Radar style boolean: {error}")))?;
    if field.key_end() + length != field.end() || value > 1 {
        return Err(Error::InvalidFormat(format!(
            "native Radar style boolean must be zero or one, not {value}"
        )));
    }
    Ok(Some(value == 1))
}

fn read_local_alpha_multiplier(data: &[u8]) -> Result<Option<f32>> {
    let Some(extension) = generated_chart_series_style_extension(data)? else {
        return Ok(None);
    };
    let Some(field) = singular_field(extension, RADAR_FILL_STROKE_ALPHA_MULTIPLIER_FIELD, 5)?
    else {
        return Ok(None);
    };
    let bytes: [u8; 4] = extension[field.payload_start()..field.end()]
        .try_into()
        .map_err(|_| Error::InvalidFormat("truncated Radar alpha multiplier".to_owned()))?;
    Ok(Some(f32::from_le_bytes(bytes)))
}

fn require_singular_field(data: &[u8], number: u32, wire_type: u8) -> Result<()> {
    let _ = singular_field(data, number, wire_type)?;
    Ok(())
}

fn singular_field(
    data: &[u8],
    number: u32,
    wire_type: u8,
) -> Result<Option<crate::wire::WireField>> {
    let fields = parse_wire_fields(data)?;
    let matches = fields
        .into_iter()
        .filter(|field| field.number() == number)
        .collect::<Vec<_>>();
    let [field] = matches.as_slice() else {
        if matches.is_empty() {
            return Ok(None);
        }
        return Err(Error::InvalidFormat(format!(
            "Radar series-style field {number} occurs more than once"
        )));
    };
    if field.wire_type() != wire_type {
        return Err(Error::InvalidFormat(format!(
            "Radar series-style field {number} has wire type {}, not {wire_type}",
            field.wire_type()
        )));
    }
    Ok(Some(*field))
}

fn patch_local_radar_style(
    data: &[u8],
    color: RgbaColor,
    style: ChartRadarSeriesStyle,
) -> Result<Vec<u8>> {
    let existing_extension = generated_chart_series_style_extension(data)?;
    let mut extension = existing_extension.unwrap_or_default().to_vec();
    tsch::generated::ChartSeriesStyleArchive::decode(extension.as_slice())?;

    let fill = match style {
        ChartRadarSeriesStyle::Fill => ShapeFill::Solid(with_alpha(color, NATIVE_FILL_ONLY_ALPHA)?),
        ChartRadarSeriesStyle::Stroke | ChartRadarSeriesStyle::FillAndStroke => ShapeFill::None,
    };
    let fill = fill_to_native(&fill).encode_to_vec();
    let fill_present = singular_field(&extension, RADAR_FILL_FIELD, 2)?.is_some();
    extension =
        patch_length_delimited_field(&extension, RADAR_FILL_FIELD, fill_present, Some(&fill))?;

    let stroke_present = singular_field(&extension, RADAR_STROKE_FIELD, 2)?.is_some();
    let stroke = if style == ChartRadarSeriesStyle::Fill {
        None
    } else {
        let stroke = ShapeStroke::new(
            color,
            StrokeWidth::new(NATIVE_RADAR_STROKE_WIDTH_POINTS)?,
            StrokePattern::Solid,
        );
        Some(stroke_to_native(stroke).encode_to_vec())
    };
    extension = patch_length_delimited_field(
        &extension,
        RADAR_STROKE_FIELD,
        stroke_present,
        stroke.as_deref(),
    )?;

    let uses_stroke_present =
        singular_field(&extension, RADAR_FILL_USES_STROKE_FIELD, 0)?.is_some();
    extension = patch_varint_field(
        &extension,
        RADAR_FILL_USES_STROKE_FIELD,
        uses_stroke_present,
        Some(u64::from(style == ChartRadarSeriesStyle::FillAndStroke)),
    )?;

    let multiplier_present =
        singular_field(&extension, RADAR_FILL_STROKE_ALPHA_MULTIPLIER_FIELD, 5)?.is_some();
    let multiplier = (style == ChartRadarSeriesStyle::FillAndStroke)
        .then_some(NATIVE_FILL_AND_STROKE_ALPHA_MULTIPLIER.to_bits());
    extension = patch_fixed32_field(
        &extension,
        RADAR_FILL_STROKE_ALPHA_MULTIPLIER_FIELD,
        multiplier_present,
        multiplier,
    )?;

    let patched = patch_length_delimited_field(
        data,
        GENERATED_CHART_SERIES_STYLE_EXTENSION_FIELD,
        existing_extension.is_some(),
        Some(&extension),
    )?;
    validate_local_radar_style(&patched, style)?;
    Ok(patched)
}

fn validate_local_radar_style(data: &[u8], expected: ChartRadarSeriesStyle) -> Result<()> {
    let fill = read_local_radar_fill(data)?
        .ok_or_else(|| Error::InvalidFormat("Radar style patch did not write a fill".to_owned()))?;
    let stroke = read_local_radar_stroke(data)?.flatten();
    let uses_stroke = read_local_fill_uses_stroke(data)?.ok_or_else(|| {
        Error::InvalidFormat("Radar style patch did not write its mode flag".to_owned())
    })?;
    if classify_radar_style(&fill, stroke.as_ref(), uses_stroke)? != expected {
        return Err(Error::InvalidFormat(
            "Radar series-style wire patch failed validation".to_owned(),
        ));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::protobuf::{tsd, tsp, tss};
    use crate::shapes::RgbColorSpace;
    use crate::wire::{append_length_delimited_field, append_varint_field};

    const UNKNOWN_OUTER_FIELD: u32 = 4_096;
    const UNKNOWN_EXTENSION_FIELD: u32 = 4_097;
    const UNKNOWN_VALUE: u64 = 42;

    #[test]
    fn radar_styles_are_closed_and_fill_and_stroke_is_default() {
        assert_eq!(
            ChartRadarSeriesStyle::default(),
            ChartRadarSeriesStyle::FillAndStroke
        );
    }

    #[test]
    fn all_native_modes_patch_canonically_and_preserve_unknown_fields() {
        let mut extension = tsch::generated::ChartSeriesStyleArchive {
            tschchartseriesradarareafill: Some(solid_fill(color(0.2, 0.5, 0.8, 1.0))),
            tschchartseriesradarareafilluseseriesstroke: Some(true),
            tschchartseriesradarareafilluseseriesstrokealphamultiplier: Some(0.25),
            tschchartseriesdefaultopacity: Some(0.8),
            ..Default::default()
        }
        .encode_to_vec();
        append_varint_field(&mut extension, UNKNOWN_EXTENSION_FIELD, UNKNOWN_VALUE).unwrap();
        let mut original = tsch::ChartSeriesStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();
        append_length_delimited_field(
            &mut original,
            GENERATED_CHART_SERIES_STYLE_EXTENSION_FIELD,
            &extension,
        )
        .unwrap();
        append_varint_field(&mut original, UNKNOWN_OUTER_FIELD, UNKNOWN_VALUE).unwrap();

        for style in [
            ChartRadarSeriesStyle::Fill,
            ChartRadarSeriesStyle::Stroke,
            ChartRadarSeriesStyle::FillAndStroke,
        ] {
            let patched =
                patch_local_radar_style(&original, color(0.2, 0.5, 0.8, 1.0), style).unwrap();
            validate_local_radar_style(&patched, style).unwrap();
            assert!(has_field(&patched, UNKNOWN_OUTER_FIELD));
            let extension = generated_chart_series_style_extension(&patched)
                .unwrap()
                .unwrap();
            assert!(has_field(extension, UNKNOWN_EXTENSION_FIELD));
            let decoded = tsch::generated::ChartSeriesStyleArchive::decode(extension).unwrap();
            assert_eq!(decoded.tschchartseriesdefaultopacity, Some(0.8));
        }
    }

    #[test]
    fn inconsistent_and_malformed_native_modes_are_rejected() {
        assert!(
            classify_radar_style(&ShapeFill::None, None, false).is_err(),
            "neither fill nor stroke is not a native mode"
        );
        assert!(
            classify_radar_style(&ShapeFill::Solid(color(0.2, 0.5, 0.8, 1.0)), None, true).is_err(),
            "derived fill requires a stroke"
        );

        let generated = tsch::generated::ChartSeriesStyleArchive {
            tschchartseriesradarareafill: Some(solid_fill(color(0.2, 0.5, 0.8, 1.0))),
            ..Default::default()
        };
        let mut extension = generated.encode_to_vec();
        append_varint_field(&mut extension, RADAR_FILL_USES_STROKE_FIELD, 2).unwrap();
        let mut outer = tsch::ChartSeriesStyleArchive::default().encode_to_vec();
        append_length_delimited_field(
            &mut outer,
            GENERATED_CHART_SERIES_STYLE_EXTENSION_FIELD,
            &extension,
        )
        .unwrap();
        assert!(read_local_fill_uses_stroke(&outer).is_err());
    }

    fn solid_fill(color: RgbaColor) -> tsd::FillArchive {
        tsd::FillArchive {
            color: Some(tsp::Color {
                model: tsp::color::ColorModel::Rgb as i32,
                r: Some(color.red()),
                g: Some(color.green()),
                b: Some(color.blue()),
                a: Some(color.alpha()),
                rgbspace: Some(tsp::color::RgbColorSpace::Srgb as i32),
                ..Default::default()
            }),
            ..Default::default()
        }
    }

    fn color(red: f32, green: f32, blue: f32, alpha: f32) -> RgbaColor {
        RgbaColor::new(red, green, blue, alpha, RgbColorSpace::Srgb).unwrap()
    }

    fn has_field(data: &[u8], number: u32) -> bool {
        parse_wire_fields(data)
            .unwrap()
            .iter()
            .any(|field| field.number() == number)
    }
}
