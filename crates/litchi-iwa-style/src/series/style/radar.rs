//! The chart-wide Radar series style (the host's
//! `charts::radar_series_style`, record 0419).
//!
//! The Radar inspector's "Style" control applies one coordinated
//! representation to every series. iWork stores it in three inherited
//! series-style properties of the generated extension: the area fill (field
//! 165), the area stroke (172) and whether the fill derives from the stroke
//! (188, with its alpha multiplier in 189). A read follows each visible
//! series' effective style and its parent chain and requires one uniform,
//! consistent value, as the host does; a write rewrites, in series order,
//! every effective style whose value differs, after requiring it to be
//! exclusive to the chart.
//!
//! Two layers: [`effective`] and [`rewrites`] work on payloads for an owner
//! that publishes its own in-place rewrites; the bounded workspace layer
//! serves `chart_style::{read_chart, write_chart}`.

use litchi_iwa_common::color::Rgba;
use litchi_iwa_common::shape::fill::Fill;
use litchi_iwa_common::shape::stroke::{Pattern, Stroke, Width};
use litchi_iwa_common::wire;
use litchi_iwa_protos::chart_style_codec::{self as codec, Limits as CodecLimits, Properties};
use litchi_iwa_protos::shape_style_codec::native as stored;

use super::{decode, fill_at, inherited, invalid, workspace_payload};
use crate::chart_style::chart::RadarSeriesStyle;
use crate::drawing;
use crate::graph::{Error, Workspace};
use crate::series::{
    Refusal, SERIES_STYLE_MESSAGE_TYPE, SymbolBudget, SymbolError, SymbolResult, bounded_payload,
    bounded_style_owner_count, bounded_style_slots, charged_patch,
};

/// `tschchartseriesradarareafill`.
const FILL_FIELD: u32 = 165;
/// `tschchartseriesradarareastroke`.
const STROKE_FIELD: u32 = 172;
/// `tschchartseriesradarareafilluseseriesstroke`.
const FILL_USES_STROKE_FIELD: u32 = 188;
/// `tschchartseriesradarareafilluseseriesstrokealphamultiplier`.
const ALPHA_MULTIPLIER_FIELD: u32 = 189;
/// The alpha of a fill-only series' fill.
const FILL_ONLY_ALPHA: f32 = 0.5;
/// The alpha multiplier of a derived fill.
const FILL_AND_STROKE_ALPHA_MULTIPLIER: f32 = 0.15;
/// The width of a Radar series outline, in points.
const STROKE_WIDTH_POINTS: f32 = 4.0;

fn codec_error(error: codec::Error) -> SymbolError {
    SymbolError::Graph(Error::Codec(error))
}

fn local_fill(properties: &Properties<'_>) -> SymbolResult<Option<Fill>> {
    fill_at(properties, FILL_FIELD)
}

/// The stored area stroke (`Some(None)`: a hidden stroke).
fn local_stroke(properties: &Properties<'_>) -> SymbolResult<Option<Option<Stroke>>> {
    let Some(bytes) = properties.message(STROKE_FIELD).map_err(codec_error)? else {
        return Ok(None);
    };
    let decoded = stored::stroke(bytes, CodecLimits::DEFAULT).map_err(invalid)?;
    Ok(Some(drawing::stroke(&decoded).map_err(invalid)?))
}

/// The derived-fill switch, a varint of zero or one (the host's strict
/// boolean).
fn local_uses_stroke(properties: &Properties<'_>) -> SymbolResult<Option<bool>> {
    match properties
        .varint(FILL_USES_STROKE_FIELD)
        .map_err(codec_error)?
    {
        None => Ok(None),
        Some(value @ (0 | 1)) => Ok(Some(value == 1)),
        Some(_) => Err(SymbolError::Invalid(
            "a Radar style switch is neither zero nor one",
        )),
    }
}

fn local_alpha_multiplier(properties: &Properties<'_>) -> SymbolResult<Option<f32>> {
    properties
        .float(ALPHA_MULTIPLIER_FIELD)
        .map_err(codec_error)
}

/// The host's `classify_radar_style`.
fn classify(
    fill: &Fill,
    stroke: Option<&Stroke>,
    uses_stroke: bool,
) -> SymbolResult<RadarSeriesStyle> {
    let has_fill = !matches!(fill, Fill::None);
    match (uses_stroke, has_fill, stroke.is_some()) {
        (false, true, false) => Ok(RadarSeriesStyle::Fill),
        (false, false, true) => Ok(RadarSeriesStyle::Stroke),
        (true, _, true) => Ok(RadarSeriesStyle::FillAndStroke),
        _ => Err(SymbolError::Invalid(
            "a Radar series style has an inconsistent fill and stroke",
        )),
    }
}

/// The effective Radar style of one series style (the host's
/// `effective_radar_style`).
fn style_of<'a>(
    style: u64,
    payload: &mut impl FnMut(u64) -> SymbolResult<&'a [u8]>,
) -> SymbolResult<RadarSeriesStyle> {
    let fill = inherited(style, payload, local_fill)?.unwrap_or_default();
    let stroke = inherited(style, payload, local_stroke)?.flatten();
    let uses_stroke = inherited(style, payload, local_uses_stroke)?.unwrap_or(false);
    if uses_stroke {
        let multiplier = inherited(style, payload, local_alpha_multiplier)?
            .unwrap_or(FILL_AND_STROKE_ALPHA_MULTIPLIER);
        if !multiplier.is_finite() || !(0.0..=1.0).contains(&multiplier) {
            return Err(SymbolError::Invalid(
                "a Radar fill alpha multiplier is outside zero to one",
            ));
        }
    }
    classify(&fill, stroke.as_ref(), uses_stroke)
}

fn with_alpha(color: Rgba, alpha: f32) -> SymbolResult<Rgba> {
    Rgba::new(
        color.red(),
        color.green(),
        color.blue(),
        alpha,
        color.color_space(),
    )
    .map_err(invalid)
}

/// The opaque series color a Radar style is built from: the effective
/// stroke's color, else the effective solid fill's (the host's
/// `effective_radar_color`).
fn color_of<'a>(
    style: u64,
    payload: &mut impl FnMut(u64) -> SymbolResult<&'a [u8]>,
) -> SymbolResult<Rgba> {
    if let Some(stroke) = inherited(style, payload, local_stroke)?.flatten() {
        return with_alpha(stroke.color, 1.0);
    }
    match inherited(style, payload, local_fill)? {
        Some(Fill::Solid(color)) => with_alpha(color, 1.0),
        Some(_) => Err(SymbolError::Invalid(
            "a Radar series needs a solid palette color",
        )),
        None => Err(SymbolError::Invalid("a Radar series has no color")),
    }
}

/// The series style payload storing `style` in series `color`, every other
/// byte kept (the host's `patch_local_radar_style`).
fn rewrite(payload: &[u8], color: Rgba, style: RadarSeriesStyle) -> SymbolResult<Vec<u8>> {
    let wire_error = |error| SymbolError::Graph(Error::Wire(error));
    let properties = decode(payload)?;
    let existing = properties.extension();
    let mut extension = existing.unwrap_or_default().to_vec();
    let fill = match style {
        RadarSeriesStyle::Fill => Fill::Solid(with_alpha(color, FILL_ONLY_ALPHA)?),
        RadarSeriesStyle::Stroke | RadarSeriesStyle::FillAndStroke => Fill::None,
    };
    let fill = drawing::encode_fill(&fill);
    let present = local_fill(&properties)?.is_some();
    extension = wire::patch_length_delimited_field(&extension, FILL_FIELD, present, Some(&fill))
        .map_err(wire_error)?;
    let present = local_stroke(&properties)?.is_some();
    let width = Width::new(STROKE_WIDTH_POINTS).map_err(invalid)?;
    let stroke = (style != RadarSeriesStyle::Fill)
        .then(|| drawing::encode_stroke(Some(Stroke::new(color, width, Pattern::Solid))));
    extension =
        wire::patch_length_delimited_field(&extension, STROKE_FIELD, present, stroke.as_deref())
            .map_err(wire_error)?;
    let present = local_uses_stroke(&properties)?.is_some();
    extension = wire::patch_varint_field(
        &extension,
        FILL_USES_STROKE_FIELD,
        present,
        Some(u64::from(style == RadarSeriesStyle::FillAndStroke)),
    )
    .map_err(wire_error)?;
    let present = local_alpha_multiplier(&properties)?.is_some();
    extension = wire::patch_fixed32_field(
        &extension,
        ALPHA_MULTIPLIER_FIELD,
        present,
        (style == RadarSeriesStyle::FillAndStroke)
            .then_some(FILL_AND_STROKE_ALPHA_MULTIPLIER.to_bits()),
    )
    .map_err(wire_error)?;
    let patched = wire::patch_length_delimited_field(
        payload,
        codec::EXTENSION_FIELD,
        existing.is_some(),
        Some(&extension),
    )
    .map_err(wire_error)?;
    // The host's `validate_local_radar_style`.
    let written = decode(&patched)?;
    let fill =
        local_fill(&written)?.ok_or(SymbolError::Invalid("a Radar style write stored no fill"))?;
    let stroke = local_stroke(&written)?.flatten();
    let uses_stroke = local_uses_stroke(&written)?
        .ok_or(SymbolError::Invalid("a Radar style write stored no switch"))?;
    if classify(&fill, stroke.as_ref(), uses_stroke)? != style {
        return Err(SymbolError::Invalid(
            "a Radar style write does not read back",
        ));
    }
    Ok(patched)
}

/// The uniform Radar style of the series styles `styles` (one per visible
/// series, in series order). `payload` returns the
/// `TSCH.ChartSeriesStyleArchive` payload of a style object.
///
/// # Errors
///
/// No series, mixed or inconsistent styles, or an invalid style or parent
/// chain.
pub fn effective<'a>(
    styles: &[u64],
    mut payload: impl FnMut(u64) -> SymbolResult<&'a [u8]>,
) -> SymbolResult<RadarSeriesStyle> {
    let mut uniform = None;
    for &style in styles {
        let value = style_of(style, &mut payload)?;
        if uniform.is_some_and(|known| known != value) {
            return Err(SymbolError::Invalid(
                "a Radar chart has mixed series styles",
            ));
        }
        uniform = Some(value);
    }
    uniform.ok_or(SymbolError::Invalid("a Radar chart has no series"))
}

/// The rewritten payloads storing `style` in the series styles `styles`
/// (one per visible series, in series order), as the host writes them: the
/// series colors are read first, then each style whose effective value
/// differs is rewritten in order (a style two series share is rewritten
/// once). Returns the changed styles in order with their new payloads; the
/// caller must refuse a style another chart uses before publishing.
///
/// # Errors
///
/// No series, a series without a solid palette color, an invalid style or
/// parent chain, or a write that does not read back.
pub fn rewrites<'a>(
    styles: &[u64],
    style: RadarSeriesStyle,
    mut payload: impl FnMut(u64) -> SymbolResult<&'a [u8]>,
) -> SymbolResult<Vec<(u64, Vec<u8>)>> {
    if styles.is_empty() {
        return Err(SymbolError::Invalid("a Radar chart has no series"));
    }
    let colors = styles
        .iter()
        .map(|&identifier| color_of(identifier, &mut payload))
        .collect::<SymbolResult<Vec<_>>>()?;
    let mut changed: Vec<(u64, Vec<u8>)> = Vec::new();
    for (&identifier, color) in styles.iter().zip(colors) {
        let current = {
            let mut overlay = |object: u64| -> SymbolResult<&[u8]> {
                match changed.iter().find(|(known, _)| *known == object) {
                    Some((_, data)) => Ok(data.as_slice()),
                    None => payload(object),
                }
            };
            if style_of(identifier, &mut overlay).is_ok_and(|current| current == style) {
                continue;
            }
            overlay(identifier)?.to_vec()
        };
        let data = rewrite(&current, color, style)?;
        match changed.iter_mut().find(|(known, _)| *known == identifier) {
            Some((_, existing)) => *existing = data,
            None => changed.push((identifier, data)),
        }
    }
    let mut overlay = |object: u64| -> SymbolResult<&[u8]> {
        match changed.iter().find(|(known, _)| *known == object) {
            Some((_, data)) => Ok(data.as_slice()),
            None => payload(object),
        }
    };
    if effective(styles, &mut overlay)? != style {
        return Err(SymbolError::Invalid(
            "a Radar style write does not read back",
        ));
    }
    Ok(changed)
}

/// The uniform Radar style of the `series_count` visible series of chart
/// drawable `chart`, charged to `budget`.
pub(crate) fn read(
    workspace: &Workspace<'_>,
    chart: u64,
    series_count: usize,
    budget: &mut SymbolBudget,
) -> SymbolResult<RadarSeriesStyle> {
    let styles = bounded_style_slots(workspace, chart, series_count, budget)?;
    effective(&styles, workspace_payload(workspace, budget))
}

/// Store `style` in the effective styles of the `series_count` visible
/// series of chart drawable `chart` (each changed one exclusive to the
/// chart), admitting each changed style to the budget's scope. Returns
/// whether a payload changed.
///
/// # Errors
///
/// [`Refusal::SharedStyle`], an invalid style, an exhausted limit, or a
/// write that does not read back.
pub(crate) fn write(
    workspace: &mut Workspace<'_>,
    chart: u64,
    series_count: usize,
    style: RadarSeriesStyle,
    budget: &mut SymbolBudget,
) -> SymbolResult<bool> {
    let styles = bounded_style_slots(workspace, chart, series_count, budget)?;
    let rewritten = rewrites(&styles, style, workspace_payload(workspace, budget))?;
    for (identifier, _) in &rewritten {
        if bounded_style_owner_count(workspace, *identifier, budget)? != 1 {
            return Err(SymbolError::Refused(Refusal::SharedStyle));
        }
    }
    let changed = !rewritten.is_empty();
    for (identifier, data) in rewritten {
        budget.work(1)?;
        let location = workspace.locate(identifier).map_err(SymbolError::from)?;
        let member = workspace.member(location).to_owned();
        if let Some(scope) = budget.scope_mut() {
            scope.allow(identifier, &member, SERIES_STYLE_MESSAGE_TYPE);
        }
        let (message, payload) = bounded_payload(
            workspace,
            identifier,
            SERIES_STYLE_MESSAGE_TYPE,
            "a series style must have exactly one payload",
            budget,
        )?;
        let data = charged_patch(payload, budget, |_| Ok(data))?;
        workspace
            .replace_payload(identifier, message, data)
            .map_err(SymbolError::from)?;
    }
    Ok(changed)
}

#[cfg(test)]
mod tests {
    use litchi_iwa_common::color::RgbColorSpace;

    use super::*;

    const UNKNOWN_OUTER_FIELD: u32 = 4_096;
    const UNKNOWN_EXTENSION_FIELD: u32 = 4_097;

    fn color() -> Rgba {
        Rgba::new(0.2, 0.5, 0.8, 1.0, RgbColorSpace::Srgb).unwrap()
    }

    fn payload(extension: &[u8]) -> Vec<u8> {
        let mut outer = Vec::new();
        wire::append_length_delimited_field(
            &mut outer,
            codec::EXTENSION_FIELD,
            extension,
            litchi_iwa_common::WireLimits::DEFAULT,
        )
        .unwrap();
        wire::append_varint_field(&mut outer, UNKNOWN_OUTER_FIELD, 42).unwrap();
        outer
    }

    #[test]
    fn every_style_reads_back_and_keeps_unknown_fields() {
        let mut extension = Vec::new();
        wire::append_length_delimited_field(
            &mut extension,
            FILL_FIELD,
            &drawing::encode_fill(&Fill::Solid(color())),
            litchi_iwa_common::WireLimits::DEFAULT,
        )
        .unwrap();
        wire::append_varint_field(&mut extension, FILL_USES_STROKE_FIELD, 1).unwrap();
        wire::append_varint_field(&mut extension, UNKNOWN_EXTENSION_FIELD, 7).unwrap();
        let original = payload(&extension);
        for style in [
            RadarSeriesStyle::Fill,
            RadarSeriesStyle::Stroke,
            RadarSeriesStyle::FillAndStroke,
        ] {
            let patched = rewrite(&original, color(), style).unwrap();
            assert_eq!(effective(&[1], |_| Ok(patched.as_slice())).unwrap(), style);
            let properties = decode(&patched).unwrap();
            assert!(
                properties
                    .varint(UNKNOWN_EXTENSION_FIELD)
                    .unwrap()
                    .is_some()
            );
            assert!(
                wire::WireView::parse(&patched, litchi_iwa_common::WireLimits::DEFAULT)
                    .unwrap()
                    .fields()
                    .any(|field| field.number() == UNKNOWN_OUTER_FIELD)
            );
        }
    }

    #[test]
    fn inconsistent_and_malformed_states_refuse() {
        assert!(classify(&Fill::None, None, false).is_err());
        assert!(classify(&Fill::Solid(color()), None, true).is_err());
        let mut extension = Vec::new();
        wire::append_varint_field(&mut extension, FILL_USES_STROKE_FIELD, 2).unwrap();
        let malformed = payload(&extension);
        assert!(effective(&[1], |_| Ok(malformed.as_slice())).is_err());
    }
}
