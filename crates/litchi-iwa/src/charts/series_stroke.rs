//! Lossless inherited per-series stroke CRUD for native 2D charts.
//!
//! Native theme styles provide the baseline appearance while app-authored
//! private styles contain sparse overrides. An explicit empty stroke is
//! distinct from an absent field so callers can both hide and reset a stroke.

use prost::Message;

use crate::charts::ChartKind;
use crate::charts::series_style::{
    ChartSeriesStyleSlot, GENERATED_CHART_SERIES_STYLE_EXTENSION_FIELD,
    effective_chart_series_style_slots, generated_chart_series_style_extension,
};
use crate::protobuf::{tsch, tsd};
use crate::shapes::{
    RgbaColor, ShapeStroke, StrokeJoin, StrokePattern, StrokeWidth, empty_stroke_archive,
    stroke_from_native, stroke_to_native,
};
use crate::wire::patch_length_delimited_field;
use crate::{Error, IWorkPackage, Result};

const AREA_STROKE_FIELD: u32 = 45;
const BAR_STROKE_FIELD: u32 = 46;
const BUBBLE_STROKE_FIELD: u32 = 47;
const LINE_STROKE_FIELD: u32 = 48;
const PIE_STROKE_FIELD: u32 = 52;
const SCATTER_STROKE_FIELD: u32 = 53;
const RADAR_AREA_STROKE_FIELD: u32 = 172;

/// Stroke patterns exposed by the native chart-series inspector.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum ChartSeriesStrokePattern {
    #[default]
    Solid,
    MediumDash,
    RoundedDash,
}

/// Typed native stroke appearance for one chart series.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ChartSeriesStroke {
    pub color: RgbaColor,
    pub width: StrokeWidth,
    pub pattern: ChartSeriesStrokePattern,
}

impl ChartSeriesStroke {
    pub const fn new(
        color: RgbaColor,
        width: StrokeWidth,
        pattern: ChartSeriesStrokePattern,
    ) -> Self {
        Self {
            color,
            width,
            pattern,
        }
    }

    pub(crate) fn from_native(stroke: &tsd::StrokeArchive) -> Result<Option<Self>> {
        let Some(stroke) = stroke_from_native(stroke)? else {
            return Ok(None);
        };
        let pattern = match stroke.pattern {
            StrokePattern::Solid => ChartSeriesStrokePattern::Solid,
            StrokePattern::MediumDash => ChartSeriesStrokePattern::MediumDash,
            StrokePattern::RoundedDash => ChartSeriesStrokePattern::RoundedDash,
            StrokePattern::ShortDash | StrokePattern::LongDash => {
                return Err(Error::InvalidFormat(format!(
                    "unsupported native chart series stroke pattern {:?}",
                    stroke.pattern
                )));
            },
        };
        Ok(Some(Self::new(stroke.color, stroke.width, pattern)))
    }

    pub(crate) fn to_native(self) -> tsd::StrokeArchive {
        let pattern = match self.pattern {
            ChartSeriesStrokePattern::Solid => StrokePattern::Solid,
            ChartSeriesStrokePattern::MediumDash => StrokePattern::MediumDash,
            ChartSeriesStrokePattern::RoundedDash => StrokePattern::RoundedDash,
        };
        stroke_to_native(
            ShapeStroke::new(self.color, self.width, pattern).with_join(StrokeJoin::Miter),
        )
    }
}

/// Native series-stroke family selected by an unambiguous 2D chart kind.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ChartSeriesStrokeKind {
    Area2d,
    BarOrColumn2d,
    Bubble2d,
    Line2d,
    Pie2d,
    Radar2d,
    Scatter2d,
}

impl ChartSeriesStrokeKind {
    /// Resolve the stroke family displayed by iWork for a chart kind.
    pub fn for_chart_kind(kind: ChartKind) -> Result<Self> {
        match kind {
            ChartKind::Area2d | ChartKind::StackedArea2d => Ok(Self::Area2d),
            ChartKind::Bar2d
            | ChartKind::StackedBar2d
            | ChartKind::MultiDataBar2d
            | ChartKind::Column2d
            | ChartKind::StackedColumn2d
            | ChartKind::MultiDataColumn2d => Ok(Self::BarOrColumn2d),
            ChartKind::Bubble2d => Ok(Self::Bubble2d),
            ChartKind::Line2d => Ok(Self::Line2d),
            ChartKind::Pie2d | ChartKind::Donut2d => Ok(Self::Pie2d),
            ChartKind::Radar2d => Ok(Self::Radar2d),
            ChartKind::Scatter2d => Ok(Self::Scatter2d),
            _ => Err(Error::InvalidFormat(format!(
                "chart kind {kind:?} has no unambiguous series stroke"
            ))),
        }
    }

    const fn field_number(self) -> u32 {
        match self {
            Self::Area2d => AREA_STROKE_FIELD,
            Self::BarOrColumn2d => BAR_STROKE_FIELD,
            Self::Bubble2d => BUBBLE_STROKE_FIELD,
            Self::Line2d => LINE_STROKE_FIELD,
            Self::Pie2d => PIE_STROKE_FIELD,
            Self::Radar2d => RADAR_AREA_STROKE_FIELD,
            Self::Scatter2d => SCATTER_STROKE_FIELD,
        }
    }
}

/// Read effective strokes in native series order.
pub(crate) fn chart_series_strokes(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    kind: ChartKind,
    series_count: usize,
) -> Result<Vec<Option<ChartSeriesStroke>>> {
    let storage = ChartSeriesStrokeKind::for_chart_kind(kind)?;
    effective_chart_series_style_slots(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        series_count,
    )?
    .iter()
    .map(|slot| read_effective_stroke(slot, package, storage))
    .collect()
}

/// Replace every effective series stroke in native series order.
pub(crate) fn set_chart_series_strokes(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    kind: ChartKind,
    series_count: usize,
    expected: &[Option<ChartSeriesStroke>],
) -> Result<()> {
    if expected.len() != series_count {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} has {series_count} series, not {} strokes",
            expected.len()
        )));
    }
    let storage = ChartSeriesStrokeKind::for_chart_kind(kind)?;
    let slots = effective_chart_series_style_slots(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        series_count,
    )?;
    let current = slots
        .iter()
        .map(|slot| read_effective_stroke(slot, package, storage))
        .collect::<Result<Vec<_>>>()?;
    for (slot, (current, replacement)) in slots.iter().zip(current.iter().zip(expected)) {
        if current == replacement {
            continue;
        }
        slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
        slot.update(package, |data| {
            patch_local_stroke(data, storage, Some(*replacement))
        })?;
    }
    if chart_series_strokes(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        kind,
        series_count,
    )? != expected
    {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} series stroke update failed validation"
        )));
    }
    Ok(())
}

/// Remove one local stroke override and expose its inherited appearance.
pub(crate) fn reset_chart_series_stroke(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    kind: ChartKind,
    series_count: usize,
    series_index: usize,
) -> Result<Option<ChartSeriesStroke>> {
    let storage = ChartSeriesStrokeKind::for_chart_kind(kind)?;
    let slots = effective_chart_series_style_slots(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        series_count,
    )?;
    let slot = slots.get(series_index).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} has {series_count} series, not series {}",
            series_index + 1
        ))
    })?;
    if read_local_stroke(slot, package, storage)?.is_none() {
        return read_effective_stroke(slot, package, storage);
    }
    slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
    slot.update(package, |data| patch_local_stroke(data, storage, None))?;
    read_effective_stroke(slot, package, storage)
}

fn read_effective_stroke(
    slot: &ChartSeriesStyleSlot,
    package: &IWorkPackage,
    storage: ChartSeriesStrokeKind,
) -> Result<Option<ChartSeriesStroke>> {
    Ok(slot
        .read_inherited(package, |data| {
            read_stroke_field(data, storage.field_number())
        })?
        .flatten())
}

fn read_local_stroke(
    slot: &ChartSeriesStyleSlot,
    package: &IWorkPackage,
    storage: ChartSeriesStrokeKind,
) -> Result<Option<Option<ChartSeriesStroke>>> {
    slot.read(package, |data| {
        read_stroke_field(data, storage.field_number())
    })
}

fn read_stroke_field(data: &[u8], field_number: u32) -> Result<Option<Option<ChartSeriesStroke>>> {
    let Some(extension) = generated_chart_series_style_extension(data)? else {
        return Ok(None);
    };
    let generated = tsch::generated::ChartSeriesStyleArchive::decode(extension)?;
    let native = stroke_field(&generated, field_number)?;
    native.map(ChartSeriesStroke::from_native).transpose()
}

fn stroke_field(
    generated: &tsch::generated::ChartSeriesStyleArchive,
    field_number: u32,
) -> Result<Option<&tsd::StrokeArchive>> {
    match field_number {
        AREA_STROKE_FIELD => Ok(generated.tschchartseriesareastroke.as_ref()),
        BAR_STROKE_FIELD => Ok(generated.tschchartseriesbarstroke.as_ref()),
        BUBBLE_STROKE_FIELD => Ok(generated.tschchartseriesbubblestroke.as_ref()),
        LINE_STROKE_FIELD => Ok(generated.tschchartserieslinestroke.as_ref()),
        PIE_STROKE_FIELD => Ok(generated.tschchartseriespiestroke.as_ref()),
        SCATTER_STROKE_FIELD => Ok(generated.tschchartseriesscatterstroke.as_ref()),
        RADAR_AREA_STROKE_FIELD => Ok(generated.tschchartseriesradarareastroke.as_ref()),
        _ => Err(Error::InvalidFormat(format!(
            "unsupported chart series stroke field {field_number}"
        ))),
    }
}

fn patch_local_stroke(
    data: &[u8],
    storage: ChartSeriesStrokeKind,
    stroke: Option<Option<ChartSeriesStroke>>,
) -> Result<Vec<u8>> {
    let field_number = storage.field_number();
    let existing_extension = generated_chart_series_style_extension(data)?;
    let extension = existing_extension.unwrap_or_default();
    tsch::generated::ChartSeriesStyleArchive::decode(extension)?;
    let present = read_stroke_field(data, field_number)?.is_some();
    let native = stroke.map(|stroke| {
        stroke
            .map_or_else(empty_stroke_archive, ChartSeriesStroke::to_native)
            .encode_to_vec()
    });
    let extension =
        patch_length_delimited_field(extension, field_number, present, native.as_deref())?;
    let patched = patch_length_delimited_field(
        data,
        GENERATED_CHART_SERIES_STYLE_EXTENSION_FIELD,
        existing_extension.is_some(),
        (!extension.is_empty()).then_some(extension.as_slice()),
    )?;
    if read_stroke_field(&patched, field_number)? != stroke {
        return Err(Error::InvalidFormat(
            "chart series stroke wire patch failed validation".to_owned(),
        ));
    }
    Ok(patched)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::protobuf::tss;
    use crate::shapes::{RgbColorSpace, RgbaColor, StrokeWidth};
    use crate::wire::{append_varint_field, parse_wire_fields};

    const UNKNOWN_OUTER_FIELD: u32 = 4_096;
    const UNKNOWN_GENERATED_FIELD: u32 = 4_097;

    #[test]
    fn every_unambiguous_stroke_kind_has_a_typed_storage_family() {
        for kind in [
            ChartKind::Area2d,
            ChartKind::StackedArea2d,
            ChartKind::Bar2d,
            ChartKind::StackedBar2d,
            ChartKind::MultiDataBar2d,
            ChartKind::Column2d,
            ChartKind::StackedColumn2d,
            ChartKind::MultiDataColumn2d,
            ChartKind::Bubble2d,
            ChartKind::Line2d,
            ChartKind::Pie2d,
            ChartKind::Donut2d,
            ChartKind::Radar2d,
            ChartKind::Scatter2d,
        ] {
            assert!(ChartSeriesStrokeKind::for_chart_kind(kind).is_ok());
        }
        for kind in [
            ChartKind::Mixed2d,
            ChartKind::TwoAxis2d,
            ChartKind::Area3d,
            ChartKind::Bar3d,
            ChartKind::Column3d,
            ChartKind::Line3d,
            ChartKind::Pie3d,
        ] {
            assert!(ChartSeriesStrokeKind::for_chart_kind(kind).is_err());
        }
    }

    #[test]
    fn series_stroke_patch_distinguishes_empty_override_and_exact_reset() {
        let original = style_with_unknown_fields();
        let stroke = ChartSeriesStroke::new(
            RgbaColor::new(0.1, 0.3, 0.8, 1.0, RgbColorSpace::Srgb).unwrap(),
            StrokeWidth::new(3.5).unwrap(),
            ChartSeriesStrokePattern::RoundedDash,
        );
        let visible = patch_local_stroke(
            &original,
            ChartSeriesStrokeKind::BarOrColumn2d,
            Some(Some(stroke)),
        )
        .unwrap();
        assert_eq!(
            read_stroke_field(&visible, BAR_STROKE_FIELD).unwrap(),
            Some(Some(stroke))
        );
        let hidden =
            patch_local_stroke(&visible, ChartSeriesStrokeKind::BarOrColumn2d, Some(None)).unwrap();
        assert_eq!(
            read_stroke_field(&hidden, BAR_STROKE_FIELD).unwrap(),
            Some(None)
        );
        assert_unknown_fields_retained(&original, &hidden);
        let reset =
            patch_local_stroke(&hidden, ChartSeriesStrokeKind::BarOrColumn2d, None).unwrap();
        assert_eq!(reset, original);
    }

    fn style_with_unknown_fields() -> Vec<u8> {
        let mut generated = tsch::generated::ChartSeriesStyleArchive::default().encode_to_vec();
        append_varint_field(&mut generated, UNKNOWN_GENERATED_FIELD, 77).unwrap();
        let mut outer = tsch::ChartSeriesStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();
        outer = patch_length_delimited_field(
            &outer,
            GENERATED_CHART_SERIES_STYLE_EXTENSION_FIELD,
            false,
            Some(generated.as_slice()),
        )
        .unwrap();
        append_varint_field(&mut outer, UNKNOWN_OUTER_FIELD, 91).unwrap();
        outer
    }

    fn assert_unknown_fields_retained(original: &[u8], patched: &[u8]) {
        let outer = |data: &[u8]| {
            parse_wire_fields(data)
                .unwrap()
                .into_iter()
                .find(|field| field.number() == UNKNOWN_OUTER_FIELD)
                .map(|field| data[field.start()..field.end()].to_vec())
        };
        assert_eq!(outer(patched), outer(original));
        let original_generated = generated_chart_series_style_extension(original)
            .unwrap()
            .unwrap();
        let patched_generated = generated_chart_series_style_extension(patched)
            .unwrap()
            .unwrap();
        let generated = |data: &[u8]| {
            parse_wire_fields(data)
                .unwrap()
                .into_iter()
                .find(|field| field.number() == UNKNOWN_GENERATED_FIELD)
                .map(|field| data[field.start()..field.end()].to_vec())
        };
        assert_eq!(generated(patched_generated), generated(original_generated));
    }
}
