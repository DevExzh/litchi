//! Lossless inherited outline CRUD for native chart data symbols.
//!
//! An absent field inherits from the theme while an explicit empty stroke
//! hides the outline. Keeping those states distinct makes reset deterministic.

use prost::Message;

use crate::charts::ChartKind;
use crate::charts::series_stroke::ChartSeriesStroke;
use crate::charts::series_style::{
    ChartSeriesStyleSlot, GENERATED_CHART_SERIES_STYLE_EXTENSION_FIELD,
    effective_chart_series_style_slots, generated_chart_series_style_extension,
};
use crate::protobuf::{tsch, tsd};
use crate::shapes::empty_stroke_archive;
use crate::wire::patch_length_delimited_field;
use crate::{Error, IWorkPackage, Result};

const AREA_OUTLINE_FIELD: u32 = 75;
const BUBBLE_OUTLINE_FIELD: u32 = 76;
const LINE_OUTLINE_FIELD: u32 = 77;
const RADAR_OUTLINE_FIELD: u32 = 183;
const SCATTER_OUTLINE_FIELD: u32 = 80;

/// Native symbol-outline field family for an unambiguous 2D chart kind.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ChartSeriesSymbolOutlineKind {
    Area2d,
    Bubble2d,
    Line2d,
    Radar2d,
    Scatter2d,
}

impl ChartSeriesSymbolOutlineKind {
    pub fn for_chart_kind(kind: ChartKind) -> Result<Self> {
        match kind {
            ChartKind::Area2d | ChartKind::StackedArea2d => Ok(Self::Area2d),
            ChartKind::Bubble2d => Ok(Self::Bubble2d),
            ChartKind::Line2d => Ok(Self::Line2d),
            ChartKind::Radar2d => Ok(Self::Radar2d),
            ChartKind::Scatter2d => Ok(Self::Scatter2d),
            _ => Err(Error::InvalidFormat(format!(
                "chart kind {kind:?} has no unambiguous data-symbol outline family"
            ))),
        }
    }

    const fn field_number(self) -> u32 {
        match self {
            Self::Area2d => AREA_OUTLINE_FIELD,
            Self::Bubble2d => BUBBLE_OUTLINE_FIELD,
            Self::Line2d => LINE_OUTLINE_FIELD,
            Self::Radar2d => RADAR_OUTLINE_FIELD,
            Self::Scatter2d => SCATTER_OUTLINE_FIELD,
        }
    }
}

pub(crate) fn chart_series_symbol_outlines(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    kind: ChartKind,
    series_count: usize,
) -> Result<Vec<Option<ChartSeriesStroke>>> {
    let storage = ChartSeriesSymbolOutlineKind::for_chart_kind(kind)?;
    effective_chart_series_style_slots(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        series_count,
    )?
    .iter()
    .map(|slot| read_effective_outline(slot, package, storage))
    .collect()
}

#[allow(clippy::too_many_arguments)]
pub(crate) fn set_chart_series_symbol_outlines(
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
            "{drawable_label} chart {drawable_object_id} has {series_count} series, not {} data-symbol outlines",
            expected.len()
        )));
    }
    let storage = ChartSeriesSymbolOutlineKind::for_chart_kind(kind)?;
    let slots = effective_chart_series_style_slots(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        series_count,
    )?;
    let current = slots
        .iter()
        .map(|slot| read_effective_outline(slot, package, storage))
        .collect::<Result<Vec<_>>>()?;
    for (slot, (current, replacement)) in slots.iter().zip(current.iter().zip(expected)) {
        if current == replacement {
            continue;
        }
        slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
        slot.update(package, |data| {
            patch_local_outline(data, storage, Some(*replacement))
        })?;
    }
    if chart_series_symbol_outlines(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        kind,
        series_count,
    )? != expected
    {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} data-symbol outline update failed validation"
        )));
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
pub(crate) fn reset_chart_series_symbol_outline(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    kind: ChartKind,
    series_count: usize,
    series_index: usize,
) -> Result<Option<ChartSeriesStroke>> {
    let storage = ChartSeriesSymbolOutlineKind::for_chart_kind(kind)?;
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
    if read_local_outline(slot, package, storage)?.is_none() {
        return read_effective_outline(slot, package, storage);
    }
    slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
    slot.update(package, |data| patch_local_outline(data, storage, None))?;
    read_effective_outline(slot, package, storage)
}

fn read_effective_outline(
    slot: &ChartSeriesStyleSlot,
    package: &IWorkPackage,
    storage: ChartSeriesSymbolOutlineKind,
) -> Result<Option<ChartSeriesStroke>> {
    Ok(slot
        .read_inherited(package, |data| {
            read_outline_field(data, storage.field_number())
        })?
        .flatten())
}

fn read_local_outline(
    slot: &ChartSeriesStyleSlot,
    package: &IWorkPackage,
    storage: ChartSeriesSymbolOutlineKind,
) -> Result<Option<Option<ChartSeriesStroke>>> {
    slot.read(package, |data| {
        read_outline_field(data, storage.field_number())
    })
}

fn read_outline_field(data: &[u8], field_number: u32) -> Result<Option<Option<ChartSeriesStroke>>> {
    let Some(extension) = generated_chart_series_style_extension(data)? else {
        return Ok(None);
    };
    let generated = tsch::generated::ChartSeriesStyleArchive::decode(extension)?;
    outline_field(&generated, field_number)?
        .map(ChartSeriesStroke::from_native)
        .transpose()
}

fn outline_field(
    generated: &tsch::generated::ChartSeriesStyleArchive,
    field_number: u32,
) -> Result<Option<&tsd::StrokeArchive>> {
    match field_number {
        AREA_OUTLINE_FIELD => Ok(generated.tschchartseriesareasymbolstroke.as_ref()),
        BUBBLE_OUTLINE_FIELD => Ok(generated.tschchartseriesbubblesymbolstroke.as_ref()),
        LINE_OUTLINE_FIELD => Ok(generated.tschchartserieslinesymbolstroke.as_ref()),
        RADAR_OUTLINE_FIELD => Ok(generated.tschchartseriesradarareasymbolstroke.as_ref()),
        SCATTER_OUTLINE_FIELD => Ok(generated.tschchartseriesscattersymbolstroke.as_ref()),
        _ => Err(Error::InvalidFormat(format!(
            "unsupported chart data-symbol outline field {field_number}"
        ))),
    }
}

fn patch_local_outline(
    data: &[u8],
    storage: ChartSeriesSymbolOutlineKind,
    outline: Option<Option<ChartSeriesStroke>>,
) -> Result<Vec<u8>> {
    let field_number = storage.field_number();
    let existing_extension = generated_chart_series_style_extension(data)?;
    let extension = existing_extension.unwrap_or_default();
    tsch::generated::ChartSeriesStyleArchive::decode(extension)?;
    let present = read_outline_field(data, field_number)?.is_some();
    let native = outline.map(|outline| {
        outline
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
    if read_outline_field(&patched, field_number)? != outline {
        return Err(Error::InvalidFormat(
            "chart data-symbol outline wire patch failed validation".to_owned(),
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
    fn every_unambiguous_symbol_outline_kind_has_a_typed_storage_family() {
        for kind in [
            ChartKind::Area2d,
            ChartKind::StackedArea2d,
            ChartKind::Bubble2d,
            ChartKind::Line2d,
            ChartKind::Radar2d,
            ChartKind::Scatter2d,
        ] {
            assert!(ChartSeriesSymbolOutlineKind::for_chart_kind(kind).is_ok());
        }
        for kind in [
            ChartKind::Bar2d,
            ChartKind::Mixed2d,
            ChartKind::TwoAxis2d,
            ChartKind::Area3d,
            ChartKind::Line3d,
        ] {
            assert!(ChartSeriesSymbolOutlineKind::for_chart_kind(kind).is_err());
        }
    }

    #[test]
    fn symbol_outline_patch_distinguishes_empty_override_and_exact_reset() {
        let original = style_with_unknown_fields();
        let outline = ChartSeriesStroke::new(
            RgbaColor::new(0.1, 0.3, 0.8, 1.0, RgbColorSpace::Srgb).unwrap(),
            StrokeWidth::new(3.5).unwrap(),
            crate::charts::ChartSeriesStrokePattern::RoundedDash,
        );
        let visible = patch_local_outline(
            &original,
            ChartSeriesSymbolOutlineKind::Line2d,
            Some(Some(outline)),
        )
        .unwrap();
        assert_eq!(
            read_outline_field(&visible, LINE_OUTLINE_FIELD).unwrap(),
            Some(Some(outline))
        );
        let hidden =
            patch_local_outline(&visible, ChartSeriesSymbolOutlineKind::Line2d, Some(None))
                .unwrap();
        assert_eq!(
            read_outline_field(&hidden, LINE_OUTLINE_FIELD).unwrap(),
            Some(None)
        );
        assert_unknown_fields_retained(&original, &hidden);
        let reset =
            patch_local_outline(&hidden, ChartSeriesSymbolOutlineKind::Line2d, None).unwrap();
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
