//! Lossless inherited per-series fill CRUD for native charts.
//!
//! iWork stores chart-kind defaults in theme series styles and sparse
//! overrides in private child styles. Reads follow that parent chain while
//! writes touch only the selected style's kind-specific fill field.

use prost::Message;

use crate::charts::ChartKind;
use crate::charts::series_style::{
    ChartSeriesStyleSlot, GENERATED_CHART_SERIES_STYLE_EXTENSION_FIELD,
    effective_chart_series_style_slots, generated_chart_series_style_extension,
};
use crate::data_reference_registry::{
    add_component_data_reference, remove_component_data_reference,
};
use crate::package_metadata::component_identifier_for_entry;
use crate::protobuf::tsch;
use crate::shapes::{
    DrawableSize, RgbaColor, ShapeFill, ShapeImageDataIdentifier, ShapeImageFill,
    ShapeImageFillTechnique, fill_from_native, fill_to_native, image_data_identifier,
    remove_orphaned_image_asset, validate_image_asset,
};
use crate::wire::patch_length_delimited_field;
use crate::{Error, IWorkMediaEditor, IWorkPackage, MediaType, Result};

const AREA_3D_FILL_FIELD: u32 = 6;
const BAR_3D_FILL_FIELD: u32 = 7;
const COLUMN_3D_FILL_FIELD: u32 = 8;
const LINE_3D_FILL_FIELD: u32 = 9;
const PIE_3D_FILL_FIELD: u32 = 10;
const AREA_FILL_FIELD: u32 = 11;
const BAR_FILL_FIELD: u32 = 12;
const COLUMN_FILL_FIELD: u32 = 13;
const DEFAULT_FILL_FIELD: u32 = 14;
const PIE_FILL_FIELD: u32 = 17;
const RADAR_AREA_FILL_FIELD: u32 = 165;

/// Native series-fill family selected by a chart kind.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ChartSeriesFillKind {
    Area2d,
    Bar2d,
    Column2d,
    Pie2d,
    Radar2d,
    Area3d,
    Bar3d,
    Column3d,
    Line3d,
    Pie3d,
}

impl ChartSeriesFillKind {
    /// Resolve the fill family displayed by iWork for a chart kind.
    pub fn for_chart_kind(kind: ChartKind) -> Result<Self> {
        match kind {
            ChartKind::Area2d | ChartKind::StackedArea2d => Ok(Self::Area2d),
            ChartKind::Bar2d | ChartKind::StackedBar2d | ChartKind::MultiDataBar2d => {
                Ok(Self::Bar2d)
            },
            ChartKind::Column2d | ChartKind::StackedColumn2d | ChartKind::MultiDataColumn2d => {
                Ok(Self::Column2d)
            },
            ChartKind::Pie2d | ChartKind::Donut2d => Ok(Self::Pie2d),
            ChartKind::Radar2d => Ok(Self::Radar2d),
            ChartKind::Area3d | ChartKind::StackedArea3d => Ok(Self::Area3d),
            ChartKind::Bar3d | ChartKind::StackedBar3d => Ok(Self::Bar3d),
            ChartKind::Column3d | ChartKind::StackedColumn3d => Ok(Self::Column3d),
            ChartKind::Line3d => Ok(Self::Line3d),
            ChartKind::Pie3d | ChartKind::Donut3d => Ok(Self::Pie3d),
            _ => Err(Error::InvalidFormat(format!(
                "chart kind {kind:?} has no unambiguous series fill"
            ))),
        }
    }

    const fn field_number(self) -> u32 {
        match self {
            Self::Area2d => AREA_FILL_FIELD,
            Self::Bar2d => BAR_FILL_FIELD,
            Self::Column2d => COLUMN_FILL_FIELD,
            Self::Pie2d => PIE_FILL_FIELD,
            Self::Radar2d => RADAR_AREA_FILL_FIELD,
            Self::Area3d => AREA_3D_FILL_FIELD,
            Self::Bar3d => BAR_3D_FILL_FIELD,
            Self::Column3d => COLUMN_3D_FILL_FIELD,
            Self::Line3d => LINE_3D_FILL_FIELD,
            Self::Pie3d => PIE_3D_FILL_FIELD,
        }
    }
}

/// Read effective fills in native series order.
pub(crate) fn chart_series_fills(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    kind: ChartKind,
    series_count: usize,
) -> Result<Vec<ShapeFill>> {
    let storage = ChartSeriesFillKind::for_chart_kind(kind)?;
    effective_chart_series_style_slots(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        series_count,
    )?
    .iter()
    .map(|slot| read_effective_fill(slot, package, storage))
    .collect()
}

/// Replace every effective series fill in native series order.
pub(crate) fn set_chart_series_fills(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    kind: ChartKind,
    series_count: usize,
    expected: &[ShapeFill],
) -> Result<()> {
    if expected.len() != series_count {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} has {series_count} series, not {} fills",
            expected.len()
        )));
    }
    let storage = ChartSeriesFillKind::for_chart_kind(kind)?;
    let slots = effective_chart_series_style_slots(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        series_count,
    )?;
    let current = slots
        .iter()
        .map(|slot| read_effective_fill(slot, package, storage))
        .collect::<Result<Vec<_>>>()?;
    for fill in expected {
        validate_image_asset(package, fill)?;
    }
    for (slot, (current, replacement)) in slots.iter().zip(current.iter().zip(expected)) {
        if current == replacement {
            continue;
        }
        slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
        let old_local = slot.read(package, |data| read_local_fill(data, storage))?;
        slot.update(package, |data| {
            patch_local_fill(data, storage, Some(replacement))
        })?;
        adjust_data_reference(
            package,
            slot,
            old_local.as_ref().and_then(image_data_identifier),
            image_data_identifier(replacement),
        )?;
        remove_orphaned_image_asset(package, old_local.as_ref().and_then(image_data_identifier))?;
    }
    if chart_series_fills(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        kind,
        series_count,
    )? != expected
    {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} series fill update failed validation"
        )));
    }
    Ok(())
}

/// Remove one local fill override and expose its inherited effective fill.
pub(crate) fn reset_chart_series_fill(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    kind: ChartKind,
    series_count: usize,
    series_index: usize,
) -> Result<ShapeFill> {
    let storage = ChartSeriesFillKind::for_chart_kind(kind)?;
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
    let old_local = slot.read(package, |data| read_local_fill(data, storage))?;
    if old_local.is_none() {
        return read_effective_fill(slot, package, storage);
    }
    slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
    slot.update(package, |data| patch_local_fill(data, storage, None))?;
    adjust_data_reference(
        package,
        slot,
        old_local.as_ref().and_then(image_data_identifier),
        None,
    )?;
    remove_orphaned_image_asset(package, old_local.as_ref().and_then(image_data_identifier))?;
    read_effective_fill(slot, package, storage)
}

/// Embed image bytes and assign them to one series.
#[allow(clippy::too_many_arguments)]
pub(crate) fn set_chart_series_image_fill_data(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    kind: ChartKind,
    series_count: usize,
    series_index: usize,
    preferred_filename: &str,
    data: &[u8],
    technique: ShapeImageFillTechnique,
    fill_size: DrawableSize,
    tint: Option<RgbaColor>,
) -> Result<ShapeImageFill> {
    if series_index >= series_count {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} has {series_count} series, not series {}",
            series_index + 1
        )));
    }
    let mut media = IWorkMediaEditor::from_package(package.clone())?;
    let asset = media.insert_unreferenced(preferred_filename, data)?;
    if asset.media_type != MediaType::Image {
        return Err(Error::Bundle(format!(
            "Chart series image fills require image data, not {}",
            asset.media_type.name()
        )));
    }
    let mut staged = media.into_package();
    let mut image = ShapeImageFill::embedded(
        ShapeImageDataIdentifier::new(asset.data_identifier)?,
        technique,
        fill_size,
    )?;
    if let Some(tint) = tint {
        image = image.with_tint(tint);
    }
    let mut fills = chart_series_fills(
        &staged,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        kind,
        series_count,
    )?;
    fills[series_index] = ShapeFill::Image(image.clone());
    set_chart_series_fills(
        &mut staged,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        kind,
        series_count,
        &fills,
    )?;
    let package_path = asset.package_path.as_deref().ok_or_else(|| {
        Error::InvalidFormat("Chart series image-fill asset is not materialized".to_owned())
    })?;
    if staged.entry(package_path) != Some(data) {
        return Err(Error::InvalidFormat(
            "Chart series image-fill insertion failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(image)
}

fn read_effective_fill(
    slot: &ChartSeriesStyleSlot,
    package: &IWorkPackage,
    storage: ChartSeriesFillKind,
) -> Result<ShapeFill> {
    if let Some(fill) = slot.read_inherited(package, |data| {
        read_fill_field(data, storage.field_number())
    })? {
        return Ok(fill);
    }
    slot.read_inherited(package, |data| read_fill_field(data, DEFAULT_FILL_FIELD))?
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "chart series style {} has no inherited fill",
                slot.object_id()
            ))
        })
}

fn read_local_fill(data: &[u8], storage: ChartSeriesFillKind) -> Result<Option<ShapeFill>> {
    read_fill_field(data, storage.field_number())
}

fn read_fill_field(data: &[u8], field_number: u32) -> Result<Option<ShapeFill>> {
    let Some(extension) = generated_chart_series_style_extension(data)? else {
        return Ok(None);
    };
    let generated = tsch::generated::ChartSeriesStyleArchive::decode(extension)?;
    let native = match field_number {
        AREA_3D_FILL_FIELD => generated.tschchartseries3dareafill.as_ref(),
        BAR_3D_FILL_FIELD => generated.tschchartseries3dbarfill.as_ref(),
        COLUMN_3D_FILL_FIELD => generated.tschchartseries3dcolumnfill.as_ref(),
        LINE_3D_FILL_FIELD => generated.tschchartseries3dlinefill.as_ref(),
        PIE_3D_FILL_FIELD => generated.tschchartseries3dpiefill.as_ref(),
        AREA_FILL_FIELD => generated.tschchartseriesareafill.as_ref(),
        BAR_FILL_FIELD => generated.tschchartseriesbarfill.as_ref(),
        COLUMN_FILL_FIELD => generated.tschchartseriescolumnfill.as_ref(),
        DEFAULT_FILL_FIELD => generated.tschchartseriesdefaultfill.as_ref(),
        PIE_FILL_FIELD => generated.tschchartseriespiefill.as_ref(),
        RADAR_AREA_FILL_FIELD => generated.tschchartseriesradarareafill.as_ref(),
        _ => {
            return Err(Error::InvalidFormat(format!(
                "unsupported chart series fill field {field_number}"
            )));
        },
    };
    native.map(fill_from_native).transpose()
}

fn patch_local_fill(
    data: &[u8],
    storage: ChartSeriesFillKind,
    fill: Option<&ShapeFill>,
) -> Result<Vec<u8>> {
    let field_number = storage.field_number();
    let existing_extension = generated_chart_series_style_extension(data)?;
    let extension = existing_extension.unwrap_or_default();
    tsch::generated::ChartSeriesStyleArchive::decode(extension)?;
    let present = read_fill_field(data, field_number)?.is_some();
    let native = fill.map(|fill| fill_to_native(fill).encode_to_vec());
    let extension =
        patch_length_delimited_field(extension, field_number, present, native.as_deref())?;
    let patched = patch_length_delimited_field(
        data,
        GENERATED_CHART_SERIES_STYLE_EXTENSION_FIELD,
        existing_extension.is_some(),
        (!extension.is_empty()).then_some(extension.as_slice()),
    )?;
    if read_fill_field(&patched, field_number)? != fill.cloned() {
        return Err(Error::InvalidFormat(
            "chart series fill wire patch failed validation".to_owned(),
        ));
    }
    Ok(patched)
}

fn adjust_data_reference(
    package: &mut IWorkPackage,
    slot: &ChartSeriesStyleSlot,
    old_data_identifier: Option<u64>,
    new_data_identifier: Option<u64>,
) -> Result<()> {
    if old_data_identifier == new_data_identifier {
        return Ok(());
    }
    let component_id = component_identifier_for_entry(package, slot.archive_name())?
        .ok_or_else(|| Error::InvalidFormat("Chart series style has no component".to_owned()))?;
    if let Some(identifier) = old_data_identifier {
        remove_component_data_reference(package, component_id, identifier, slot.object_id())?;
    }
    if let Some(identifier) = new_data_identifier {
        add_component_data_reference(package, component_id, identifier, slot.object_id())?;
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::protobuf::tss;
    use crate::shapes::{RgbColorSpace, ShapeGradient, ShapeGradientAngle};
    use crate::wire::{append_varint_field, parse_wire_fields};

    const UNKNOWN_OUTER_FIELD: u32 = 4_096;
    const UNKNOWN_GENERATED_FIELD: u32 = 4_097;

    #[test]
    fn every_unambiguous_fill_kind_has_a_typed_storage_family() {
        for kind in [
            ChartKind::Area2d,
            ChartKind::StackedArea2d,
            ChartKind::Bar2d,
            ChartKind::StackedBar2d,
            ChartKind::MultiDataBar2d,
            ChartKind::Column2d,
            ChartKind::StackedColumn2d,
            ChartKind::MultiDataColumn2d,
            ChartKind::Pie2d,
            ChartKind::Donut2d,
            ChartKind::Radar2d,
            ChartKind::Area3d,
            ChartKind::StackedArea3d,
            ChartKind::Bar3d,
            ChartKind::StackedBar3d,
            ChartKind::Column3d,
            ChartKind::StackedColumn3d,
            ChartKind::Line3d,
            ChartKind::Pie3d,
            ChartKind::Donut3d,
        ] {
            assert!(ChartSeriesFillKind::for_chart_kind(kind).is_ok());
        }
        for kind in [
            ChartKind::Line2d,
            ChartKind::Scatter2d,
            ChartKind::Bubble2d,
            ChartKind::Mixed2d,
            ChartKind::TwoAxis2d,
        ] {
            assert!(ChartSeriesFillKind::for_chart_kind(kind).is_err());
        }
    }

    #[test]
    fn series_fill_patch_is_lossless_and_resets_exactly() {
        let original = style_with_unknown_fields(solid(0.2, 0.5, 0.8));
        let replacement = ShapeFill::Gradient(ShapeGradient::linear(
            color(0.8, 0.1, 0.2),
            color(0.1, 0.3, 0.9),
            ShapeGradientAngle::from_degrees(35.0).unwrap(),
        ));
        let patched =
            patch_local_fill(&original, ChartSeriesFillKind::Column2d, Some(&replacement)).unwrap();
        assert_eq!(
            read_local_fill(&patched, ChartSeriesFillKind::Column2d).unwrap(),
            Some(replacement)
        );
        assert_unknown_fields_retained(&original, &patched);
        let reset = patch_local_fill(&patched, ChartSeriesFillKind::Column2d, None).unwrap();
        assert_eq!(reset, original);
    }

    fn style_with_unknown_fields(default_fill: ShapeFill) -> Vec<u8> {
        let mut generated = tsch::generated::ChartSeriesStyleArchive {
            tschchartseriesdefaultfill: Some(fill_to_native(&default_fill)),
            ..Default::default()
        }
        .encode_to_vec();
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
        let original_outer = parse_wire_fields(original)
            .unwrap()
            .into_iter()
            .find(|field| field.number() == UNKNOWN_OUTER_FIELD)
            .map(|field| original[field.start()..field.end()].to_vec());
        let patched_outer = parse_wire_fields(patched)
            .unwrap()
            .into_iter()
            .find(|field| field.number() == UNKNOWN_OUTER_FIELD)
            .map(|field| patched[field.start()..field.end()].to_vec());
        assert_eq!(patched_outer, original_outer);
        let original_generated = generated_chart_series_style_extension(original)
            .unwrap()
            .unwrap();
        let patched_generated = generated_chart_series_style_extension(patched)
            .unwrap()
            .unwrap();
        let original_unknown = parse_wire_fields(original_generated)
            .unwrap()
            .into_iter()
            .find(|field| field.number() == UNKNOWN_GENERATED_FIELD)
            .map(|field| original_generated[field.start()..field.end()].to_vec());
        let patched_unknown = parse_wire_fields(patched_generated)
            .unwrap()
            .into_iter()
            .find(|field| field.number() == UNKNOWN_GENERATED_FIELD)
            .map(|field| patched_generated[field.start()..field.end()].to_vec());
        assert_eq!(patched_unknown, original_unknown);
    }

    fn solid(red: f32, green: f32, blue: f32) -> ShapeFill {
        ShapeFill::Solid(color(red, green, blue))
    }

    fn color(red: f32, green: f32, blue: f32) -> RgbaColor {
        RgbaColor::new(red, green, blue, 1.0, RgbColorSpace::Srgb).unwrap()
    }
}
