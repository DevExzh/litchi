//! Lossless inherited fill CRUD for native chart data symbols.
//!
//! A symbol fill is either a concrete drawing fill or a native request to use
//! the owning series' fill/stroke color. The three related wire fields are
//! treated as one tagged value so conflicting modes cannot leak into the API.

use prost::Message;

use litchi_iwa_common::media::Type as MediaType;

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
use crate::wire::{patch_length_delimited_field, patch_varint_field};
use crate::{Error, IWorkMediaEditor, IWorkPackage, Result};

const AREA_FILL_FIELD: u32 = 54;
const BUBBLE_FILL_FIELD: u32 = 55;
const LINE_FILL_FIELD: u32 = 56;
const RADAR_FILL_FIELD: u32 = 175;
const SCATTER_FILL_FIELD: u32 = 59;

const AREA_USE_SERIES_FILL_FIELD: u32 = 60;
const LINE_USE_SERIES_FILL_FIELD: u32 = 61;
const RADAR_USE_SERIES_FILL_FIELD: u32 = 177;

const AREA_USE_SERIES_STROKE_FIELD: u32 = 64;
const BUBBLE_USE_SERIES_STROKE_FIELD: u32 = 65;
const LINE_USE_SERIES_STROKE_FIELD: u32 = 66;
const RADAR_USE_SERIES_STROKE_FIELD: u32 = 179;
const SCATTER_USE_SERIES_STROKE_FIELD: u32 = 69;

/// Effective fill displayed inside one chart data symbol.
#[derive(Debug, Clone, PartialEq)]
pub enum ChartSeriesSymbolFill {
    /// A concrete no-fill, color, gradient, or image fill.
    Custom(ShapeFill),
    /// Follow the owning series' fill color.
    SeriesFill,
    /// Follow the owning series' stroke color.
    SeriesStroke,
}

impl From<ShapeFill> for ChartSeriesSymbolFill {
    fn from(fill: ShapeFill) -> Self {
        Self::Custom(fill)
    }
}

/// Native symbol-fill field family for an unambiguous 2D chart kind.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ChartSeriesSymbolFillKind {
    Area2d,
    Bubble2d,
    Line2d,
    Radar2d,
    Scatter2d,
}

#[derive(Debug, Clone, Copy)]
struct SymbolFillFields {
    fill: u32,
    use_series_fill: Option<u32>,
    use_series_stroke: u32,
}

impl ChartSeriesSymbolFillKind {
    pub fn for_chart_kind(kind: ChartKind) -> Result<Self> {
        match kind {
            ChartKind::Area2d | ChartKind::StackedArea2d => Ok(Self::Area2d),
            ChartKind::Bubble2d => Ok(Self::Bubble2d),
            ChartKind::Line2d => Ok(Self::Line2d),
            ChartKind::Radar2d => Ok(Self::Radar2d),
            ChartKind::Scatter2d => Ok(Self::Scatter2d),
            _ => Err(Error::InvalidFormat(format!(
                "chart kind {kind:?} has no unambiguous data-symbol fill family"
            ))),
        }
    }

    const fn fields(self) -> SymbolFillFields {
        match self {
            Self::Area2d => SymbolFillFields {
                fill: AREA_FILL_FIELD,
                use_series_fill: Some(AREA_USE_SERIES_FILL_FIELD),
                use_series_stroke: AREA_USE_SERIES_STROKE_FIELD,
            },
            Self::Bubble2d => SymbolFillFields {
                fill: BUBBLE_FILL_FIELD,
                use_series_fill: None,
                use_series_stroke: BUBBLE_USE_SERIES_STROKE_FIELD,
            },
            Self::Line2d => SymbolFillFields {
                fill: LINE_FILL_FIELD,
                use_series_fill: Some(LINE_USE_SERIES_FILL_FIELD),
                use_series_stroke: LINE_USE_SERIES_STROKE_FIELD,
            },
            Self::Radar2d => SymbolFillFields {
                fill: RADAR_FILL_FIELD,
                use_series_fill: Some(RADAR_USE_SERIES_FILL_FIELD),
                use_series_stroke: RADAR_USE_SERIES_STROKE_FIELD,
            },
            Self::Scatter2d => SymbolFillFields {
                fill: SCATTER_FILL_FIELD,
                use_series_fill: None,
                use_series_stroke: SCATTER_USE_SERIES_STROKE_FIELD,
            },
        }
    }

    const fn default_fill(self) -> ChartSeriesSymbolFill {
        match self {
            Self::Area2d | Self::Radar2d => ChartSeriesSymbolFill::SeriesFill,
            Self::Bubble2d | Self::Line2d | Self::Scatter2d => ChartSeriesSymbolFill::SeriesStroke,
        }
    }
}

pub(crate) fn chart_series_symbol_fills(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    kind: ChartKind,
    series_count: usize,
) -> Result<Vec<ChartSeriesSymbolFill>> {
    let storage = ChartSeriesSymbolFillKind::for_chart_kind(kind)?;
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

#[allow(clippy::too_many_arguments)]
pub(crate) fn set_chart_series_symbol_fills(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    kind: ChartKind,
    series_count: usize,
    expected: &[ChartSeriesSymbolFill],
) -> Result<()> {
    if expected.len() != series_count {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} has {series_count} series, not {} data-symbol fills",
            expected.len()
        )));
    }
    let storage = ChartSeriesSymbolFillKind::for_chart_kind(kind)?;
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
        if let ChartSeriesSymbolFill::Custom(fill) = fill {
            validate_image_asset(package, fill)?;
        }
    }
    for (slot, (current, replacement)) in slots.iter().zip(current.iter().zip(expected)) {
        if current == replacement {
            continue;
        }
        slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
        let old_image = slot.read(package, |data| {
            read_underlying_fill(data, storage)
                .map(|fill| fill.as_ref().and_then(image_data_identifier))
        })?;
        slot.update(package, |data| {
            patch_local_fill(data, storage, Some(replacement))
        })?;
        let new_image = match replacement {
            ChartSeriesSymbolFill::Custom(fill) => image_data_identifier(fill),
            ChartSeriesSymbolFill::SeriesFill | ChartSeriesSymbolFill::SeriesStroke => None,
        };
        adjust_data_reference(package, slot, old_image, new_image)?;
        remove_orphaned_image_asset(package, old_image)?;
    }
    if chart_series_symbol_fills(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        kind,
        series_count,
    )? != expected
    {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} data-symbol fill update failed validation"
        )));
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
pub(crate) fn reset_chart_series_symbol_fill(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    kind: ChartKind,
    series_count: usize,
    series_index: usize,
) -> Result<ChartSeriesSymbolFill> {
    let storage = ChartSeriesSymbolFillKind::for_chart_kind(kind)?;
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
    let old_image = slot.read(package, |data| {
        read_underlying_fill(data, storage)
            .map(|fill| fill.as_ref().and_then(image_data_identifier))
    })?;
    let local = slot.read(package, |data| read_local_fill(data, storage))?;
    if local.is_none() {
        return read_effective_fill(slot, package, storage);
    }
    slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
    slot.update(package, |data| patch_local_fill(data, storage, None))?;
    adjust_data_reference(package, slot, old_image, None)?;
    remove_orphaned_image_asset(package, old_image)?;
    read_effective_fill(slot, package, storage)
}

/// Embed image bytes and assign them to one data-symbol fill.
#[allow(clippy::too_many_arguments)]
pub(crate) fn set_chart_series_symbol_image_fill_data(
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
            "Chart data-symbol image fills require image data, not {}",
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
    let mut fills = chart_series_symbol_fills(
        &staged,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        kind,
        series_count,
    )?;
    fills[series_index] = ChartSeriesSymbolFill::Custom(ShapeFill::Image(image.clone()));
    set_chart_series_symbol_fills(
        &mut staged,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        kind,
        series_count,
        &fills,
    )?;
    let package_path = asset.package_path.as_deref().ok_or_else(|| {
        Error::InvalidFormat("Chart data-symbol image-fill asset is not materialized".to_owned())
    })?;
    if staged.entry(package_path) != Some(data) {
        return Err(Error::InvalidFormat(
            "Chart data-symbol image-fill insertion failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(image)
}

fn read_effective_fill(
    slot: &ChartSeriesStyleSlot,
    package: &IWorkPackage,
    storage: ChartSeriesSymbolFillKind,
) -> Result<ChartSeriesSymbolFill> {
    Ok(slot
        .read_inherited(package, |data| read_local_fill(data, storage))?
        .unwrap_or_else(|| storage.default_fill()))
}

fn read_local_fill(
    data: &[u8],
    storage: ChartSeriesSymbolFillKind,
) -> Result<Option<ChartSeriesSymbolFill>> {
    let Some(extension) = generated_chart_series_style_extension(data)? else {
        return Ok(None);
    };
    let generated = tsch::generated::ChartSeriesStyleArchive::decode(extension)?;
    let fields = storage.fields();
    let use_fill = fields
        .use_series_fill
        .and_then(|field| bool_field(&generated, field));
    let use_stroke = bool_field(&generated, fields.use_series_stroke);
    if use_fill == Some(true) && use_stroke == Some(true) {
        return Err(Error::InvalidFormat(
            "chart data-symbol fill uses both series fill and series stroke".to_owned(),
        ));
    }
    if use_fill == Some(true) {
        return Ok(Some(ChartSeriesSymbolFill::SeriesFill));
    }
    if use_stroke == Some(true) {
        return Ok(Some(ChartSeriesSymbolFill::SeriesStroke));
    }
    read_fill_archive(&generated, fields.fill)?
        .map(fill_from_native)
        .transpose()
        .map(|fill| fill.map(ChartSeriesSymbolFill::Custom))
}

fn read_underlying_fill(
    data: &[u8],
    storage: ChartSeriesSymbolFillKind,
) -> Result<Option<ShapeFill>> {
    let Some(extension) = generated_chart_series_style_extension(data)? else {
        return Ok(None);
    };
    let generated = tsch::generated::ChartSeriesStyleArchive::decode(extension)?;
    read_fill_archive(&generated, storage.fields().fill)?
        .map(fill_from_native)
        .transpose()
}

fn patch_local_fill(
    data: &[u8],
    storage: ChartSeriesSymbolFillKind,
    fill: Option<&ChartSeriesSymbolFill>,
) -> Result<Vec<u8>> {
    let existing = generated_chart_series_style_extension(data)?;
    let mut extension = existing.unwrap_or_default().to_vec();
    let generated = tsch::generated::ChartSeriesStyleArchive::decode(extension.as_slice())?;
    let fields = storage.fields();
    let fill_present = read_fill_archive(&generated, fields.fill)?.is_some();
    extension = patch_length_delimited_field(
        &extension,
        fields.fill,
        fill_present,
        match fill {
            Some(ChartSeriesSymbolFill::Custom(fill)) => Some(fill_to_native(fill).encode_to_vec()),
            _ => None,
        }
        .as_deref(),
    )?;
    if let Some(field) = fields.use_series_fill {
        extension = patch_varint_field(
            &extension,
            field,
            bool_field(&generated, field).is_some(),
            matches!(fill, Some(ChartSeriesSymbolFill::SeriesFill)).then_some(1),
        )?;
    } else if matches!(fill, Some(ChartSeriesSymbolFill::SeriesFill)) {
        return Err(Error::InvalidFormat(
            "this chart data-symbol family cannot use the series fill".to_owned(),
        ));
    }
    extension = patch_varint_field(
        &extension,
        fields.use_series_stroke,
        bool_field(&generated, fields.use_series_stroke).is_some(),
        matches!(fill, Some(ChartSeriesSymbolFill::SeriesStroke)).then_some(1),
    )?;
    let patched = patch_length_delimited_field(
        data,
        GENERATED_CHART_SERIES_STYLE_EXTENSION_FIELD,
        existing.is_some(),
        (!extension.is_empty()).then_some(extension.as_slice()),
    )?;
    if read_local_fill(&patched, storage)? != fill.cloned() {
        return Err(Error::InvalidFormat(
            "chart data-symbol fill wire patch failed validation".to_owned(),
        ));
    }
    Ok(patched)
}

fn read_fill_archive(
    generated: &tsch::generated::ChartSeriesStyleArchive,
    field: u32,
) -> Result<Option<&crate::protobuf::tsd::FillArchive>> {
    match field {
        AREA_FILL_FIELD => Ok(generated.tschchartseriesareasymbolfill.as_ref()),
        BUBBLE_FILL_FIELD => Ok(generated.tschchartseriesbubblesymbolfill.as_ref()),
        LINE_FILL_FIELD => Ok(generated.tschchartserieslinesymbolfill.as_ref()),
        RADAR_FILL_FIELD => Ok(generated.tschchartseriesradarareasymbolfill.as_ref()),
        SCATTER_FILL_FIELD => Ok(generated.tschchartseriesscattersymbolfill.as_ref()),
        _ => Err(Error::InvalidFormat(format!(
            "unsupported chart data-symbol fill field {field}"
        ))),
    }
}

fn bool_field(generated: &tsch::generated::ChartSeriesStyleArchive, field: u32) -> Option<bool> {
    match field {
        AREA_USE_SERIES_FILL_FIELD => generated.tschchartseriesareasymbolfilluseseriesfill,
        LINE_USE_SERIES_FILL_FIELD => generated.tschchartserieslinesymbolfilluseseriesfill,
        RADAR_USE_SERIES_FILL_FIELD => generated.tschchartseriesradarareasymbolfilluseseriesfill,
        AREA_USE_SERIES_STROKE_FIELD => generated.tschchartseriesareasymbolfilluseseriesstroke,
        BUBBLE_USE_SERIES_STROKE_FIELD => generated.tschchartseriesbubblesymbolfilluseseriesstroke,
        LINE_USE_SERIES_STROKE_FIELD => generated.tschchartserieslinesymbolfilluseseriesstroke,
        RADAR_USE_SERIES_STROKE_FIELD => {
            generated.tschchartseriesradarareasymbolfilluseseriesstroke
        },
        SCATTER_USE_SERIES_STROKE_FIELD => {
            generated.tschchartseriesscattersymbolfilluseseriesstroke
        },
        _ => None,
    }
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
    use crate::shapes::RgbColorSpace;
    use crate::wire::{append_varint_field, parse_wire_fields};
    use litchi_iwa_common::shape::fill::{Angle, Gradient};

    const UNKNOWN_OUTER_FIELD: u32 = 4_096;
    const UNKNOWN_GENERATED_FIELD: u32 = 4_097;

    #[test]
    fn symbol_fill_modes_round_trip_and_reset_exactly() {
        let original = style_with_unknown_fields();
        let modes = [
            ChartSeriesSymbolFill::SeriesFill,
            ChartSeriesSymbolFill::SeriesStroke,
            ChartSeriesSymbolFill::Custom(ShapeFill::None),
            ChartSeriesSymbolFill::Custom(ShapeFill::Solid(color(0.2, 0.5, 0.8))),
            ChartSeriesSymbolFill::Custom(ShapeFill::Gradient(Gradient::linear(
                color(0.8, 0.1, 0.2),
                color(0.1, 0.3, 0.9),
                Angle::from_degrees(35.0).unwrap(),
            ))),
        ];
        for mode in modes {
            let patched =
                patch_local_fill(&original, ChartSeriesSymbolFillKind::Line2d, Some(&mode))
                    .unwrap();
            assert_eq!(
                read_local_fill(&patched, ChartSeriesSymbolFillKind::Line2d).unwrap(),
                Some(mode)
            );
            assert_unknown_fields_retained(&original, &patched);
            assert_eq!(
                patch_local_fill(&patched, ChartSeriesSymbolFillKind::Line2d, None).unwrap(),
                original
            );
        }
    }

    #[test]
    fn conflicting_native_symbol_fill_modes_are_rejected() {
        let original = style_with_unknown_fields();
        let fields = ChartSeriesSymbolFillKind::Line2d.fields();
        let extension = generated_chart_series_style_extension(&original)
            .unwrap()
            .unwrap();
        let extension =
            patch_varint_field(extension, fields.use_series_fill.unwrap(), false, Some(1)).unwrap();
        let extension =
            patch_varint_field(&extension, fields.use_series_stroke, false, Some(1)).unwrap();
        let malformed = patch_length_delimited_field(
            &original,
            GENERATED_CHART_SERIES_STYLE_EXTENSION_FIELD,
            true,
            Some(extension.as_slice()),
        )
        .unwrap();
        assert!(read_local_fill(&malformed, ChartSeriesSymbolFillKind::Line2d).is_err());
    }

    fn style_with_unknown_fields() -> Vec<u8> {
        let mut generated = Vec::new();
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

    fn color(red: f32, green: f32, blue: f32) -> RgbaColor {
        RgbaColor::new(red, green, blue, 1.0, RgbColorSpace::Srgb).unwrap()
    }
}
