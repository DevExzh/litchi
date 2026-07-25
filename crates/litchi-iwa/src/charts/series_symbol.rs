//! Strict, lossless data-symbol CRUD for native 2D chart series.
//!
//! Visibility and shape live in sparse per-series non-style objects, while
//! optional point size is inherited through the series-style parent chain.

use prost::Message;

use crate::charts::ChartKind;
use crate::charts::series_non_style::{
    chart_series_non_style_values, generated_chart_series_non_style_extension,
    patch_chart_series_non_style_extension, set_chart_series_non_style_values,
};
use crate::charts::series_style::{
    GENERATED_CHART_SERIES_STYLE_EXTENSION_FIELD, effective_chart_series_style_slots,
    generated_chart_series_style_extension,
};
use crate::protobuf::tsch;
use crate::wire::{patch_fixed32_field, patch_length_delimited_field, patch_varint_field};
use crate::{Error, IWorkPackage, Result};

const AREA_SHOW_FIELD: u32 = 32;
const LINE_SHOW_FIELD: u32 = 33;
const MIXED_AREA_SHOW_FIELD: u32 = 34;
const MIXED_LINE_SHOW_FIELD: u32 = 35;
const SCATTER_SHOW_FIELD: u32 = 36;
const RADAR_SHOW_FIELD: u32 = 160;

const AREA_TYPE_FIELD: u32 = 47;
const LINE_TYPE_FIELD: u32 = 48;
const MIXED_AREA_TYPE_FIELD: u32 = 49;
const MIXED_LINE_TYPE_FIELD: u32 = 50;
const SCATTER_TYPE_FIELD: u32 = 51;
const RADAR_TYPE_FIELD: u32 = 163;

const AREA_SIZE_FIELD: u32 = 70;
const LINE_SIZE_FIELD: u32 = 71;
const MIXED_AREA_SIZE_FIELD: u32 = 72;
const MIXED_LINE_SIZE_FIELD: u32 = 73;
const SCATTER_SIZE_FIELD: u32 = 74;
const RADAR_SIZE_FIELD: u32 = 181;

/// Marker shapes offered by the native iWork data-symbol menu.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum ChartSeriesSymbolShape {
    #[default]
    Circle,
    Triangle,
    InvertedTriangle,
    Diamond,
    Square,
}

impl ChartSeriesSymbolShape {
    const fn native(self) -> i32 {
        match self {
            Self::Circle => 0,
            Self::Triangle => 2,
            Self::InvertedTriangle => 3,
            Self::Diamond => 5,
            Self::Square => 4,
        }
    }

    fn from_native(value: i32) -> Result<Self> {
        match value {
            0 => Ok(Self::Circle),
            2 => Ok(Self::Triangle),
            3 => Ok(Self::InvertedTriangle),
            4 => Ok(Self::Square),
            5 => Ok(Self::Diamond),
            _ => Err(Error::InvalidFormat(format!(
                "unsupported native chart series symbol type {value}"
            ))),
        }
    }
}

/// Explicit marker size in points. An absent size means native automatic sizing.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct ChartSeriesSymbolSize(f32);

impl ChartSeriesSymbolSize {
    pub fn new(points: f32) -> Result<Self> {
        if !points.is_finite() || points <= 0.0 {
            return Err(Error::InvalidFormat(format!(
                "chart series symbol size must be finite and positive, got {points}"
            )));
        }
        Ok(Self(points))
    }

    pub const fn points(self) -> f32 {
        self.0
    }
}

/// One visible data symbol. `size: None` preserves native automatic sizing.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ChartSeriesSymbol {
    pub shape: ChartSeriesSymbolShape,
    pub size: Option<ChartSeriesSymbolSize>,
}

impl ChartSeriesSymbol {
    pub const fn automatic(shape: ChartSeriesSymbolShape) -> Self {
        Self { shape, size: None }
    }

    pub const fn sized(shape: ChartSeriesSymbolShape, size: ChartSeriesSymbolSize) -> Self {
        Self {
            shape,
            size: Some(size),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct SymbolNonStyle {
    visible: bool,
    shape: ChartSeriesSymbolShape,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ChartSeriesSymbolKind {
    Area2d,
    Line2d,
    MixedArea2d,
    MixedLine2d,
    Radar2d,
    Scatter2d,
}

impl ChartSeriesSymbolKind {
    pub fn for_chart_kind(kind: ChartKind) -> Result<Self> {
        match kind {
            ChartKind::Area2d | ChartKind::StackedArea2d => Ok(Self::Area2d),
            ChartKind::Line2d => Ok(Self::Line2d),
            ChartKind::Radar2d => Ok(Self::Radar2d),
            ChartKind::Scatter2d => Ok(Self::Scatter2d),
            _ => Err(Error::InvalidFormat(format!(
                "chart kind {kind:?} has no unambiguous data-symbol family"
            ))),
        }
    }

    const fn show_field(self) -> u32 {
        match self {
            Self::Area2d => AREA_SHOW_FIELD,
            Self::Line2d => LINE_SHOW_FIELD,
            Self::MixedArea2d => MIXED_AREA_SHOW_FIELD,
            Self::MixedLine2d => MIXED_LINE_SHOW_FIELD,
            Self::Radar2d => RADAR_SHOW_FIELD,
            Self::Scatter2d => SCATTER_SHOW_FIELD,
        }
    }

    const fn type_field(self) -> u32 {
        match self {
            Self::Area2d => AREA_TYPE_FIELD,
            Self::Line2d => LINE_TYPE_FIELD,
            Self::MixedArea2d => MIXED_AREA_TYPE_FIELD,
            Self::MixedLine2d => MIXED_LINE_TYPE_FIELD,
            Self::Radar2d => RADAR_TYPE_FIELD,
            Self::Scatter2d => SCATTER_TYPE_FIELD,
        }
    }

    const fn size_field(self) -> u32 {
        match self {
            Self::Area2d => AREA_SIZE_FIELD,
            Self::Line2d => LINE_SIZE_FIELD,
            Self::MixedArea2d => MIXED_AREA_SIZE_FIELD,
            Self::MixedLine2d => MIXED_LINE_SIZE_FIELD,
            Self::Radar2d => RADAR_SIZE_FIELD,
            Self::Scatter2d => SCATTER_SIZE_FIELD,
        }
    }
}

pub(crate) fn chart_series_symbols(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    kind: ChartKind,
    series_count: usize,
) -> Result<Vec<Option<ChartSeriesSymbol>>> {
    let storage = ChartSeriesSymbolKind::for_chart_kind(kind)?;
    let default = SymbolNonStyle {
        visible: false,
        shape: ChartSeriesSymbolShape::Circle,
    };
    let non_styles = chart_series_non_style_values(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        series_count,
        default,
        |data| read_non_style(data, storage),
    )?;
    let slots = effective_chart_series_style_slots(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        series_count,
    )?;
    non_styles
        .into_iter()
        .zip(slots)
        .map(|(non_style, slot)| {
            if !non_style.visible {
                return Ok(None);
            }
            let size = slot.read_inherited(package, |data| read_local_size(data, storage))?;
            Ok(Some(ChartSeriesSymbol {
                shape: non_style.shape,
                size,
            }))
        })
        .collect()
}

#[allow(clippy::too_many_arguments)]
pub(crate) fn set_chart_series_symbols(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    kind: ChartKind,
    series_count: usize,
    expected: &[Option<ChartSeriesSymbol>],
) -> Result<()> {
    if expected.len() != series_count {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} has {series_count} series, not {} symbol settings",
            expected.len()
        )));
    }
    let storage = ChartSeriesSymbolKind::for_chart_kind(kind)?;
    let default = SymbolNonStyle {
        visible: false,
        shape: ChartSeriesSymbolShape::Circle,
    };
    let non_styles = expected
        .iter()
        .map(|symbol| {
            symbol.map_or(default, |symbol| SymbolNonStyle {
                visible: true,
                shape: symbol.shape,
            })
        })
        .collect::<Vec<_>>();
    set_chart_series_non_style_values(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        "data symbols",
        &non_styles,
        default,
        |data| read_non_style(data, storage),
        |data, value| patch_non_style(data, storage, *value),
    )?;

    let slots = effective_chart_series_style_slots(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        series_count,
    )?;
    for (slot, symbol) in slots.iter().zip(expected) {
        let replacement = symbol.and_then(|symbol| symbol.size);
        let current = slot.read_inherited(package, |data| read_local_size(data, storage))?;
        if current == replacement {
            continue;
        }
        slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
        slot.update(package, |data| patch_local_size(data, storage, replacement))?;
    }
    if chart_series_symbols(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        kind,
        series_count,
    )? != expected
    {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} data-symbol update failed validation"
        )));
    }
    Ok(())
}

fn read_non_style(data: &[u8], storage: ChartSeriesSymbolKind) -> Result<SymbolNonStyle> {
    let Some(extension) = generated_chart_series_non_style_extension(data)? else {
        return Ok(SymbolNonStyle {
            visible: false,
            shape: ChartSeriesSymbolShape::Circle,
        });
    };
    let generated = tsch::generated::ChartSeriesNonStyleArchive::decode(extension)?;
    let visible = bool_field(&generated, storage.show_field()).unwrap_or(false);
    let shape = ChartSeriesSymbolShape::from_native(
        int_field(&generated, storage.type_field()).unwrap_or(0),
    )?;
    Ok(SymbolNonStyle { visible, shape })
}

fn patch_non_style(
    data: &[u8],
    storage: ChartSeriesSymbolKind,
    value: SymbolNonStyle,
) -> Result<Vec<u8>> {
    let existing = generated_chart_series_non_style_extension(data)?;
    let extension = existing.unwrap_or_default();
    let generated = tsch::generated::ChartSeriesNonStyleArchive::decode(extension)?;
    let visible_present = bool_field(&generated, storage.show_field()).is_some();
    let type_present = int_field(&generated, storage.type_field()).is_some();
    let extension = patch_varint_field(
        extension,
        storage.show_field(),
        visible_present,
        value.visible.then_some(1),
    )?;
    let extension = patch_varint_field(
        &extension,
        storage.type_field(),
        type_present,
        (value.shape != ChartSeriesSymbolShape::Circle).then_some(value.shape.native() as u64),
    )?;
    let patched = patch_chart_series_non_style_extension(
        data,
        existing.is_some(),
        (!extension.is_empty()).then_some(extension.as_slice()),
    )?;
    if read_non_style(&patched, storage)? != value {
        return Err(Error::InvalidFormat(
            "chart series data-symbol wire patch failed validation".to_owned(),
        ));
    }
    Ok(patched)
}

fn read_local_size(
    data: &[u8],
    storage: ChartSeriesSymbolKind,
) -> Result<Option<ChartSeriesSymbolSize>> {
    let Some(extension) = generated_chart_series_style_extension(data)? else {
        return Ok(None);
    };
    let generated = tsch::generated::ChartSeriesStyleArchive::decode(extension)?;
    float_field(&generated, storage.size_field())
        .map(ChartSeriesSymbolSize::new)
        .transpose()
}

fn patch_local_size(
    data: &[u8],
    storage: ChartSeriesSymbolKind,
    size: Option<ChartSeriesSymbolSize>,
) -> Result<Vec<u8>> {
    let existing = generated_chart_series_style_extension(data)?;
    let extension = existing.unwrap_or_default();
    let generated = tsch::generated::ChartSeriesStyleArchive::decode(extension)?;
    let present = float_field(&generated, storage.size_field()).is_some();
    let extension = patch_fixed32_field(
        extension,
        storage.size_field(),
        present,
        size.map(|size| size.points().to_bits()),
    )?;
    patch_length_delimited_field(
        data,
        GENERATED_CHART_SERIES_STYLE_EXTENSION_FIELD,
        existing.is_some(),
        (!extension.is_empty()).then_some(extension.as_slice()),
    )
}

fn bool_field(generated: &tsch::generated::ChartSeriesNonStyleArchive, field: u32) -> Option<bool> {
    match field {
        AREA_SHOW_FIELD => generated.tschchartseriesareashowsymbol,
        LINE_SHOW_FIELD => generated.tschchartserieslineshowsymbol,
        MIXED_AREA_SHOW_FIELD => generated.tschchartseriesmixedareashowsymbol,
        MIXED_LINE_SHOW_FIELD => generated.tschchartseriesmixedlineshowsymbol,
        RADAR_SHOW_FIELD => generated.tschchartseriesradarareashowsymbol,
        SCATTER_SHOW_FIELD => generated.tschchartseriesscattershowsymbol,
        _ => None,
    }
}

fn int_field(generated: &tsch::generated::ChartSeriesNonStyleArchive, field: u32) -> Option<i32> {
    match field {
        AREA_TYPE_FIELD => generated.tschchartseriesareasymboltype,
        LINE_TYPE_FIELD => generated.tschchartserieslinesymboltype,
        MIXED_AREA_TYPE_FIELD => generated.tschchartseriesmixedareasymboltype,
        MIXED_LINE_TYPE_FIELD => generated.tschchartseriesmixedlinesymboltype,
        RADAR_TYPE_FIELD => generated.tschchartseriesradarareasymboltype,
        SCATTER_TYPE_FIELD => generated.tschchartseriesscattersymboltype,
        _ => None,
    }
}

fn float_field(generated: &tsch::generated::ChartSeriesStyleArchive, field: u32) -> Option<f32> {
    match field {
        AREA_SIZE_FIELD => generated.tschchartseriesareasymbolsize,
        LINE_SIZE_FIELD => generated.tschchartserieslinesymbolsize,
        MIXED_AREA_SIZE_FIELD => generated.tschchartseriesmixedareasymbolsize,
        MIXED_LINE_SIZE_FIELD => generated.tschchartseriesmixedlinesymbolsize,
        RADAR_SIZE_FIELD => generated.tschchartseriesradarareasymbolsize,
        SCATTER_SIZE_FIELD => generated.tschchartseriesscattersymbolsize,
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::protobuf::tss;
    use crate::wire::{append_varint_field, parse_wire_fields};

    #[test]
    fn symbol_shapes_match_native_wire_values() {
        for (shape, native) in [
            (ChartSeriesSymbolShape::Circle, 0),
            (ChartSeriesSymbolShape::Triangle, 2),
            (ChartSeriesSymbolShape::InvertedTriangle, 3),
            (ChartSeriesSymbolShape::Square, 4),
            (ChartSeriesSymbolShape::Diamond, 5),
        ] {
            assert_eq!(shape.native(), native);
            assert_eq!(ChartSeriesSymbolShape::from_native(native).unwrap(), shape);
        }
        assert!(ChartSeriesSymbolShape::from_native(1).is_err());
    }

    #[test]
    fn symbol_non_style_patch_preserves_unknown_fields() {
        let mut generated = Vec::new();
        append_varint_field(&mut generated, 4_097, 91).unwrap();
        let mut original = tsch::ChartSeriesNonStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();
        append_varint_field(&mut original, 4_096, 73).unwrap();
        original =
            patch_chart_series_non_style_extension(&original, false, Some(generated.as_slice()))
                .unwrap();
        let expected = SymbolNonStyle {
            visible: true,
            shape: ChartSeriesSymbolShape::Square,
        };
        let patched = patch_non_style(&original, ChartSeriesSymbolKind::Line2d, expected).unwrap();
        assert_eq!(
            read_non_style(&patched, ChartSeriesSymbolKind::Line2d).unwrap(),
            expected
        );
        assert!(
            parse_wire_fields(&patched)
                .unwrap()
                .iter()
                .any(|field| field.number == 4_096)
        );
        let extension = generated_chart_series_non_style_extension(&patched)
            .unwrap()
            .unwrap();
        assert!(
            parse_wire_fields(extension)
                .unwrap()
                .iter()
                .any(|field| field.number == 4_097)
        );
    }

    #[test]
    fn symbol_size_patch_distinguishes_auto_and_explicit_points() {
        let original = tsch::ChartSeriesStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();
        let size = ChartSeriesSymbolSize::new(18.0).unwrap();
        let explicit =
            patch_local_size(&original, ChartSeriesSymbolKind::Line2d, Some(size)).unwrap();
        assert_eq!(
            read_local_size(&explicit, ChartSeriesSymbolKind::Line2d).unwrap(),
            Some(size)
        );
        let automatic = patch_local_size(&explicit, ChartSeriesSymbolKind::Line2d, None).unwrap();
        assert_eq!(
            read_local_size(&automatic, ChartSeriesSymbolKind::Line2d).unwrap(),
            None
        );
    }

    #[test]
    fn symbol_size_rejects_non_finite_and_non_positive_values() {
        for value in [0.0, -1.0, f32::INFINITY, f32::NAN] {
            assert!(ChartSeriesSymbolSize::new(value).is_err());
        }
    }
}
