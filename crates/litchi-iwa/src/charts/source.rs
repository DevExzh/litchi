//! Shared construction and decoding primitives for source-built iWork charts.

use std::collections::{HashMap, HashSet};

mod build;
mod data;
mod ids;
mod stylesheet;

pub(crate) use build::source_chart_objects;
pub(crate) use data::{
    chart_data, chart_geometry, chart_grid, drawable_geometry, geometry_archive, reference,
    require_creatable_kind,
};
pub(crate) use ids::SourceChartObjectIds;
pub(crate) use stylesheet::{
    local_chart_style_ids, register_chart_styles, unregister_chart_styles,
    validate_chart_styles_registered,
};

use prost::Message;

use super::{ChartData, ChartGapSpacing, ChartKind, IWorkChartArchive};
use crate::archive::{ArchiveObject, RawMessage};
use crate::protobuf::{tn, tsch, tsd, tsk, tsp, tss};
use crate::shapes::{DrawableGeometry, DrawablePoint, DrawableSize};
use crate::wire::{
    append_length_delimited_field, append_varint_field, patch_varint_field,
    transform_length_delimited_fields_at_path,
};
use crate::{Error, Result};

pub(crate) const CHART_MESSAGE_TYPE: u32 = 5_021;
pub(crate) const CHART_MEDIATOR_MESSAGE_TYPE: u32 = 12_006;
pub(crate) const STANDIN_MESSAGE_TYPE: u32 = 3_097;
pub(crate) const CHART_PRESET_MESSAGE_TYPE: u32 = 5_020;
pub(crate) const CHART_STYLE_MESSAGE_TYPE: u32 = 5_022;
pub(crate) const CHART_NON_STYLE_MESSAGE_TYPE: u32 = 5_023;
pub(crate) const LEGEND_STYLE_MESSAGE_TYPE: u32 = 5_024;
pub(crate) const LEGEND_NON_STYLE_MESSAGE_TYPE: u32 = 5_025;
pub(crate) const AXIS_STYLE_MESSAGE_TYPE: u32 = 5_026;
pub(crate) const AXIS_NON_STYLE_MESSAGE_TYPE: u32 = 5_027;
pub(crate) const SERIES_STYLE_MESSAGE_TYPE: u32 = 5_028;
pub(crate) const SERIES_NON_STYLE_MESSAGE_TYPE: u32 = 5_029;

const CURRENT_STYLE_EXTENSION_FIELD: u32 = 10_000;
const CHART_SCENE_DEPTH_EXTENSION_FIELD: u32 = 10_002;
const CHART_APPEARANCE_PRESERVED_EXTENSION_FIELD: u32 = 10_023;
const CHART_PROPORTIONAL_CALLOUTS_EXTENSION_FIELD: u32 = 10_024;
const CHART_ROUNDED_CORNERS_EXTENSION_FIELD: u32 = 10_026;
const CHART_VALUE_LABEL_SPACING_EXTENSION_FIELD: u32 = 10_027;
const CHART_ERROR_BAR_SPACING_EXTENSION_FIELD: u32 = 10_028;
const CHART_STACKED_SUMMARY_LABELS_EXTENSION_FIELD: u32 = 10_029;
const CHART_CACHED_FORMATTERS_EXTENSION_FIELD: u32 = 10_030;
const STANDARD_MESSAGE_VERSION: &[u32] = &[1, 0, 5];
const STANDIN_MESSAGE_VERSION: &[u32] = &[10, 1, 0];
const DEFAULT_DRAWABLE_FLAGS: u32 = 3;
const DEFAULT_ROTATION_DEGREES: f32 = 0.0;
const DEFAULT_TEXT_WRAP_MARGIN_POINTS: f32 = 12.0;
const DEFAULT_TEXT_WRAP_ALPHA_THRESHOLD: f32 = 0.5;
const MEDIATOR_LOCAL_SERIES_SENTINEL: u32 = u32::MAX;
const MEDIATOR_REMOTE_SERIES_INDEX: u32 = 0;
const MEDIATOR_FORMULA_DIRECTION: i32 = 0;
const MEDIATOR_FORMULA_SCHEME: i32 = 0;
const DEFAULT_CHART_DATASET_INDEX: u32 = 0;
const DEFAULT_CHART_NUMBER_FORMAT_TYPE: u32 = 256;
const AUTOMATIC_CHART_DECIMAL_PLACES: u32 = 253;
const DEFAULT_CHART_NEGATIVE_STYLE: u32 = 0;
const NUMBERS_PARAGRAPH_STYLE_COUNT: usize = 30;
const PAGES_PARAGRAPH_STYLE_COUNT: usize = 30;
const KEYNOTE_PARAGRAPH_STYLE_COUNT: usize = 25;
const SERIES_STYLE_COUNT: usize = 6;
const VALUE_AXIS_COUNT: usize = 2;

/// Native graph differences that cannot be shared across iWork applications.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum ChartApplicationProfile {
    Numbers,
    Pages,
    Keynote,
}

impl ChartApplicationProfile {
    const fn paragraph_style_count(self) -> usize {
        match self {
            Self::Numbers => NUMBERS_PARAGRAPH_STYLE_COUNT,
            Self::Pages => PAGES_PARAGRAPH_STYLE_COUNT,
            Self::Keynote => KEYNOTE_PARAGRAPH_STYLE_COUNT,
        }
    }

    const fn uses_mediator(self) -> bool {
        matches!(self, Self::Numbers)
    }

    const fn owns_preset(self) -> bool {
        matches!(self, Self::Pages)
    }
}

/// Remap a wire-preserved native chart style preset.
///
/// Style presets contain the private style-object references that accompany a
/// chart clone. Unknown fields are retained verbatim, while metadata prevents
/// silently retaining an unhandled private reference in an opaque extension.
pub(crate) fn remap_chart_preset_wire(
    data: &[u8],
    recorded_references: &[u64],
    remap: &HashMap<u64, u64>,
) -> Result<Vec<u8>> {
    const REFERENCE_FIELDS: &[u32] = &[1, 2, 3, 4, 5, 6];

    let mut expected = tsch::ChartStylePreset::decode(data)?;
    let mut typed_references = HashSet::new();
    collect_optional_preset_reference(&mut typed_references, expected.chart_style.as_ref());
    collect_optional_preset_reference(&mut typed_references, expected.legend_style.as_ref());
    collect_preset_references(&mut typed_references, &expected.value_axis_styles);
    collect_preset_references(&mut typed_references, &expected.category_axis_styles);
    collect_preset_references(&mut typed_references, &expected.series_styles);
    collect_preset_references(&mut typed_references, &expected.paragraph_styles);
    if let Some(identifier) = recorded_references
        .iter()
        .copied()
        .find(|identifier| remap.contains_key(identifier) && !typed_references.contains(identifier))
    {
        return Err(Error::InvalidFormat(format!(
            "chart preset payload has an unrecognized private reference {identifier}"
        )));
    }

    remap_optional_preset_reference(&mut expected.chart_style, remap);
    remap_optional_preset_reference(&mut expected.legend_style, remap);
    remap_preset_references(&mut expected.value_axis_styles, remap);
    remap_preset_references(&mut expected.category_axis_styles, remap);
    remap_preset_references(&mut expected.series_styles, remap);
    remap_preset_references(&mut expected.paragraph_styles, remap);

    let mut rewritten = data.to_vec();
    for field in REFERENCE_FIELDS {
        rewritten =
            transform_length_delimited_fields_at_path(&rewritten, &[*field], |raw_reference| {
                let reference = tsp::Reference::decode(raw_reference)?;
                let Some(identifier) = remap.get(&reference.identifier).copied() else {
                    return Ok(raw_reference.to_vec());
                };
                let rewritten = patch_varint_field(raw_reference, 1, true, Some(identifier))?;
                if tsp::Reference::decode(rewritten.as_slice())?.identifier != identifier {
                    return Err(Error::InvalidFormat(
                        "chart preset reference wire remap failed validation".to_owned(),
                    ));
                }
                Ok(rewritten)
            })?;
    }
    if tsch::ChartStylePreset::decode(rewritten.as_slice())? != expected {
        return Err(Error::InvalidFormat(
            "chart preset wire remap failed validation".to_owned(),
        ));
    }
    Ok(rewritten)
}

fn remap_preset_reference(reference: &mut tsp::Reference, remap: &HashMap<u64, u64>) {
    if let Some(identifier) = remap.get(&reference.identifier) {
        reference.identifier = *identifier;
    }
}

fn remap_optional_preset_reference(
    reference: &mut Option<tsp::Reference>,
    remap: &HashMap<u64, u64>,
) {
    if let Some(reference) = reference {
        remap_preset_reference(reference, remap);
    }
}

fn remap_preset_references(references: &mut [tsp::Reference], remap: &HashMap<u64, u64>) {
    for reference in references {
        remap_preset_reference(reference, remap);
    }
}

fn collect_preset_reference(identifiers: &mut HashSet<u64>, reference: &tsp::Reference) {
    if reference.identifier != 0 {
        identifiers.insert(reference.identifier);
    }
}

fn collect_optional_preset_reference(
    identifiers: &mut HashSet<u64>,
    reference: Option<&tsp::Reference>,
) {
    if let Some(reference) = reference {
        collect_preset_reference(identifiers, reference);
    }
}

fn collect_preset_references(identifiers: &mut HashSet<u64>, references: &[tsp::Reference]) {
    for reference in references {
        collect_preset_reference(identifiers, reference);
    }
}

fn deterministic_uuid(seed: u64) -> String {
    let suffix = seed & 0x0000_ffff_ffff_ffff;
    format!("00000000-0000-4000-8000-{suffix:012X}")
}
