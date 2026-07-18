//! Shared construction and decoding primitives for source-built iWork charts.

mod build;
mod data;
mod ids;

pub(crate) use build::source_chart_objects;
pub(crate) use data::{
    chart_data, chart_geometry, chart_grid, drawable_geometry, geometry_archive, reference,
    require_creatable_kind,
};
pub(crate) use ids::SourceChartObjectIds;

use prost::Message;

use super::{ChartData, ChartKind, IWorkChartArchive};
use crate::archive::{ArchiveObject, RawMessage};
use crate::protobuf::{tn, tsch, tsd, tsk, tsp, tss};
use crate::shapes::{DrawableGeometry, DrawablePoint, DrawableSize};
use crate::wire::append_length_delimited_field;
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
const KEYNOTE_PARAGRAPH_STYLE_COUNT: usize = 25;
const SERIES_STYLE_COUNT: usize = 6;
const VALUE_AXIS_COUNT: usize = 2;

/// Native graph differences that cannot be shared across iWork applications.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum ChartApplicationProfile {
    Numbers,
    Keynote,
}

impl ChartApplicationProfile {
    const fn paragraph_style_count(self) -> usize {
        match self {
            Self::Numbers => NUMBERS_PARAGRAPH_STYLE_COUNT,
            Self::Keynote => KEYNOTE_PARAGRAPH_STYLE_COUNT,
        }
    }

    const fn uses_mediator(self) -> bool {
        matches!(self, Self::Numbers)
    }
}

fn deterministic_uuid(seed: u64) -> String {
    let suffix = seed & 0x0000_ffff_ffff_ffff;
    format!("00000000-0000-4000-8000-{suffix:012X}")
}
