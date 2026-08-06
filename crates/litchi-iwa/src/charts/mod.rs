//! Chart Support for iWork Documents
//!
//! This module provides support for extracting metadata and content from
//! charts in iWork documents (Numbers, Pages, Keynote).
//!
//! Charts in iWork contain:
//! - Chart titles and axis labels
//! - Data series names
//! - Legend text
//! - Grid data (row/column names and values)

use crate::{Error, IWorkPackage, Result};

mod archive;

/// Native protobuf-backed chart data for advanced format-level integrations.
pub mod raw {
    pub use super::archive::IWorkChartArchive;
}

pub(crate) mod arrangement;
pub(crate) mod axis;
pub(crate) mod axis_bounds;
pub(crate) mod axis_gridline_stroke;
pub(crate) mod axis_label_affixes;
pub(crate) mod axis_label_angle;
pub(crate) mod axis_label_position_3d;
pub(crate) mod axis_number_format;
pub(crate) mod axis_scale;
pub(crate) mod axis_steps;
pub(crate) mod axis_style;
pub(crate) mod background_fill;
pub(crate) mod bar_shape_3d;
pub(crate) mod border;
pub(crate) mod border_stroke;
pub mod category_labels;
mod data;
pub(crate) mod depth_3d;
pub(crate) mod donut_inner_radius;
pub(crate) mod font;
pub(crate) mod gaps;
pub(crate) mod hidden_data;
pub use litchi_iwa_common::chart::kind::Kind;
pub(crate) mod legend_fill;
pub(crate) mod legend_font;
pub(crate) mod legend_frame;
pub(crate) mod legend_shadow;
pub(crate) mod legend_stroke;
pub(crate) mod legend_style;
pub(crate) mod lighting_3d;
pub(crate) mod metadata_extractor;
pub(crate) mod non_style;
pub(crate) mod number_format;
pub(crate) mod object_container;
pub(crate) mod options;
pub(crate) mod pie_label_distance;
pub(crate) mod pie_labels;
pub(crate) mod pie_leader_lines;
pub(crate) mod pie_start_angle;
pub(crate) mod pie_wedge_explosion;
pub(crate) mod radar_grid_shape;
pub(crate) mod radar_series_style;
pub(crate) mod radar_start_angle;
pub mod reference_line;
pub(crate) mod rounded_corners;
pub(crate) mod scene_3d;
pub(crate) mod series_connection_line;
pub(crate) mod series_error_bar_auto_fit;
pub(crate) mod series_error_bars;
pub(crate) mod series_fill;
pub(crate) mod series_gap_3d;
pub(crate) mod series_non_style;
pub(crate) mod series_stroke;
pub(crate) mod series_style;
pub(crate) mod series_symbol;
pub(crate) mod series_symbol_fill;
pub(crate) mod series_symbol_outline;
pub(crate) mod series_topology;
pub(crate) mod series_trendline;
pub(crate) mod series_value_label_affixes;
pub(crate) mod series_value_label_auto_fit;
pub(crate) mod series_value_label_location;
pub(crate) mod series_value_label_number_format;
pub(crate) mod series_value_labels;
pub(crate) mod shadow;
pub(crate) mod source;
pub(crate) mod style;

pub use litchi_iwa_common::chart::gaps::{Percentage, Spacing};

pub(crate) use archive::IWorkChartArchive;
pub use arrangement::ChartArrangement;
pub use axis_gridline_stroke::{ChartAxisGridline, ChartAxisGridlineStroke};
pub use bar_shape_3d::Chart3dBarShape;
pub use data::ChartData;
pub use depth_3d::Chart3dDepth;
pub use donut_inner_radius::ChartDonutInnerRadius;
pub use font::{ChartFont, ChartFontSize};
pub use legend_fill::ChartLegendFill;
pub use legend_font::{ChartLegendFont, ChartLegendFontSize};
pub use legend_frame::{
    ChartLegendCoordinate, ChartLegendExtent, ChartLegendFrame, ChartLegendRect,
};
pub use legend_shadow::ChartLegendShadow;
pub use legend_stroke::ChartLegendStroke;
pub use lighting_3d::Chart3dLightingStyle;
pub use litchi_iwa_common::chart::axis::{
    Axis, Bound, Bounds, LabelAngle, LabelPosition3d, MajorStepCount, MinorStepCount, Scale, Steps,
    TickMarkLocation,
};
pub use litchi_iwa_common::chart::number_format::{
    DecimalPlaces, FixedDecimalPlaces, LabelAffixes, NegativeStyle, NumberFormat,
};
pub use litchi_iwa_common::chart::pie::{
    LabelVisibility, LeaderLineVisibility, LeaderLineVisibilityKind,
};
pub use litchi_iwa_common::chart::series_labels::{Index, Visibility};
pub use litchi_iwa_common::chart::{Direction, DirectionKind};
pub use metadata_extractor::ChartMetadata;
pub use pie_label_distance::ChartPieLabelDistance;
pub use pie_start_angle::ChartPieStartAngle;
pub use pie_wedge_explosion::{ChartPieWedgeExplosion, ChartPieWedgeIndex};
pub use radar_grid_shape::ChartRadarGridShape;
pub use radar_series_style::ChartRadarSeriesStyle;
pub use radar_start_angle::ChartRadarStartAngle;
pub use rounded_corners::{ChartCornerRadius, ChartRoundedCorners};
pub use scene_3d::Chart3dRotation;
pub use series_connection_line::{ChartSeriesConnectionLine, ChartSeriesConnectionLineKind};
pub use series_error_bar_auto_fit::ChartSeriesErrorBarAutoFit;
pub use series_error_bars::{
    ChartErrorBarCustomValue, ChartErrorBarCustomValues, ChartErrorBarDirection,
    ChartErrorBarFixedValue, ChartErrorBarPercentage, ChartErrorBarStandardDeviationCount,
    ChartSeriesErrorBars,
};
pub use series_fill::ChartSeriesFillKind;
pub use series_gap_3d::Chart3dSeriesGap;
pub use series_stroke::{ChartSeriesStroke, ChartSeriesStrokeKind, ChartSeriesStrokePattern};
pub use series_symbol::{
    ChartSeriesSymbol, ChartSeriesSymbolKind, ChartSeriesSymbolShape, ChartSeriesSymbolSize,
};
pub use series_symbol_fill::{ChartSeriesSymbolFill, ChartSeriesSymbolFillKind};
pub use series_symbol_outline::ChartSeriesSymbolOutlineKind;
pub use series_trendline::{
    ChartSeriesTrendline, ChartSeriesTrendlineMovingAveragePeriod,
    ChartSeriesTrendlinePolynomialOrder, ChartSeriesTrendlineType,
};
pub use series_value_label_auto_fit::ChartSeriesValueLabelAutoFit;
pub use series_value_label_location::ChartSeriesValueLabelLocation;
pub use shadow::ChartShadow;

impl From<litchi_iwa_common::chart::axis::bounds::Error> for crate::Error {
    fn from(error: litchi_iwa_common::chart::axis::bounds::Error) -> Self {
        Self::InvalidFormat(error.to_string())
    }
}

impl From<litchi_iwa_common::chart::axis::label_angle::Error> for crate::Error {
    fn from(error: litchi_iwa_common::chart::axis::label_angle::Error) -> Self {
        Self::InvalidFormat(error.to_string())
    }
}

impl From<litchi_iwa_common::chart::axis::steps::Error> for crate::Error {
    fn from(error: litchi_iwa_common::chart::axis::steps::Error) -> Self {
        Self::InvalidFormat(error.to_string())
    }
}

impl From<litchi_iwa_common::chart::category_labels::Error> for crate::Error {
    fn from(error: litchi_iwa_common::chart::category_labels::Error) -> Self {
        Self::InvalidFormat(error.to_string())
    }
}

impl From<litchi_iwa_common::chart::number_format::Error> for crate::Error {
    fn from(error: litchi_iwa_common::chart::number_format::Error) -> Self {
        Self::InvalidFormat(error.to_string())
    }
}

impl From<litchi_iwa_common::chart::reference_line::Error> for crate::Error {
    fn from(error: litchi_iwa_common::chart::reference_line::Error) -> Self {
        Self::InvalidFormat(error.to_string())
    }
}

/// Locate one chart-private object and reject ambiguous cross-component IDs.
pub(crate) fn unique_chart_object_archive_name(
    package: &IWorkPackage,
    identifier: u64,
    object_label: &str,
) -> Result<String> {
    let mut archive_name = None;
    for name in package.iwa_entry_names() {
        if package.archive(name)?.object(identifier).is_none() {
            continue;
        }
        if archive_name.replace(name.to_owned()).is_some() {
            return Err(Error::Archive(format!(
                "{object_label} {identifier} occurs in multiple IWA components"
            )));
        }
    }
    archive_name
        .ok_or_else(|| Error::InvalidFormat(format!("{object_label} {identifier} is missing")))
}
