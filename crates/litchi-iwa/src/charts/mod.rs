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
pub(crate) mod arrangement;
pub(crate) mod axis;
pub(crate) mod axis_bounds;
pub(crate) mod axis_gridline_stroke;
pub(crate) mod axis_label_affixes;
pub(crate) mod axis_label_angle;
pub(crate) mod axis_number_format;
pub(crate) mod axis_scale;
pub(crate) mod axis_steps;
pub(crate) mod axis_style;
pub(crate) mod background_fill;
pub(crate) mod border;
pub(crate) mod border_stroke;
pub(crate) mod category_labels;
mod data;
mod direction;
pub(crate) mod donut_inner_radius;
pub(crate) mod font;
pub(crate) mod gaps;
pub(crate) mod hidden_data;
mod kind;
pub(crate) mod legend_fill;
pub(crate) mod legend_style;
pub mod metadata_extractor;
pub(crate) mod non_style;
pub(crate) mod number_format;
pub(crate) mod options;
pub(crate) mod pie_label_distance;
pub(crate) mod pie_labels;
pub(crate) mod pie_start_angle;
pub(crate) mod pie_wedge_explosion;
pub(crate) mod reference_lines;
pub(crate) mod rounded_corners;
pub(crate) mod series_connection_line;
pub(crate) mod series_error_bar_auto_fit;
pub(crate) mod series_error_bars;
pub(crate) mod series_fill;
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

pub use archive::IWorkChartArchive;
pub use arrangement::ChartArrangement;
pub use axis::ChartAxis;
pub use axis_bounds::{ChartAxisBound, ChartValueAxisBounds};
pub use axis_gridline_stroke::{ChartAxisGridline, ChartAxisGridlineStroke};
pub use axis_label_angle::ChartAxisLabelAngle;
pub use axis_scale::ChartValueAxisScale;
pub use axis_steps::{ChartAxisMajorStepCount, ChartAxisMinorStepCount, ChartValueAxisSteps};
pub use axis_style::ChartAxisTickMarkLocation;
pub use category_labels::{
    ChartCategoryLabelFrequency, ChartCategoryLabelInterval, ChartCategoryLabelLayout,
};
pub use data::ChartData;
pub use direction::ChartSeriesDirection;
pub use donut_inner_radius::ChartDonutInnerRadius;
pub use font::{ChartFont, ChartFontSize};
pub use gaps::{ChartGapPercentage, ChartGapSpacing};
pub use kind::ChartKind;
pub use legend_fill::ChartLegendFill;
pub use metadata_extractor::{ChartMetadata, ChartMetadataExtractor};
pub use number_format::{
    ChartDecimalPlaces, ChartFixedDecimalPlaces, ChartLabelAffixes, ChartNegativeStyle,
    ChartNumberFormat,
};
pub use pie_label_distance::ChartPieLabelDistance;
pub use pie_labels::ChartPieLabelVisibility;
pub use pie_start_angle::ChartPieStartAngle;
pub use pie_wedge_explosion::{ChartPieWedgeExplosion, ChartPieWedgeIndex};
pub use reference_lines::{ChartReferenceLine, ChartReferenceLineKind, ChartReferenceLineValue};
pub use rounded_corners::{ChartCornerRadius, ChartRoundedCorners};
pub use series_connection_line::{ChartSeriesConnectionLine, ChartSeriesConnectionLineKind};
pub use series_error_bar_auto_fit::ChartSeriesErrorBarAutoFit;
pub use series_error_bars::{
    ChartErrorBarCustomValue, ChartErrorBarCustomValues, ChartErrorBarDirection,
    ChartErrorBarFixedValue, ChartErrorBarPercentage, ChartErrorBarStandardDeviationCount,
    ChartSeriesErrorBars,
};
pub use series_fill::ChartSeriesFillKind;
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
pub use series_value_label_affixes::ChartSeriesValueLabelAffixes;
pub use series_value_label_auto_fit::ChartSeriesValueLabelAutoFit;
pub use series_value_label_location::ChartSeriesValueLabelLocation;
pub use series_value_label_number_format::{
    ChartSeriesValueLabelDecimalPlaces, ChartSeriesValueLabelFixedDecimalPlaces,
    ChartSeriesValueLabelNegativeStyle, ChartSeriesValueLabelNumberFormat,
};
pub use series_value_labels::{ChartSeriesIndex, ChartSeriesValueLabelVisibility};
pub use shadow::ChartShadow;

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
