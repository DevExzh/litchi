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
pub(crate) mod axis;
pub(crate) mod axis_bounds;
pub(crate) mod axis_scale;
pub(crate) mod axis_steps;
pub(crate) mod axis_style;
pub(crate) mod background_fill;
pub(crate) mod border;
pub(crate) mod border_stroke;
mod data;
mod direction;
pub(crate) mod donut_inner_radius;
pub(crate) mod gaps;
pub(crate) mod hidden_data;
mod kind;
pub mod metadata_extractor;
pub(crate) mod non_style;
pub(crate) mod options;
pub(crate) mod pie_label_distance;
pub(crate) mod pie_labels;
pub(crate) mod pie_start_angle;
pub(crate) mod pie_wedge_explosion;
pub(crate) mod rounded_corners;
pub(crate) mod series_non_style;
pub(crate) mod series_style;
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
pub use axis::ChartAxis;
pub use axis_bounds::{ChartAxisBound, ChartValueAxisBounds};
pub use axis_scale::ChartValueAxisScale;
pub use axis_steps::{ChartAxisMajorStepCount, ChartAxisMinorStepCount, ChartValueAxisSteps};
pub use axis_style::ChartAxisTickMarkLocation;
pub use data::ChartData;
pub use direction::ChartSeriesDirection;
pub use donut_inner_radius::ChartDonutInnerRadius;
pub use gaps::{ChartGapPercentage, ChartGapSpacing};
pub use kind::ChartKind;
pub use metadata_extractor::{ChartMetadata, ChartMetadataExtractor};
pub use pie_label_distance::ChartPieLabelDistance;
pub use pie_labels::ChartPieLabelVisibility;
pub use pie_start_angle::ChartPieStartAngle;
pub use pie_wedge_explosion::{ChartPieWedgeExplosion, ChartPieWedgeIndex};
pub use rounded_corners::{ChartCornerRadius, ChartRoundedCorners};
pub use series_trendline::ChartSeriesTrendline;
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
