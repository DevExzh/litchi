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
mod data;
mod direction;
mod kind;
pub mod metadata_extractor;
pub(crate) mod options;
pub(crate) mod source;

pub use archive::IWorkChartArchive;
pub use axis::ChartAxis;
pub use axis_bounds::{ChartAxisBound, ChartValueAxisBounds};
pub use axis_scale::ChartValueAxisScale;
pub use axis_steps::{ChartAxisMajorStepCount, ChartAxisMinorStepCount, ChartValueAxisSteps};
pub use axis_style::ChartAxisTickMarkLocation;
pub use data::ChartData;
pub use direction::ChartSeriesDirection;
pub use kind::ChartKind;
pub use metadata_extractor::{ChartMetadata, ChartMetadataExtractor};

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
