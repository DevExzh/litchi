//! Layered, inert XLSX pivot-chart ownership.
//!
//! Semantic objects live in [`model`], bounded DrawingML conversion in
//! [`codec`], and workbook relationship-graph resolution in [`package`].
//! Pivot charts remain read-only and inert: this owner never refreshes a
//! pivot cache or renders a chart.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use codec::parse_binding;
pub use model::{
    Binding, Chart, DropZoneVisibility, FieldType, Options, Series, SheetCharts, SheetKind, Source,
};
pub use package::{load, load_sheet};

pub(crate) use codec::default_options_extension_xml;

#[cfg(test)]
pub(crate) use package::resolve_source_name;

use crate::error::Error;

/// Extension URI of the MS-XLSX `c14:pivotOptions` series extension.
pub const OPTIONS_EXTENSION_URI: &str = "{781A3756-C4B2-4CAC-9D66-4F8C8630D5DC}";

/// Default `c:fmtId` written for authored pivot charts (Excel writes 0).
pub const DEFAULT_FORMAT_ID: u32 = 0;

pub(super) const C14_CHART_NAMESPACE: &str =
    "http://schemas.microsoft.com/office/drawing/2007/8/2/chart";
pub(super) const CHARTSHEET_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";
pub(super) const STRICT_CHARTSHEET_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/chartsheet";
pub(super) const CHARTSHEET_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml";
pub(super) const MAX_CHART_PART_BYTES: usize = 32 * 1024 * 1024;
pub(super) const MAX_WORKBOOK_BYTES: usize = 32 * 1024 * 1024;
pub(super) const MAX_DRAWING_PART_BYTES: usize = 32 * 1024 * 1024;
pub(super) const MAX_DRAWINGS_PER_WORKSHEET: usize = 1024;
pub(super) const MAX_PIVOT_CHARTS_PER_WORKSHEET: usize = 4096;
pub(super) const MAX_SERIES_PER_CHART: usize = 16_384;
pub(super) const MAX_EXTENSION_URIS: usize = 256;
pub(super) const MAX_TEXT_BYTES: usize = 1024 * 1024;
pub(super) const MAX_DEPTH: usize = 256;

pub(super) fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

pub(super) fn limit(name: &str) -> Error {
    invalid(format!("pivot chart {name} limit exceeded"))
}
