//! Layered SpreadsheetML chartsheet package graph.
//!
//! This package boundary owns the inert OPC graph around one chartsheet:
//! drawings, classic and extended charts, companion parts, media, VML,
//! Printer Settings, and extension relationships. The semantic chartsheet
//! grammar remains in [`super`]. Its `model`, `codec`, and `operations`
//! children keep graph values, bounded leaf XML, and package mutation/query
//! orchestration separate.

mod codec;
mod model;
mod operations;

#[cfg(test)]
mod tests;

pub use model::{
    BackgroundImageContentType, BackgroundPicture, ChartCompanionResource,
    ChartEmbeddedPackageContentType, ChartEmbeddedPackageResource, ChartOutboundResource,
    ChartResource, ChartResourceKind, ChartThemeOverrideResource, ChartUserShapesResource,
    DrawingResource, Entry, ExtensionRelationship, ExtensionRelationshipTarget, ImageContentType,
    ImageResource, Package, PrinterSettings, VmlDrawingResource,
};
pub use operations::{load_chartsheet, store_chartsheet, validate_package};

use crate::error::Error;

pub(super) const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
pub(super) const STRICT_SML: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
#[cfg(test)]
pub(super) const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
#[cfg(test)]
pub(super) const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
#[cfg(test)]
pub(super) const XDR: &str = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
#[cfg(test)]
pub(super) const STRICT_XDR: &str = "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing";
#[cfg(test)]
pub(super) const CHART: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
#[cfg(test)]
pub(super) const STRICT_CHART: &str = "http://purl.oclc.org/ooxml/drawingml/chart";
pub(super) const DRAWING_MAIN: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
pub(super) const STRICT_DRAWING_MAIN: &str = "http://purl.oclc.org/ooxml/drawingml/main";
pub(super) const CHART_EX: &str = "http://schemas.microsoft.com/office/drawing/2014/chartex";
pub(super) const CHART_EX_CHOICE: &str =
    "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex";
#[cfg(test)]
pub(super) const CHART_STYLE: &str = "http://schemas.microsoft.com/office/drawing/2012/chartStyle";
pub(super) const CHARTSHEET_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";
pub(super) const STRICT_CHARTSHEET_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/chartsheet";
#[cfg(test)]
pub(super) const IMAGE_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
pub(super) const CHART_EX_REL: &str =
    "http://schemas.microsoft.com/office/2014/relationships/chartEx";
pub(super) const CHART_STYLE_REL: &str =
    "http://schemas.microsoft.com/office/2011/relationships/chartStyle";
pub(super) const CHART_COLOR_STYLE_REL: &str =
    "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle";
#[cfg(test)]
pub(super) const VML_DRAWING_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";
pub(super) const CHARTSHEET_CT: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml";
pub(super) const DRAWING_CT: &str = "application/vnd.openxmlformats-officedocument.drawing+xml";
pub(super) const CHART_CT: &str =
    "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
pub(super) const CHART_EX_CT: &str = "application/vnd.ms-office.chartex+xml";
pub(super) const CHART_STYLE_CT: &str = "application/vnd.ms-office.chartstyle+xml";
pub(super) const CHART_COLOR_STYLE_CT: &str = "application/vnd.ms-office.chartcolorstyle+xml";
pub(super) const CHART_USER_SHAPES_CT: &str =
    "application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml";
pub(super) const THEME_OVERRIDE_CT: &str =
    "application/vnd.openxmlformats-officedocument.themeOverride+xml";
pub(super) const VML_DRAWING_CT: &str = "application/vnd.openxmlformats-officedocument.vmlDrawing";
pub(super) const MAX_XML_BYTES: usize = 32 * 1024 * 1024;
pub(super) const MAX_DRAWING_BYTES: usize = 16 * 1024 * 1024;
pub(super) const MAX_CHART_BYTES: usize = 32 * 1024 * 1024;
pub(super) const MAX_BACKGROUND_IMAGE_BYTES: usize = 32 * 1024 * 1024;
pub(super) const MAX_VML_DRAWING_BYTES: usize = 16 * 1024 * 1024;
pub(super) const MAX_TOTAL_RESOURCE_BYTES: usize = 128 * 1024 * 1024;
pub(super) const MAX_NODES: usize = 500_000;
pub(super) const MAX_DEPTH: usize = 256;
pub(super) const MAX_NAMESPACE_BINDINGS: usize = 4096;
pub(super) const MAX_STRING_BYTES: usize = 4 * 1024 * 1024;
pub(super) const MAX_CHARTS: usize = 256;
pub(super) const MAX_CHART_EX_BYTES: usize = 16 * 1024 * 1024;
pub(super) const MAX_CHART_STYLE_BYTES: usize = 4 * 1024 * 1024;
pub(super) const MAX_CHART_COLOR_STYLE_BYTES: usize = 4 * 1024 * 1024;
pub(super) const MAX_CHART_STYLE_PARTS: usize = 16;
pub(super) const MAX_CHART_USER_SHAPES_BYTES: usize = 8 * 1024 * 1024;
pub(super) const MAX_CHART_USER_SHAPE_IMAGES: usize = 256;
pub(super) const MAX_CHART_USER_SHAPE_IMAGE_BYTES: usize = 32 * 1024 * 1024;
pub(super) const MAX_CHART_DIRECT_IMAGES: usize = 256;
// ChartEx relationships are limited by the bounded direct-image, companion,
// and single-resource families owned by this module.  The chartEx schema
// permits one externalData package relationship (MS-ODRAWXML 2.24).
pub(super) const MAX_CHART_RELATIONSHIPS: usize =
    MAX_CHART_DIRECT_IMAGES + (MAX_CHART_STYLE_PARTS * 2) + 3;
pub(super) const MAX_CHART_THEME_OVERRIDE_BYTES: usize = 4 * 1024 * 1024;
pub(super) const MAX_CHART_THEME_IMAGES: usize = 256;
pub(super) const MAX_CHART_EMBEDDED_PACKAGE_BYTES: usize = 64 * 1024 * 1024;
#[cfg(test)]
pub(super) const MAX_WEB_PUBLISH_ITEMS: usize = 4096;
#[cfg(test)]
pub(super) const MAX_WEB_PUBLISH_STRING_BYTES: usize = 64 * 1024;
#[cfg(test)]
pub(super) const MAX_EXTENSIONS: usize = 1024;
#[cfg(test)]
pub(super) const MAX_EXTENSION_URI_BYTES: usize = 1024;
pub(super) const MAX_EXTENSION_PAYLOAD_BYTES: usize = 1024 * 1024;
pub(super) const MAX_EXTENSION_RELATIONSHIPS: usize = 1024;
pub(super) const MAX_EXTENSION_RELATIONSHIP_STRING_BYTES: usize = 64 * 1024;

pub(super) fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

pub(super) fn limit(name: &str) -> Error {
    invalid(format!("chartsheet {name} limit exceeded"))
}
