//! Layered SpreadsheetML timeline cache and worksheet timeline owners.
//!
//! Semantic values live in [`model`], bounded XML conversion in [`codec`],
//! and OPC relationship/part ownership in [`package`]. The historical
//! `litchi_xlsx::timelines` path remains the public facade.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

use crate::error::{Error, Result};

pub(super) const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
pub(super) const STRICT_SML: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
pub(super) const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
pub(super) const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
pub(super) const X15: &str = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main";
pub(super) const XR10: &str = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10";

pub const TIMELINE_CACHE_CONTENT_TYPE: &str = "application/vnd.ms-excel.TimelineCache+xml";
pub const TIMELINE_CACHE_RELATIONSHIP_TYPE: &str =
    "http://schemas.microsoft.com/office/2010/relationships/TimelineCache";
pub const TIMELINES_CONTENT_TYPE: &str = "application/vnd.ms-excel.Timeline+xml";
pub const TIMELINES_RELATIONSHIP_TYPE: &str =
    "http://schemas.microsoft.com/office/2010/relationships/Timeline";
pub const TIMELINE_CACHE_EXTENSION_URI: &str = "{D0CA8CA8-9F24-4464-BF8E-62219DCF47F9}";
pub const TIMELINES_EXTENSION_URI: &str = "{7E03D99C-DC04-49d9-9315-930204A7B6E9}";

pub(super) const MAX_XML_BYTES: usize = 16 * 1024 * 1024;
pub(super) const MAX_REWRITE_BYTES: usize = 32 * 1024 * 1024;
pub(super) const MAX_OPAQUE_BYTES: usize = 4 * 1024 * 1024;
pub(super) const MAX_TOTAL_OPAQUE_BYTES: usize = 16 * 1024 * 1024;
pub(super) const MAX_STRING_BYTES: usize = 1024 * 1024;
pub(super) const MAX_NODES: usize = 250_000;
pub(super) const MAX_DEPTH: usize = 128;
pub(super) const MAX_CACHES: usize = 4096;
pub(super) const MAX_TIMELINES: usize = 16_384;
pub(super) const MAX_PIVOT_TABLES: usize = 65_536;

pub(super) fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

pub(super) fn limit(name: &str) -> Error {
    invalid(format!("Timeline {name} limit exceeded"))
}

pub(super) fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(litchi_ooxml_common::XmlError::Malformed(error.to_string()))
}

pub(super) fn bounded(value: &str, name: &str) -> Result<()> {
    if value.len() <= MAX_STRING_BYTES {
        Ok(())
    } else {
        Err(limit(name))
    }
}

pub(super) fn bounded_nonempty(value: &str, name: &str) -> Result<()> {
    if value.is_empty() {
        return Err(invalid(format!("{name} cannot be empty")));
    }
    bounded(value, name)
}

pub use codec::{
    parse_timeline_cache_definition, parse_timelines, write_timeline_cache_definition,
    write_timelines,
};
pub use model::{
    Cache, CacheDefinition, CachePivotTable, FilterType, Level, OpaqueXml, PivotFilter, Range,
    State, View, Views, WorksheetView,
};
pub use package::{
    load_timeline_caches, load_timelines, store_timeline_caches, store_worksheet_timelines,
};

// Compatibility aliases preserve the pre-layering public surface. Canonical
// names above are contextual to this owner and do not repeat its prefix.
pub type Timeline = View;
pub type TimelineCacheDefinition = CacheDefinition;
pub type TimelineCachePivotTable = CachePivotTable;
pub type TimelineLevel = Level;
pub type TimelineOpaqueXml = OpaqueXml;
pub type TimelinePivotFilter = PivotFilter;
pub type TimelineRange = Range;
pub type TimelineState = State;
pub type Timelines = Views;
pub type WorkbookTimelineCache = Cache;
pub type WorksheetTimelines = WorksheetView;
pub type PivotFilterType = FilterType;
