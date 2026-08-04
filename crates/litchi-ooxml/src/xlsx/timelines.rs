//! Compatibility adapter for the canonical XLSX timeline codec cluster.
//!
//! Timeline cache/view models, bounded XML handling, and OPC graph validation
//! live in `litchi_xlsx::timelines`. This module retains the historical host
//! path and error variants for the OOXML facade.

use crate::error::{OoxmlError, Result};
use litchi_opc::{OpcPackage, PackURI};

pub use litchi_xlsx::timelines::{
    PivotFilterType, TIMELINE_CACHE_CONTENT_TYPE, TIMELINE_CACHE_EXTENSION_URI,
    TIMELINE_CACHE_RELATIONSHIP_TYPE, TIMELINES_CONTENT_TYPE, TIMELINES_EXTENSION_URI,
    TIMELINES_RELATIONSHIP_TYPE, Timeline, TimelineCacheDefinition, TimelineCachePivotTable,
    TimelineLevel, TimelineOpaqueXml, TimelinePivotFilter, TimelineRange, TimelineState, Timelines,
    WorkbookTimelineCache, WorksheetTimelines,
};

fn map_error(error: litchi_xlsx::Error) -> OoxmlError {
    match error {
        litchi_xlsx::Error::Package(error) => OoxmlError::Opc(error),
        litchi_xlsx::Error::MarkupCompatibility(error) => OoxmlError::from(error),
        litchi_xlsx::Error::Xml(error) => OoxmlError::Xml(error.to_string()),
        litchi_xlsx::Error::Common(error) => OoxmlError::Common(error),
        litchi_xlsx::Error::Invalid(message) => OoxmlError::InvalidFormat(message),
        other => OoxmlError::Xlsx(other),
    }
}

pub fn parse_timeline_cache_definition(xml: &[u8]) -> Result<TimelineCacheDefinition> {
    litchi_xlsx::timelines::parse_timeline_cache_definition(xml).map_err(map_error)
}

pub fn write_timeline_cache_definition(value: &TimelineCacheDefinition) -> Result<Vec<u8>> {
    litchi_xlsx::timelines::write_timeline_cache_definition(value).map_err(map_error)
}

pub fn parse_timelines(xml: &[u8]) -> Result<Timelines> {
    litchi_xlsx::timelines::parse_timelines(xml).map_err(map_error)
}

pub fn write_timelines(value: &Timelines) -> Result<Vec<u8>> {
    litchi_xlsx::timelines::write_timelines(value).map_err(map_error)
}

pub fn load_timeline_caches(
    package: &OpcPackage,
    workbook_name: &PackURI,
) -> Result<Vec<WorkbookTimelineCache>> {
    litchi_xlsx::timelines::load_timeline_caches(package, workbook_name).map_err(map_error)
}

pub fn store_timeline_caches(
    package: &mut OpcPackage,
    workbook_name: &PackURI,
    caches: &[WorkbookTimelineCache],
) -> Result<()> {
    litchi_xlsx::timelines::store_timeline_caches(package, workbook_name, caches).map_err(map_error)
}

pub fn load_timelines(
    package: &OpcPackage,
    workbook_name: &PackURI,
) -> Result<Vec<WorksheetTimelines>> {
    litchi_xlsx::timelines::load_timelines(package, workbook_name).map_err(map_error)
}

pub fn store_worksheet_timelines(
    package: &mut OpcPackage,
    workbook_name: &PackURI,
    value: &WorksheetTimelines,
) -> Result<()> {
    litchi_xlsx::timelines::store_worksheet_timelines(package, workbook_name, value)
        .map_err(map_error)
}
