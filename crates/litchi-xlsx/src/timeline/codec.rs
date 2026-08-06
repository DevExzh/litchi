//! Bounded SpreadsheetML timeline XML conversion.

use super::model::{CacheDefinition, Views};
use crate::error::Result;

/// Parse one `timelineCacheDefinition` part.
pub fn read_cache(xml: &[u8]) -> Result<CacheDefinition> {
    crate::timelines::parse_timeline_cache_definition(xml)
}

/// Serialize one `timelineCacheDefinition` part deterministically.
pub fn write_cache(value: &CacheDefinition) -> Result<Vec<u8>> {
    crate::timelines::write_timeline_cache_definition(value)
}

/// Parse one worksheet `timelines` part.
pub fn read_views(xml: &[u8]) -> Result<Views> {
    crate::timelines::parse_timelines(xml)
}

/// Serialize one worksheet `timelines` part deterministically.
pub fn write_views(value: &Views) -> Result<Vec<u8>> {
    crate::timelines::write_timelines(value)
}
