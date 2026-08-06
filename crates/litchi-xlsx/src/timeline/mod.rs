//! Typed, bounded SpreadsheetML timeline ownership.
//!
//! Timelines are inert package controls. This owner preserves their XML and
//! graph state but never refreshes a PivotCache, recalculates formulas, or
//! renders/applies a host UI.

pub mod codec;
pub mod model;
pub mod package;
pub mod transaction;
pub mod validation;

pub use model::{
    Cache, CacheDefinition, CachePivotTable, FilterType, Level, OpaqueXml, Part, PivotFilter,
    Range, State, View, Views,
};
pub use package::{load_caches, load_parts, store_caches, store_part};
pub use transaction::Transaction;

pub const CACHE_CONTENT_TYPE: &str = crate::timelines::TIMELINE_CACHE_CONTENT_TYPE;
pub const CACHE_RELATIONSHIP_TYPE: &str = crate::timelines::TIMELINE_CACHE_RELATIONSHIP_TYPE;
pub const CONTENT_TYPE: &str = crate::timelines::TIMELINES_CONTENT_TYPE;
pub const RELATIONSHIP_TYPE: &str = crate::timelines::TIMELINES_RELATIONSHIP_TYPE;

/// Read one timeline cache definition from its XML part.
pub fn read_cache(xml: &[u8]) -> crate::error::Result<CacheDefinition> {
    codec::read_cache(xml)
}

/// Read one worksheet timelines part from its XML.
pub fn read_views(xml: &[u8]) -> crate::error::Result<Views> {
    codec::read_views(xml)
}

/// Validate one timeline cache definition without mutating it.
pub fn validate(value: &CacheDefinition) -> crate::error::Result<()> {
    validation::cache(value)
}

/// Explicitly refuses runtime/UI behavior that this inert owner does not
/// implement.
pub fn unsupported_ui() -> crate::error::Result<()> {
    Err(crate::error::Error::Unsupported {
        feature: "timeline UI, refresh, and filter application",
    })
}
