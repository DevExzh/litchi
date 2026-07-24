//! PivotCache definition stream parsing (MS-XLSB 2.1.7.38).
//!
//! Parses a `pivotCacheDefinition*.bin` part into a typed
//! [`PivotCacheDefinition`] model covering the cache source (worksheet range,
//! consolidation, external, or OLAP), cache fields with their shared items,
//! grouping, OLAP hierarchies, calculated items and members, and refresh
//! metadata.
//!
//! Everything parsed here is an inert data snapshot: external connection
//! identifiers, relationship identifiers, MDX expressions, and formula token
//! streams are stored verbatim and are never dereferenced, contacted,
//! refreshed, or evaluated.
//!
//! The parser tolerates unknown record types (they are skipped) and rejects
//! structurally malformed payloads for the records it understands.

mod model;
mod parse;

pub use model::*;
pub use parse::parse_pivot_cache_definition;

#[cfg(test)]
mod tests;
