//! Layered XLSB timeline cache and worksheet timeline owner.
//!
//! Timeline caches/views use the XML payloads specified by `[MS-XLSB]`, while
//! workbook and worksheet relationship references remain BIFF12 records.
//! This owner is metadata-only: filtering, refresh, and `PivotCache` execution
//! are deliberately outside the bounded transaction slice.

mod codec;
mod model;
pub mod package;
pub(crate) mod transaction;
mod validation;

#[cfg(test)]
mod tests;

pub use codec::{parse_cache, parse_views, write_cache, write_views};
pub use model::{Cache, Filter, FilterType, Level, PivotTable, Range, State, View, Views};
pub use package::{CachePart, ViewPart};
pub use transaction::{Commit, Content, Patch, Snapshot, Transaction};
