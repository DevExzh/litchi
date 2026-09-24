//! Inert XLSB Volatile Dependencies support ([MS-XLSB] 2.1.7.60).
//!
//! The owner reads the Workbook implicit part lazily at the package boundary,
//! keeps cached RTD/cube values inert, and never refreshes, evaluates, or
//! contacts an external service. Unsupported source records remain safe for
//! exact no-ops and owner removal; typed rewrites refuse them.

mod codec;
mod model;
mod package;
mod patch;
mod snapshot;
mod transaction;
mod workbook;

#[cfg(test)]
mod tests;

#[cfg(test)]
mod codec_review_tests;

pub use model::{
    CachedValue, CellReference, Dependencies, DependencyKind, ErrorCode, HARD_MAX_ITEMS,
    HARD_MAX_PART_BYTES, HARD_MAX_RECORDS, HARD_MAX_RELATIONSHIPS, HARD_MAX_STRING_UNITS,
    HARD_MAX_TOTAL_STRING_UNITS, MainTopic, ReadLimits, Topic, VolatileType,
};
pub use package::{CONTENT_TYPE, DEFAULT_PART_NAME, RELATIONSHIP_TYPE};
pub use patch::{Commit, Patch};
pub use snapshot::Snapshot;
pub use transaction::Transaction;

/// Alias matching other XLSB feature owners.
pub type Limits = ReadLimits;

pub use crate::package::error::{Error, Result};
