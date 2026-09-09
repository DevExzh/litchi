//! Inert XLSB Calculation Chain ownership ([MS-XLSB] 2.1.7.4).
//!
//! The local XLSB specification defines the Calculation Chain as an optional
//! Workbook-owned performance cache, but does not provide its binary record
//! grammar. This owner therefore validates only the OPC graph and bounded
//! opaque bytes. Generic BIFF12 framing is available only as an explicit
//! diagnostic and is not an admission requirement. The payload remains
//! opaque and is never evaluated, recalculated, refreshed, or rewritten.

mod model;
mod package;
mod patch;
mod snapshot;
mod transaction;
mod workbook;

#[cfg(test)]
mod tests;

pub use model::{
    ChainPart, HARD_MAX_GRAPH_PARTS, HARD_MAX_PART_BYTES, HARD_MAX_RECORDS, HARD_MAX_RELATIONSHIPS,
    ReadLimits,
};
pub use package::{CONTENT_TYPE, RELATIONSHIP_TYPE};
pub use patch::{Commit, Patch};
pub use snapshot::Snapshot;
pub use transaction::Transaction;

/// Alias matching the concise limits naming used by other XLSB feature owners.
pub type Limits = ReadLimits;

pub use crate::package::error::{Error, Result};
