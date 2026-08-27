//! Typed, inert XLSB External Link data (MS-XLSB 2.1.7.25).
//!
//! Semantic values and bounded BIFF12 codecs are layered under this owner.
//! These APIs never open an external workbook,
//! contact DDE, instantiate OLE, refresh data, evaluate formulas, or execute
//! code.

use crate::raw::Error as WireError;
use thiserror::Error as ThisError;

mod codec;
mod limits;
mod model;
mod package;
mod transaction;
mod validation;

#[cfg(test)]
mod tests;

/// Result type for the standalone XLSB external-link codec.
pub type Result<T> = std::result::Result<T, Error>;

pub(crate) use codec::parse_external_link_with_budget;
pub(crate) use limits::Budget;
pub use limits::{
    ExternalLinkLimits, ExternalLinkLimitsBuilder, ExternalLinkLimitsError, ExternalLinkResource,
};

/// Error returned by the standalone XLSB external-link codec.
#[derive(Debug, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A validated BIFF12 wire operation failed.
    #[error(transparent)]
    Wire(#[from] WireError),
    /// A modeled external-link invariant was violated.
    #[error("invalid external link: {0}")]
    InvalidFormula(String),
    /// A fixed-width or length-prefixed structure has the wrong size.
    #[error("invalid length: expected {expected}, found {found}")]
    InvalidLength { expected: usize, found: usize },
    /// A bounded collection could not reserve its validated capacity.
    #[error("allocation failed for {resource}: {source}")]
    Allocation {
        resource: &'static str,
        #[source]
        source: std::collections::TryReserveError,
    },
    /// A governed external-link resource budget was exceeded.
    #[error("{resource} limit exceeded: actual {actual}, maximum {maximum}")]
    LimitExceeded {
        /// Resource whose configured ceiling was crossed.
        resource: ExternalLinkResource,
        /// Observed or requested amount.
        actual: usize,
        /// Configured maximum.
        maximum: usize,
    },
}

pub(crate) const MAX_COLLECTION_ITEMS: usize = 65_535;
pub(crate) const MAX_WIDE_STRING_UNITS: usize = 32_767;

/// Maximum row count accepted in an authored or parsed DDE/OLE cache.
pub const MAX_XLSB_EXTERNAL_CACHE_ROWS: u32 = 1_048_576;
/// Maximum column count accepted in an authored or parsed DDE/OLE cache.
pub const MAX_XLSB_EXTERNAL_CACHE_COLUMNS: u32 = 16_384;
/// Safety limit for materialized values across one DDE/OLE cache matrix.
pub const MAX_XLSB_EXTERNAL_CACHED_VALUES: usize = 1_048_576;

pub const EXTERNAL_REFERENCE_WORKBOOK: u16 = 0;
pub const EXTERNAL_REFERENCE_DDE: u16 = 1;
pub const EXTERNAL_REFERENCE_OLE: u16 = 2;
pub const EXTERNAL_NAME_BUILT_IN: u8 = 1;
pub const EXTERNAL_NAME_RESERVED_MASK: u8 = 0b0011_1110;
pub const DATA_ITEM_WANT_ADVISE: u8 = 1 << 1;
pub const DATA_ITEM_WANT_PICTURE: u8 = 1 << 2;
pub const DDE_ITEM_SUPPORTS_OLE: u8 = 1 << 3;
pub const DDE_ITEM_RESERVED_MASK: u8 = 0b0011_0001;
pub const OLE_ITEM_REQUIRED_CLASS_FLAG: u8 = 1 << 4;
pub const OLE_ITEM_DISPLAY_AS_ICON: u8 = 1 << 5;
pub const OLE_ITEM_RESERVED_MASK: u8 = 0b0000_1001;
pub const DATA_ITEM_REQUIRED_TRAILING_FLAG: u8 = 1;

/// Maximum bytes accepted or emitted for one external-link part stream.
pub const MAX_LINK_PART_BYTES: usize = 32 * 1024 * 1024;
/// Maximum number of opaque BIFF12 records retained by one source snapshot.
pub const MAX_UNKNOWN_RECORDS: usize = 65_535;
/// Maximum aggregate bytes retained for opaque BIFF12 records.
pub const MAX_UNKNOWN_BYTES: usize = 8 * 1024 * 1024;
pub use codec::{
    parse_external_link, parse_external_link_model, parse_external_link_model_with_limits,
    parse_external_link_with_limits, parse_external_link_with_relationship,
    parse_external_link_with_relationship_with_limits,
};
pub use model::{
    AreaReference, CachedValue, CellLocation, CellReference, DdeItem, DefinedName, Entries,
    ErrorValue, Kind, Link, NameFormula, NameFormulaKind, OleItem, Parsed, SheetRange,
    UnknownRecord, ValueMatrix, validate_number,
};
pub use package::{write_external_link_stream, write_external_link_stream_with_limits};
pub use transaction::{
    Commit, Patch, Snapshot, Transaction, apply, apply_with_limits, read, read_with_limits,
};
pub use validation::validate_link;
