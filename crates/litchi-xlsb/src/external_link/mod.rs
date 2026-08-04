//! Typed, inert XLSB External Link data (MS-XLSB 2.1.7.25).
//!
//! Semantic values, bounded BIFF12 codecs, and compatibility exports are
//! layered under this owner. These APIs never open an external workbook,
//! contact DDE, instantiate OLE, refresh data, evaluate formulas, or execute
//! code.

use crate::raw::Error as WireError;
use thiserror::Error as ThisError;

mod codec;
mod model;
mod package;

/// Result type for the standalone XLSB external-link codec.
pub type Result<T> = std::result::Result<T, Error>;

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
/// Historical constant spelling retained for source compatibility.
pub const MAX_XLSB_EXTERNAL_LINK_PART_BYTES: usize = MAX_LINK_PART_BYTES;

pub use codec::{
    parse_external_link, parse_external_link_model, parse_external_link_with_relationship,
};
pub use model::{
    AreaReference, CachedValue, CellLocation, CellReference, DdeItem, DefinedName, Entries,
    ErrorValue, Kind, Link, NameFormula, NameFormulaKind, OleItem, Parsed, SheetRange, ValueMatrix,
    validate_number,
};
pub use package::write_external_link_stream;

// Compatibility aliases for the former prefixed public owner vocabulary.
pub use model::AreaReference as XlsbExternalAreaReference;
pub use model::CachedValue as XlsbExternalCachedValue;
pub use model::CellLocation as XlsbExternalCellLocation;
pub use model::CellReference as XlsbExternalCellReference;
pub use model::DdeItem as XlsbDdeItem;
pub use model::DefinedName as XlsbExternalDefinedName;
pub use model::Entries as XlsbExternalEntries;
pub use model::ErrorValue as XlsbExternalErrorValue;
pub use model::Kind as XlsbExternalLinkKind;
pub use model::Link as XlsbExternalLink;
pub use model::NameFormula as XlsbExternalNameFormula;
pub use model::NameFormulaKind as XlsbExternalNameFormulaKind;
pub use model::OleItem as XlsbOleItem;
pub use model::Parsed as ParsedExternalLink;
pub use model::SheetRange as XlsbExternalSheetRange;
pub use model::ValueMatrix as XlsbExternalValueMatrix;
pub use model::validate_number as validate_external_number;
