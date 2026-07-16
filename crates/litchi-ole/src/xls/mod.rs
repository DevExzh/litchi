//! Legacy Excel (.xls) file format reader
//!
//! This module provides functionality to parse Microsoft Excel files
//! in the legacy binary format (.xls files), which are OLE2-based files.
//! The implementation is based on the BIFF (Binary Interchange File Format)
//! specification and draws inspiration from other spreadsheet libraries.

/// Error types for XLS parsing
mod error;

/// BIFF8 password-to-open encryption support
mod encryption;

/// BIFF record parsing utilities
pub mod records;

/// Workbook parsing implementation
mod workbook;

/// Worksheet parsing implementation
mod worksheet;

/// Cell value parsing and representation
mod cell;

/// BIFF formula token rendering
mod formula;

/// Shape extraction
pub mod shapes;

/// Shared parsing utilities
mod utils;

/// Merged cell range parsing (MERGECELLS 0x00E5)
pub mod merged_cells;

/// Hyperlink parsing (HLINK 0x01B8)
pub mod hyperlinks;

/// Comment/note parsing (NOTE 0x001C)
pub mod comments;

/// AutoFilter and sort parsing (AUTOFILTERINFO 0x009D, AUTOFILTER 0x009E, SORT 0x0090)
pub mod autofilter;

/// Pivot table parsing (SXVIEW, SXVD, SXVI, SXDI, SXVS, SXPI)
pub mod pivot_table;

/// Sheet protection parsing (PROTECT, OBJECTPROTECT, SCENPROTECT, PASSWORD)
pub mod protection;

/// XLS file writing
pub mod writer;

pub use cell::XlsCell;
pub use error::{XlsEncryptionKind, XlsError, XlsResult};
pub use records::{
    PhoneticAlignment, PhoneticRun, PhoneticString, PhoneticType, SharedStringFormatRun,
    SharedStringProperties,
};
pub use shapes::XlsShape;
pub use workbook::{XlsOpenOptions, XlsWorkbook};
pub use worksheet::XlsWorksheet;
pub use writer::XlsWriter;
