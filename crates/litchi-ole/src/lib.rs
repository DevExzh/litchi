/// Constants for OLE file format
pub mod consts;

// CFB substrate types re-exported so callers can reach them through the
// `litchi::ole` namespace as well as `litchi_cfb` directly.
pub use litchi_cfb::{
    DOCUMENT_SUMMARY_INFORMATION_FMTID, DirectoryEntry, OleError, OleFile, OleMetadata,
    OlePropertySetEditor, OleWriter, PropertySet, PropertySetGuid, PropertySetStream,
    PropertyValue, SUMMARY_INFORMATION_FMTID, StandardPropertySet, USER_DEFINED_PROPERTIES_FMTID,
    is_ole_file,
};

pub use litchi_cfb::writer;

/// MTEF extractor for OLE documents (internal use only)
#[cfg(feature = "formula")]
mod mtef_extractor;

/// Property List with Character Positions (PLCF) parser.
///
/// PLCF is a data structure used extensively in legacy Office binary formats
/// to map character positions to properties or data.
pub mod plcf;

/// Shared SPRM (Single Property Modifier) parsing
///
/// SPRM parsing logic shared between DOC and PPT formats.
/// Based on Apache POI's SPRM handling.
pub mod sprm;

/// SPRM operation constants and utilities
///
/// Complete SPRM operation definitions based on Apache POI.
pub mod sprm_operations;

// Migration-only OfficeArt writers that have not yet moved into `litchi-odraw`.
// The old numeric Escher facade is intentionally not part of the public API.
#[allow(dead_code)]
mod escher;

/// Legacy Word document (.doc) reader
///
/// This module provides functionality to parse Microsoft Word documents
/// in the legacy binary format (.doc files), which are OLE2-based files.
pub mod doc;

/// Legacy PowerPoint presentation (.ppt) reader
///
/// This module provides functionality to parse Microsoft PowerPoint presentations
/// in the legacy binary format (.ppt files), which are OLE2-based files.
pub mod ppt;

/// Legacy Excel spreadsheet (.xls) reader
///
/// This module provides functionality to parse Microsoft Excel spreadsheets
/// in the legacy binary format (.xls files), which are OLE2-based files.
pub mod xls;

pub use xls::{XlsError, XlsWorkbook};

/// Image extraction bridge between typed OfficeArt records and optional codecs.
pub mod extractor;
