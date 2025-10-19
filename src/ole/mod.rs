/// Constants for OLE file format
pub mod consts;

/// Main OLE file parsing implementation
mod file;

/// Metadata extraction from OLE property streams
mod metadata;

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

// Re-export public types for convenient access
pub use file::{is_ole_file, DirectoryEntry, OleError, OleFile};
pub use metadata::{OleMetadata, PropertyValue};
pub use xls::{XlsError, XlsWorkbook};
