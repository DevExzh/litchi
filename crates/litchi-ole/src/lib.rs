/// Constants for OLE file format
pub mod consts;

// CFB substrate — moved to the litchi-cfb crate. Re-exports preserved for
// backward compatibility through the umbrella `litchi::ole` namespace.
pub use litchi_cfb::{
    DirectoryEntry, OleError, OleFile, OleMetadata, OleWriter, PropertyValue, is_ole_file,
};

// CFB writer module retained as a re-export so `crate::writer::*` paths
// inside the format parsers keep resolving until P4c moves litchi-ole out.
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

/// Shared OfficeArt (Escher) functionality for Office binary formats
///
/// Escher is Microsoft's drawing layer format used across Office applications
/// (DOC, XLS, PPT) for shapes, connectors, and graphical elements.
/// This module provides shared zero-copy parsing and writing utilities.
pub mod escher;

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

/// Image extraction bridge between OLE Escher records and `litchi-imgconv`.
///
/// Relocated into this crate from the umbrella `litchi::images` module
/// during P4c (workspace-split) because the extractor depends on private
/// `crate::escher` and `crate::ppt::escher` types that cannot be referenced
/// across crate boundaries from the umbrella.
#[cfg(feature = "imgconv")]
pub mod extractor;
