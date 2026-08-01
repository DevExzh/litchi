/// Constants for OLE file format
pub mod consts;

/// Shared MS-OFFCRYPTO structures for encryption and rights management.
pub mod office_crypto;

/// Shared bounded MS-OVBA project, module, and compression APIs.
pub mod ovba {
    pub use litchi_cfb::ovba::*;
}

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

/// Shared OfficeArt (Escher) functionality for Office binary formats
///
/// Escher is Microsoft's drawing layer format used across Office applications
/// (DOC, XLS, PPT) for shapes, connectors, and graphical elements.
/// This module provides shared zero-copy parsing and writing utilities.
pub mod escher;

/// Safe typed discovery and inert CFB replacement for legacy embedded objects.
pub mod embedded_object;

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

/// Trust-neutral verification and authoring of legacy binary Office signatures.
pub mod signature;

/// Legacy Excel spreadsheet (.xls) reader
///
/// This module provides functionality to parse Microsoft Excel spreadsheets
/// in the legacy binary format (.xls files), which are OLE2-based files.
pub mod xls;

pub use embedded_object::{
    CompObjMetadata, DocObjectDescriptor, LegacyOfficeEmbeddedObject, LegacyOfficeObjectCollection,
    LegacyOfficeObjectEditor, LegacyOfficeObjectFormat, LegacyOfficeObjectKind,
    LegacyOfficeObjectLimits, LegacyOfficePreview, LegacyOfficePreviewKind, OleNativePackage,
    discover_legacy_office_objects,
};

pub use xls::{XlsError, XlsWorkbook};

/// Image extraction bridge between OLE Escher records and `litchi-imgconv`.
///
/// Lives here rather than in the umbrella because it depends on private
/// `crate::escher` and `crate::ppt::escher` types that cannot cross crate
/// boundaries.
#[cfg(feature = "imgconv")]
pub mod extractor;
