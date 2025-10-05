/// Constants for OLE file format
pub mod consts;

/// Main OLE file parsing implementation
mod file;

/// Metadata extraction from OLE property streams
mod metadata;

/// Legacy Word document (.doc) reader
///
/// This module provides functionality to parse Microsoft Word documents
/// in the legacy binary format (.doc files), which are OLE2-based files.
pub mod doc;

// Re-export public types for convenient access
pub use file::{is_ole_file, DirectoryEntry, OleError, OleFile};
pub use metadata::{OleMetadata, PropertyValue};
