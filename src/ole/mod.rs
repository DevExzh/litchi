/// Constants for OLE file format
pub mod consts;

/// Main OLE file parsing implementation
mod file;

/// Metadata extraction from OLE property streams
mod metadata;

// Re-export public types for convenient access
pub use file::{is_ole_file, DirectoryEntry, OleError, OleFile};
pub use metadata::{OleMetadata, PropertyValue};
