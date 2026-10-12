//! Low-level access to the eager legacy PowerPoint reader.
//!
//! The canonical [`crate::Package`] owner keeps source identity and
//! publication policy separate from this generic `Read + Seek` projection.
//! Callers that need the native mutable reader should use this explicit
//! namespace.

/// The legacy generic PowerPoint reader.
pub use crate::package::model::Reader;
