//! OLE2 file writing module
//!
//! This module provides comprehensive support for creating and modifying
//! OLE2 compound documents.
//!
//! # Implementation Notes
//!
//! The writer is based on Apache POI's POIFS implementation and follows
//! the Microsoft Compound File Binary Format specification.

/// FAT (File Allocation Table) generation
mod fat;

/// MiniFAT (Mini File Allocation Table) generation
mod minifat;

/// DIFAT (Double Indirect FAT) generation
mod difat;

/// Directory tree generation
mod directory;
#[cfg(test)]
pub(crate) use directory::DirectoryBuilder;

/// OLE2 header generation
mod header;

/// Core OLE writer implementation
mod core;

/// Integration tests for OLE writer
#[cfg(test)]
mod tests;

// Re-export public types
#[allow(
    clippy::module_name_repetitions,
    reason = "`OleWriter` is the established public API name; renaming it would break downstream crates"
)]
pub use core::OleWriter;
