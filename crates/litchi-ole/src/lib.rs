#![forbid(unsafe_code)]

//! Migration host for the legacy Word (`.doc`) implementation.
//!
//! Compound-file and OfficeArt primitives live in `litchi-cfb`,
//! `litchi-ole-common`, and `litchi-odraw`. Legacy PowerPoint support lives in
//! the independent `litchi-ppt` crate.

/// MTEF extractor for OLE documents (internal use only)
#[cfg(feature = "formula")]
mod mtef_extractor;

/// Shared SPRM (Single Property Modifier) parsing
///
/// SPRM parsing logic shared between DOC and PPT formats.
/// Based on Apache POI's SPRM handling.
pub mod sprm;

/// SPRM operation constants and utilities
///
/// Complete SPRM operation definitions based on Apache POI.
pub mod sprm_operations;

/// Legacy Word document (.doc) reader
///
/// This module provides functionality to parse Microsoft Word documents
/// in the legacy binary format (.doc files), which are OLE2-based files.
pub mod doc;
