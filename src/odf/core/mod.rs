//! Core ODF parsing functionality.
//!
//! This module provides the fundamental building blocks for parsing OpenDocument files.
//! It handles ZIP package management, XML parsing, manifest processing, and metadata extraction.
//!
//! # Implementation Progress
//!
//! ## ✅ Package System (`package.rs`) - COMPLETE
//! - ✅ `Package<R>` - Generic ZIP archive reader
//! - ✅ `from_reader()` - Open from any Read + Seek source
//! - ✅ `from_zip_archive()` - Reuse already-parsed ZIP archives
//! - ✅ `mimetype()` - Get MIME type from mimetype file
//! - ✅ `get_file()` - Extract file by path
//! - ✅ `has_file()` - Check file existence
//! - ✅ `files()` - List all files in package
//! - ✅ RefCell-based interior mutability for safe archive access
//!
//! ## ✅ Manifest (`manifest.rs`) - COMPLETE
//! - ✅ `Manifest` parsing from META-INF/manifest.xml
//! - ✅ File entry tracking with media types
//! - ✅ Encryption info parsing (basic)
//! - ✅ Manifest validation
//!
//! ## ✅ XML Processing (`xml.rs`) - COMPLETE
//! - ✅ `Content` - Parse content.xml (main document content)
//! - ✅ `Styles` - Parse styles.xml (document-wide styles)
//! - ✅ `Meta` - Parse meta.xml (metadata)
//! - ✅ `from_bytes()` - Parse from byte buffers
//! - ✅ High-performance quick-xml based parsing
//! - ✅ Namespace-aware processing
//! - ✅ Error handling with detailed messages
//!
//! ## ✅ Package Writing (`writer.rs`) - COMPLETE
//! - ✅ `PackageWriter<W>` - Generic ZIP archive writer
//! - ✅ `new()` / `with_writer()` - Create writers
//! - ✅ `set_mimetype()` - Set MIME type (stored uncompressed)
//! - ✅ `add_file()` - Add files to package
//! - ✅ `add_file_with_media_type()` - Add with manifest entry
//! - ✅ `finish()` / `finish_to_bytes()` - Finalize package
//! - ✅ Default template generation (content.xml, styles.xml, meta.xml, settings.xml)
//! - ✅ Manifest.xml auto-generation
//!
//! ## ✅ Metadata (`metadata.rs`) - COMPLETE
//! - ✅ Unified metadata structure
//! - ✅ Dublin Core fields (title, creator, description, etc.)
//! - ✅ ODF-specific fields (editing cycles, generator, etc.)
//! - ✅ Creation and modification timestamps
//!
//! # References
//! - ODF Specification: §2 (Documents), §3 (Metadata)
//! - ODF Toolkit: ODFDOM package classes
//! - ZIP format: PKZIP Application Note

/// ODF manifest parsing
mod manifest;
/// ODF metadata parsing
mod metadata;
/// ODF package handling
mod package;
/// ODF package writing
mod writer;
/// ODF XML utilities
mod xml;

// Re-export main types for convenience
// Manifest is internal to the package system
#[allow(unused_imports)]
pub use manifest::Manifest;
pub use package::Package;
pub use writer::{OdfStructure, PackageWriter};
pub use xml::{Content, Meta, Styles};
