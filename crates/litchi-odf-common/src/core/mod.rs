//! Core ODF parsing functionality.
//!
//! This module provides the fundamental building blocks for parsing
//! `OpenDocument` files. It handles ZIP package management, XML parsing,
//! manifest processing, and metadata extraction.
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
//! ## ✅ Metadata (`metadata/`) - COMPLETE
//! - ✅ Unified metadata structure
//! - ✅ Dublin Core fields (title, creator, description, etc.)
//! - ✅ ODF-specific fields (editing cycles, generator, etc.)
//! - ✅ Creation and modification timestamps
//!
//! # References
//! - ODF Specification: §2 (Documents), §3 (Metadata)
//! - ODF Toolkit: ODFDOM package classes
//! - ZIP format: PKZIP Application Note

/// ODF package-entry decryption.
mod encryption;
/// Shared ownership for simple packaged ODF families.
pub mod family;
/// ODF manifest parsing
mod manifest;
/// ODF metadata parsing
pub mod metadata;
/// ODF package handling
pub mod package;
/// ODF package writing
pub mod writer;
/// ODF XML utilities
pub mod xml;

// Re-export main types for convenience
// Manifest is internal to the package system
pub use encryption::{Cipher, Kdf, Profile, StartKey};
pub use family::{Package, validate_content_part};
#[allow(
    unused_imports,
    reason = "The manifest descriptors are intentionally re-exported as the core public API."
)]
pub use manifest::{
    Manifest, ManifestChecksum, ManifestChecksumAlgorithm, ManifestEncryption,
    ManifestEncryptionAlgorithm, ManifestEntry, ManifestKeyDerivation, ManifestStartKeyGeneration,
};
pub use metadata::{
    AutoReloadMetadata, DocumentStatistics, HyperlinkBehaviourMetadata, Metadata, TemplateMetadata,
    UserDefinedMetadata, UserDefinedValueType,
};
pub use metadata::{MetaXmlPatch, patch_meta_xml};
pub use package::OwnedPackage;
pub use writer::{PackageWriter, Structure};
pub use xml::{Content, Meta, Styles};
