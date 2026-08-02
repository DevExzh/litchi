//! Office Open XML (OOXML) format implementation.
//!
//! This module provides parsing and manipulation of Office Open XML documents,
//! including Word (.docx), Excel (.xlsx, .xlsb), and PowerPoint (.pptx) files.
//!
//! The implementation is based on the Open Packaging Conventions (OPC) and
//! follows the structure of the python-docx library, adapted for Rust with
//! performance optimizations.
//!
//! # Architecture
//!
//! The module is organized into several layers:
//!
//! 1. **OPC Layer** (`opc`): Low-level package handling (ZIP, parts, relationships)
//! 2. **Shared Utilities** (`shared`, `error`): Common types used across formats
//! 3. **Format-Specific Modules**:
//!    - `docx`: Word documents
//!    - `xlsx`: Excel spreadsheets
//!    - `xlsb`: Excel binary spreadsheets
//!    - `pptx`: PowerPoint presentations (placeholder)
//!    - `metadata`: Core properties/metadata extraction
//!
//! # Example: Working with Word Documents
//!
//! ```rust,no_run
//! use litchi_ooxml::docx::Package;
//!
//! // Open and read a document
//! let pkg = Package::open("document.docx")?;
//! let doc = pkg.document()?;
//!
//! // Extract text content
//! let text = doc.text()?;
//! println!("Document contains {} paragraphs", doc.paragraph_count()?);
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```
pub mod api;
pub mod docx;
pub mod error;
// OPC implementation lives in the focused litchi-opc crate.
pub mod opc {
    pub use litchi_opc::*;
}
pub mod pivot;
pub mod pptx;
mod vba_package;
pub mod xlsb;
pub mod xlsx;

/// Bounded, runtime-neutral OOXML package encryption.
///
/// The contextual facade keeps callers on concise names such as
/// `encryption::Mode` and `encryption::Limits` while the canonical
/// implementation remains owned by `litchi-crypto`.
#[cfg(feature = "encryption")]
pub use litchi_crypto::ooxml as encryption;

#[cfg(feature = "fonts")]
pub mod fonts;

pub use litchi_sign::{Coverage, Limits, Policy, Signer, Status, Trust};
pub use opc::{OpcPackage, PackURI};

// Re-export common utilities
pub use litchi_ooxml_common::{
    DocumentProperties, ExpandedName, MceCapabilities, MceError, MceLimits, MceOutput, MceReport,
    process_markup_compatibility, ribbon, web,
};

// Re-export error types
pub use error::{OoxmlError, Result};
