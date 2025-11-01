pub mod common;
pub mod docx;
pub mod error;
pub mod metadata;
/// Office Open XML (OOXML) format implementation.
///
/// This module provides parsing and manipulation of Office Open XML documents,
/// including Word (.docx), Excel (.xlsx, .xlsb), and PowerPoint (.pptx) files.
///
/// The implementation is based on the Open Packaging Conventions (OPC) and
/// follows the structure of the python-docx library, adapted for Rust with
/// performance optimizations.
///
/// # Architecture
///
/// The module is organized into several layers:
///
/// 1. **OPC Layer** (`opc`): Low-level package handling (ZIP, parts, relationships)
/// 2. **Shared Utilities** (`shared`, `error`): Common types used across formats
/// 3. **Format-Specific Modules**:
///    - `docx`: Word documents
///    - `xlsx`: Excel spreadsheets
///    - `xlsb`: Excel binary spreadsheets
///    - `pptx`: PowerPoint presentations (placeholder)
///    - `metadata`: Core properties/metadata extraction
///
/// # Example: Working with Word Documents
///
/// ```rust,no_run
/// use litchi::ooxml::docx::Package;
///
/// // Open and read a document
/// let pkg = Package::open("document.docx")?;
/// let doc = pkg.document()?;
///
/// // Extract text content
/// let text = doc.text()?;
/// println!("Document contains {} paragraphs", doc.paragraph_count()?);
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub mod opc;
pub mod pptx;
pub mod shared;
pub mod xlsb;
pub mod xlsx;

// Re-export commonly used types from OPC layer
pub use opc::{OpcPackage, PackURI};

// Re-export shared utilities
pub use shared::{Length, RGBColor};

// Re-export common utilities
pub use common::DocumentProperties;

// Re-export error types
pub use error::{OoxmlError, Result};
