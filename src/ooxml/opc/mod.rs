/// Open Packaging Conventions (OPC) implementation.
///
/// This module provides a complete implementation of the OPC specification,
/// which defines the structure and packaging format for Office Open XML documents.
/// It includes support for:
///
/// - Package structure (parts, relationships)
/// - Content type management
/// - ZIP-based physical packaging
/// - Efficient parsing and minimal memory allocation
///
/// # Performance Features
///
/// - Uses `memchr` for fast string searching in XML
/// - Uses `atoi_simd` for fast integer parsing
/// - Uses `quick-xml` for efficient zero-copy XML parsing
/// - Minimizes allocations by borrowing data where possible
/// - Uses hash maps for O(1) lookups

pub mod constants;
pub mod error;
pub mod package;
pub mod packuri;
pub mod part;
pub mod phys_pkg;
pub mod pkgreader;
pub mod rel;

// Re-export commonly used types
pub use package::OpcPackage;
pub use packuri::PackURI;
pub use part::{Part, XmlPart, BlobPart};
pub use rel::{Relationship, Relationships};

