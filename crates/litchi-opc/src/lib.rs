//! Open Packaging Conventions (OPC) implementation.
//!
//! This module provides a complete implementation of the OPC specification,
//! which defines the structure and packaging format for Office Open XML documents.
//! It includes support for:
//!
//! - Package structure (parts, relationships)
//! - Content type management
//! - ZIP-based physical packaging
//! - Efficient parsing and minimal memory allocation
//!
//! # Performance Features
//!
//! - Uses `memchr` for fast string searching in XML
//! - Uses `atoi_simd` for fast integer parsing
//! - Uses `quick-xml` for efficient zero-copy XML parsing
//! - Minimizes allocations by borrowing data where possible
//! - Uses hash maps for O(1) lookups
//!
//! # Bounded ingestion
//!
//! Every ordinary package constructor uses [`ReadLimits::default`]. Build a
//! checked, tighter profile with [`ReadLimits::builder`] and pass it to an
//! `*_with_limits` constructor. The policy bounds package input; ZIP members,
//! names, metadata, and declared sizes; materialized OPC parts;
//! `[Content_Types].xml`; and relationship XML, attributes, targets, events,
//! depth, and graph traversal. It is a Litchi safety policy, not a set of
//! specification maxima, and provides a defensive-consumer boundary for
//! ECMA-376 Part 2 sections 7.3.6 and 10 and MS-OI29500 sections 2.1.1749-1752.
//!
//! Macros, VBA, ActiveX, controls, OLE objects, and embedded code are retained
//! only as inert blobs when exposed or preserved. This crate never executes or
//! activates them.

pub mod atomic;

pub mod constants;
pub mod content_type;
pub mod error;
pub mod limits;
pub mod members;
pub mod package;
pub mod packuri;
pub mod part;
pub mod phys_pkg;
pub mod pkgreader;
pub mod pkgwriter;
pub mod rel;
#[cfg(feature = "sign")]
pub mod sign;

// Re-export commonly used types
pub use content_type::ContentType;
pub use error::{OpcError, Result};
pub use limits::{ReadLimits, ReadLimitsBuilder, ReadResource};
pub use members::{NonPartMember, NonPartReason};
pub use package::{FontEmbedding, OpcPackage, SaveOptions};
pub use packuri::PackURI;
pub use part::{BlobPart, Part, XmlPart};
pub use pkgwriter::PackageWriter;
pub use rel::{Relationship, Relationships, TargetMode};
