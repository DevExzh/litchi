//! iWork Archive Format Support
//!
//! This module provides comprehensive support for parsing Apple's iWork file formats
//! (Pages, Keynote, Numbers) which use the IWA (iWork Archive) format.
//!
//! ## Quick Start
//!
//! ```rust,no_run
//! use litchi::iwa::Document;
//!
//! // Open an iWork document
//! let doc = Document::open("document.pages")?;
//!
//! // Extract text content
//! let text = doc.text()?;
//! println!("{}", text);
//!
//! // Get document statistics
//! let stats = doc.stats();
//! println!("Objects: {}", stats.total_objects);
//! println!("Application: {:?}", stats.application);
//!
//! // Extract structured data (tables, slides, sections)
//! let structured = doc.extract_structured_data()?;
//! println!("{}", structured.summary());
//! # Ok::<(), litchi::iwa::Error>(())
//! ```
//!
//! ## iWork File Structure
//!
//! iWork documents are bundles containing:
//! - `Index.zip`: Contains IWA files with serialized objects
//! - `Data/`: Directory containing media assets (images, videos, audio)
//! - `Metadata/`: Document metadata and properties
//! - Preview images at root level
//!
//! ## IWA Format
//!
//! Each `.iwa` file contains:
//! - Snappy-compressed data (custom framing without stream identifier)
//! - Protobuf-encoded messages
//! - Variable-length integers for message lengths
//! - ArchiveInfo and MessageInfo headers for metadata
//!
//! ## Features
//!
//! ### Text Extraction
//! - Automatic extraction from TSWP storage messages
//! - Support for all iWork applications
//! - Preserves document structure
//!
//! ### Media Management
//! - Automatic media asset discovery
//! - Support for images, videos, audio, PDFs
//! - Media extraction and statistics
//!
//! ### Structured Data
//! - Tables from Numbers (with CSV export)
//! - Slides from Keynote (with titles and content)
//! - Sections from Pages (with headings and paragraphs)
//!
//! ### Parsing from Bytes
//! - No file system access required
//! - Direct memory parsing
//! - Useful for web services and embedded systems
//!
//! ## Examples
//!
//! ### Parse from bytes
//!
//! ```rust,no_run
//! use litchi::iwa::Document;
//! use std::fs;
//!
//! let bytes = fs::read("document.pages")?;
//! let doc = Document::from_bytes(&bytes)?;
//! let text = doc.text()?;
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```
//!
//! ### Extract media
//!
//! ```rust,no_run
//! use litchi::iwa::Document;
//!
//! let doc = Document::open("presentation.key")?;
//!
//! // Get media statistics
//! if let Some(stats) = doc.media_stats() {
//!     println!("Media: {}", stats.summary());
//! }
//!
//! // Extract specific media file
//! if let Ok(data) = doc.extract_media("image.png") {
//!     std::fs::write("extracted.png", data)?;
//! }
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```
//!
//! ### Extract tables
//!
//! ```rust,no_run
//! use litchi::iwa::Document;
//!
//! let doc = Document::open("spreadsheet.numbers")?;
//! let structured = doc.extract_structured_data()?;
//!
//! for table in &structured.tables {
//!     let csv = table.to_csv();
//!     println!("Table: {}\n{}", table.name, csv);
//! }
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```
//!
//! ## Performance
//!
//! The implementation is optimized for:
//! - Fast decompression (50-100 MB/s per core)
//! - Efficient parsing (100-200 MB/s per core)
//! - Low memory overhead (~2-3x document size)
//! - O(1) message type lookups (perfect hash maps)
//!
//! ## Reference
//!
//! This implementation is based on:
//! - `libetonyek` - C++ library from Document Liberation Project
//! - `pyiwa` - Python iWork format reader
//! - `iWorkFileFormat` - Reverse-engineered format documentation

// Core parsing modules
pub mod snappy;
pub mod varint;
pub mod archive;
pub mod bundle;
pub mod registry;
pub mod object_index;
pub mod ref_graph;
pub mod protobuf;
pub mod media;
pub mod structured;

/// Shared text extraction utilities
pub mod text;

/// High-level iWork document types
pub mod document;

/// Application-specific modules
pub mod pages;
pub mod numbers;
pub mod keynote;

/// Re-export commonly used types
pub use archive::{ArchiveInfo, MessageInfo};
pub use bundle::{Bundle, BundleMetadata, PropertyValue};
pub use document::Document;
pub use snappy::SnappyStream;
pub use media::{MediaManager, MediaAsset, MediaType, MediaStats};
pub use structured::{Table, Slide, Section, StructuredData, CellValue};
pub use text::{TextExtractor, TextStorage, TextFragment, TextStyle, ParagraphStyle};
pub use ref_graph::ReferenceGraph;

/// Error types for iWork parsing
#[derive(Debug, thiserror::Error)]
pub enum Error {
    #[error("I/O error: {0}")]
    Io(#[from] std::io::Error),

    #[error("Invalid IWA format: {0}")]
    InvalidFormat(String),

    #[error("Snappy decompression error: {0}")]
    Snappy(String),

    #[error("Protobuf decoding error: {0}")]
    ProtobufDecode(#[from] prost::DecodeError),

    #[error("Unsupported message type: {0}")]
    UnsupportedMessageType(u32),

    #[error("Archive parsing error: {0}")]
    Archive(String),

    #[error("Bundle structure error: {0}")]
    Bundle(String),
}

/// Result type alias
pub type Result<T> = std::result::Result<T, Error>;
