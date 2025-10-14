//! iWork Archive Format Support
//!
//! This module provides support for parsing Apple's iWork file formats
//! (Pages, Keynote, Numbers) which use the IWA (iWork Archive) format.
//!
//! ## iWork File Structure
//!
//! iWork documents are bundles containing:
//! - `Index.zip`: Contains IWA files with serialized objects
//! - `Data/`: Directory containing media assets
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

pub mod snappy;
pub mod varint;
pub mod archive;
pub mod bundle;
pub mod registry;
pub mod object_index;
pub mod protobuf;

/// High-level iWork document types
pub mod document;

/// Re-export commonly used types
pub use archive::{ArchiveInfo, MessageInfo};
pub use bundle::Bundle;
pub use document::Document;
pub use snappy::SnappyStream;

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
