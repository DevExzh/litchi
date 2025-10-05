//! Litchi - A Rust library for parsing Microsoft Office file formats
//!
//! This library provides efficient parsing of OLE2 (Object Linking and Embedding)
//! and OOXML (Office Open XML) file formats used by Microsoft Office.
//!
//! # Features
//!
//! - **OLE2 Parser**: Parse legacy Microsoft Office files (.doc, .xls, .ppt)
//! - **OOXML Parser**: Parse modern Office files (.docx, .xlsx, .pptx) - Coming soon
//! - **Zero-copy parsing**: Minimizes memory allocations for better performance
//! - **Metadata extraction**: Extract document properties and metadata
//!
//! # Example
//!
//! ```no_run
//! use std::fs::File;
//! use litchi::ole::OleFile;
//!
//! # fn main() -> Result<(), Box<dyn std::error::Error>> {
//! // Open an OLE file
//! let file = File::open("document.doc")?;
//! let mut ole = OleFile::open(file)?;
//!
//! // List all streams
//! let streams = ole.list_streams();
//! for stream in streams {
//!     println!("Stream: {:?}", stream);
//! }
//!
//! // Open a specific stream
//! let data = ole.open_stream(&["WordDocument"])?;
//! println!("Stream size: {} bytes", data.len());
//!
//! // Extract metadata
//! let metadata = ole.get_metadata()?;
//! if let Some(title) = metadata.title {
//!     println!("Title: {}", title);
//! }
//! # Ok(())
//! # }
//! ```

/// OLE2 (Object Linking and Embedding) file format parser
///
/// This module provides functionality to parse OLE2 structured storage files,
/// which are used by legacy Microsoft Office formats (.doc, .xls, .ppt).
pub mod ole;

/// OOXML (Office Open XML) file format parser
///
/// This module provides functionality to parse modern Office formats
/// (.docx, .xlsx, .pptx). Currently under development.
pub mod ooxml;
