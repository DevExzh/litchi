//! Litchi - A Rust library for parsing Microsoft Office file formats
//!
//! This library provides efficient parsing of OLE2 (Object Linking and Embedding)
//! and OOXML (Office Open XML) file formats used by Microsoft Office.
//!
//! # Features
//!
//! - **OLE2 Parser**: Parse legacy Microsoft Office files (.doc, .xls, .ppt)
//! - **DOC Reader**: Parse legacy Word documents (.doc)
//! - **PPT Reader**: Parse legacy PowerPoint presentations (.ppt)
//! - **OOXML Parser**: Parse modern Office files (.docx, .xlsx, .pptx)
//! - **Zero-copy parsing**: Minimizes memory allocations for better performance
//! - **Metadata extraction**: Extract document properties and metadata
//!
//! # Example - Reading a DOC file
//!
//! ```no_run
//! use litchi::ole::doc::Package;
//!
//! # fn main() -> Result<(), Box<dyn std::error::Error>> {
//! // Open a .doc file
//! let mut pkg = Package::open("document.doc")?;
//! let doc = pkg.document()?;
//!
//! // Extract all text
//! let text = doc.text()?;
//! println!("Document text: {}", text);
//!
//! // Access paragraphs
//! for para in doc.paragraphs()? {
//!     println!("Paragraph: {}", para.text()?);
//! }
//! # Ok(())
//! # }
//! ```
//!
//! # Example - Reading a DOCX file
//!
//! ```no_run
//! use litchi::ooxml::docx::Package;
//!
//! # fn main() -> Result<(), Box<dyn std::error::Error>> {
//! // Open a .docx file
//! let pkg = Package::open("document.docx")?;
//! let doc = pkg.document()?;
//!
//! // Extract all text
//! let text = doc.text()?;
//! println!("Document text: {}", text);
//! # Ok(())
//! # }
//! ```
//!
//! # Example - Reading a PPT file
//!
//! ```no_run
//! use litchi::ole::ppt::Package;
//!
//! # fn main() -> Result<(), Box<dyn std::error::Error>> {
//! // Open a .ppt file
//! let mut pkg = Package::open("presentation.ppt")?;
//! let pres = pkg.presentation()?;
//!
//! // Extract all text
//! let text = pres.text()?;
//! println!("Presentation text: {}", text);
//!
//! // Access slides
//! for slide in pres.slides()? {
//!     println!("Slide: {}", slide.text()?);
//! }
//! # Ok(())
//! # }
//! ```
//!
//! # Example - Low-level OLE access
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
//! # Ok(())
//! # }
//! ```

/// OLE2 (Object Linking and Embedding) file format parser
///
/// This module provides functionality to parse OLE2 structured storage files,
/// which are used by legacy Microsoft Office formats (.doc, .xls, .ppt).
///
/// The `ole` module also contains the `doc` submodule for parsing legacy
/// Word documents, since .doc files are OLE2-based.
pub mod ole;

/// OOXML (Office Open XML) file format parser
///
/// This module provides functionality to parse modern Office formats
/// (.docx, .xlsx, .pptx).
pub mod ooxml;

// Re-export commonly used types for convenience
pub use ole::{doc, ppt};

// Re-export shape types for PPT
pub use ole::ppt::{shapes, Shape, TextBox, Placeholder, PlaceholderType, PlaceholderSize, AutoShape};
