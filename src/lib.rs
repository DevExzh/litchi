//! Litchi - High-performance Rust library for Microsoft Office file formats
//!
//! Litchi provides a unified, user-friendly API for parsing Microsoft Office documents
//! in both legacy (OLE2) and modern (OOXML) formats. The library automatically detects
//! file formats and provides consistent interfaces for working with documents and presentations.
//!
//! # Features
//!
//! - **Unified API**: Work with .doc and .docx files using the same interface
//! - **Format Auto-detection**: No need to specify file format - it's detected automatically
//! - **High Performance**: Zero-copy parsing with SIMD optimizations where possible
//! - **Production Ready**: Clean API inspired by python-docx and python-pptx
//! - **Type Safe**: Leverages Rust's type system for safety and correctness
//!
//! # Quick Start - Word Documents
//!
//! ```no_run
//! use litchi::Document;
//!
//! # fn main() -> Result<(), litchi::Error> {
//! // Open any Word document (.doc or .docx) - format auto-detected
//! let doc = Document::open("document.doc")?;
//!
//! // Extract all text
//! let text = doc.text()?;
//! println!("Document text: {}", text);
//!
//! // Access paragraphs
//! for para in doc.paragraphs()? {
//!     println!("Paragraph: {}", para.text()?);
//!     
//!     // Access runs with formatting
//!     for run in para.runs()? {
//!         println!("  Text: {}", run.text()?);
//!         if run.bold()? == Some(true) {
//!             println!("    (bold)");
//!         }
//!     }
//! }
//!
//! // Access tables
//! for table in doc.tables()? {
//!     println!("Table with {} rows", table.row_count()?);
//!     for row in table.rows()? {
//!         for cell in row.cells()? {
//!             println!("  Cell: {}", cell.text()?);
//!         }
//!     }
//! }
//! # Ok(())
//! # }
//! ```
//!
//! # Quick Start - PowerPoint Presentations
//!
//! ```no_run
//! use litchi::Presentation;
//!
//! # fn main() -> Result<(), litchi::Error> {
//! // Open any PowerPoint presentation (.ppt or .pptx) - format auto-detected
//! let pres = Presentation::open("presentation.ppt")?;
//!
//! // Extract all text
//! let text = pres.text()?;
//! println!("Presentation text: {}", text);
//!
//! // Get slide count
//! println!("Total slides: {}", pres.slide_count()?);
//!
//! // Access individual slides
//! for (i, slide) in pres.slides()?.iter().enumerate() {
//!     println!("Slide {}: {}", i + 1, slide.text()?);
//! }
//! # Ok(())
//! # }
//! ```
//!
//! # Architecture
//!
//! The library is organized into several layers:
//!
//! ## High-Level API (Recommended)
//!
//! - [`Document`] - Unified Word document interface (.doc and .docx)
//! - [`Presentation`] - Unified PowerPoint interface (.ppt and .pptx)
//!
//! These automatically detect file formats and provide a consistent API.
//!
//! ## Common Types
//!
//! - [`common::Error`] - Unified error type
//! - [`common::Result`] - Result type alias
//! - [`common::ShapeType`] - Common shape types
//! - [`common::RGBColor`] - Color representation
//! - [`common::Length`] - Measurement with units
//!
//! ## Low-Level Modules (Advanced Use)
//!
//! - [`ole`] - Direct access to OLE2 format parsers
//! - [`ooxml`] - Direct access to OOXML format parsers
//!
//! Most users should use the high-level API and only access low-level modules
//! when format-specific features are needed.

/// Common types, traits, and utilities shared across formats
pub mod common;

/// Unified Word document API
///
/// Provides format-agnostic interface for both .doc and .docx files.
/// Use [`Document::open()`] to get started.
///
/// **Note**: This requires at least one of the `ole` or `ooxml` features to be enabled.
#[cfg(any(feature = "ole", feature = "ooxml"))]
pub mod document;

/// Image processing and conversion module
///
/// Provides functionality to parse and convert Office Drawing formats
/// (EMF, WMF, PICT) to modern image standards (PNG, JPEG, WebP).
///
/// **Note**: This requires the `imgconv` feature to be enabled.
#[cfg(feature = "imgconv")]
pub mod images;

/// Unified PowerPoint presentation API
///
/// Provides format-agnostic interface for both .ppt and .pptx files.
/// Use [`Presentation::open()`] to get started.
///
/// **Note**: This requires at least one of the `ole` or `ooxml` features to be enabled.
#[cfg(any(feature = "ole", feature = "ooxml"))]
pub mod presentation;

/// Unified Excel spreadsheet API (placeholder for future functionality)
///
/// **Note**: This requires at least one of the `ole` or `ooxml` features to be enabled.
#[cfg(any(feature = "ole", feature = "ooxml"))]
pub mod sheet;

/// Markdown conversion module
///
/// Provides functionality to convert Office documents and presentations to Markdown.
/// Use the [`markdown::ToMarkdown`] trait on Document or Presentation types.
pub mod markdown;

// Low-level format-specific modules (advanced use)
/// OLE2 format parser (legacy .doc, .ppt files)
///
/// This module provides direct access to OLE2 parsing functionality.
/// Most users should use the high-level [`Document`] and [`Presentation`]
/// APIs instead, which automatically handle format detection.
///
/// **Note**: This requires the `ole` feature to be enabled.
#[cfg(feature = "ole")]
pub mod ole;

/// OOXML format parser (modern .docx, .pptx files)
///
/// This module provides direct access to OOXML parsing functionality.
/// Most users should use the high-level [`Document`] and [`Presentation`]
/// APIs instead, which automatically handle format detection.
///
/// **Note**: This requires the `ooxml` feature to be enabled.
#[cfg(feature = "ooxml")]
pub mod ooxml;

/// Formula module
///
/// This module provides functionality to parse and convert mathematical formulas between different formats.
///
/// **Note**: This requires the `formula` feature to be enabled.
#[cfg(feature = "formula")]
pub mod formula;

/// iWork Archive Format Support
///
/// Provides support for parsing Apple's iWork file formats
/// (Pages, Keynote, Numbers) which use the IWA (iWork Archive) format.
/// Use [`iwa::Document::open()`] to get started.
///
/// **Note**: This requires the `iwa` feature to be enabled.
#[cfg(feature = "iwa")]
pub mod iwa;

/// OpenDocument Format (ODF) Support
///
/// Provides unified APIs for working with OpenDocument files (.odt, .ods, .odp).
/// The format is automatically detected and handled transparently.
/// Use [`odf::Document`], [`odf::Spreadsheet`], or [`odf::Presentation`] to get started.
///
/// **Note**: This requires the `odf` feature to be enabled.
#[cfg(feature = "odf")]
pub mod odf;

// Re-export high-level APIs
pub use common::{Error, Result};

#[cfg(any(feature = "ole", feature = "ooxml"))]
pub use document::Document;

#[cfg(any(feature = "ole", feature = "ooxml"))]
pub use presentation::Presentation;

// Re-export commonly used types
pub use common::{
    FileFormat, Length, PlaceholderType, RGBColor, ShapeType, detect_file_format,
    detect_file_format_from_bytes,
};
