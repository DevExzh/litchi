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
//! # Quick Start - Word Documents (Read)
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
//! # Quick Start - Word Documents (Write)
//!
//! ```no_run
//! use litchi::ooxml::docx::Package;
//!
//! # fn main() -> Result<(), Box<dyn std::error::Error>> {
//! // Create a new empty document
//! let mut pkg = Package::new()?;
//!
//! // Save the document
//! pkg.save("new_document.docx")?;
//!
//! // Open and verify
//! let reopened = Package::open("new_document.docx")?;
//! let doc = reopened.document()?;
//! println!("Created document with {} paragraphs", doc.paragraph_count()?);
//! # Ok(())
//! # }
//! ```
//!
//! # Quick Start - PowerPoint Presentations (Read)
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
//! # Quick Start - PowerPoint Presentations (Write)
//!
//! ```no_run
//! use litchi::ooxml::pptx::Package;
//!
//! # fn main() -> Result<(), Box<dyn std::error::Error + Send + Sync>> {
//! // Create a new empty presentation
//! let mut pkg = Package::new()?;
//!
//! // Save the presentation
//! pkg.save("new_presentation.pptx")?;
//!
//! // Open and verify
//! let reopened = Package::open("new_presentation.pptx")?;
//! let pres = reopened.presentation()?;
//! println!("Created presentation with {} slides", pres.slide_count()?);
//! # Ok(())
//! # }
//! ```
//!
//! # Quick Start - Excel Workbooks (Write)
//!
//! ```no_run
//! use litchi::ooxml::xlsx::Workbook;
//! use litchi::sheet::WorkbookTrait;
//!
//! # fn main() -> Result<(), Box<dyn std::error::Error + Send + Sync>> {
//! // Create a new empty workbook
//! let mut workbook = Workbook::create()?;
//!
//! // Save the workbook
//! workbook.save("new_workbook.xlsx")?;
//!
//! // Open and verify
//! let reopened = Workbook::open("new_workbook.xlsx")?;
//! println!("Created workbook with {} worksheets", reopened.worksheet_count());
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
//! - `Document` - Unified Word document interface (.doc and .docx)
//! - `Presentation` - Unified PowerPoint interface (.ppt and .pptx)
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
//! - `doc` - Direct access to the legacy Word parser and writer
//! - `ppt` - Direct access to the legacy PowerPoint parser and writer
//! - `xls` - Direct access to legacy Excel BIFF parsers and writers
//! - `ooxml` - Direct access to OOXML format parsers
//!
//! Most users should use the high-level API and only access low-level modules
//! when format-specific features are needed.

#![forbid(unsafe_code)]

/// Common types, traits, and utilities shared across formats.
///
/// Re-export of the `litchi-core` crate under the historical `common` path.
///
/// Smart-detection items (`DetectedFormat`, `detect_format_smart`) live in
/// the umbrella's `detection_smart` module; this shim re-exports them
/// under `common::detection` to preserve API compat.
pub mod common {
    pub use litchi_core::*;

    // Re-export top-level smart-detection entry points so legacy callers like
    // `litchi::common::detect_file_format(...)` keep working.
    #[cfg(any(
        feature = "doc",
        feature = "ppt",
        feature = "xls",
        feature = "ooxml",
        feature = "iwa",
        feature = "odf",
        feature = "rtf"
    ))]
    pub use crate::detection_smart::{detect_file_format, detect_file_format_from_bytes};

    /// Detection re-exports — merges `litchi-core`'s signature detection with
    /// the umbrella's smart-detection entry points.
    pub mod detection {
        #[cfg(feature = "iwa")]
        pub use crate::detection_smart::detect_iwork_format_from_path;
        #[cfg(any(
            feature = "doc",
            feature = "ppt",
            feature = "xls",
            feature = "ooxml",
            feature = "iwa",
            feature = "odf",
            feature = "rtf"
        ))]
        pub use crate::detection_smart::{
            DetectedFormat, detect_file_format, detect_file_format_from_bytes,
            detect_format_from_reader, detect_format_smart,
        };
        pub use litchi_core::detection::*;
    }
}

// Smart format detection (depends on per-format crates; can't live in litchi-core).
#[cfg(any(
    feature = "doc",
    feature = "ppt",
    feature = "xls",
    feature = "ooxml",
    feature = "iwa",
    feature = "odf",
    feature = "rtf"
))]
pub mod detection_smart;

mod metadata_ext;

pub use metadata_ext::MetadataYaml;

/// Unified Word document API
///
/// Provides format-agnostic interface for .doc, .docx, .rtf, and .odt files.
/// Use [`Document::open()`] to get started.
#[cfg(any(
    feature = "doc",
    feature = "ooxml",
    feature = "rtf",
    feature = "odf",
    feature = "iwa"
))]
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
/// **Note**: This requires at least one presentation-format feature to be enabled.
#[cfg(any(feature = "ppt", feature = "ooxml", feature = "odf", feature = "iwa"))]
pub mod presentation;

/// Unified Excel/Spreadsheet API (.xls, .xlsx, .xlsb, .ods, .numbers)
///
/// Requires the corresponding feature flags:
/// - `xls` for .xls
/// - `ooxml` for .xlsx and .xlsb
/// - `odf` for .ods
/// - `iwa` for .numbers
#[cfg(any(feature = "xls", feature = "ooxml", feature = "odf", feature = "iwa"))]
pub mod sheet;

/// Markdown conversion module
///
/// Provides functionality to convert Office documents and presentations to Markdown.
/// Use the [`markdown::ToMarkdown`] trait on Document or Presentation types.
pub mod markdown;

// Low-level format-specific modules (advanced use)
/// Legacy Word binary parser (`.doc`).
///
/// This is the canonical low-level facade for the independently owned
/// `litchi-doc` package. Compound-file primitives live in `litchi-cfb`, while
/// legacy PowerPoint support is exposed independently through `litchi::ppt`.
/// Most users should use the high-level [`Document`] API instead.
///
/// **Note**: This requires the `doc` feature to be enabled.
#[cfg(feature = "doc")]
pub mod doc {
    pub use litchi_doc::*;
}

/// Legacy PowerPoint binary parser and writer (`.ppt`).
///
/// This is the canonical low-level facade for the independently owned
/// `litchi-ppt` package. It is independent from [`doc`].
///
/// **Note**: This requires the `ppt` feature to be enabled.
#[cfg(feature = "ppt")]
pub mod ppt {
    pub use litchi_ppt::*;
}

/// Legacy Excel BIFF parser and writer (`.xls`).
///
/// This is the canonical low-level facade for the independently owned
/// `litchi-xls` package. It is independent from [`doc`].
///
/// **Note**: This requires the `xls` feature to be enabled.
#[cfg(feature = "xls")]
pub mod xls {
    pub use litchi_xls::*;
}

/// Shared DrawingML chart and diagram vocabulary.
///
/// This is the concise, host-neutral facade for drawing types shared by OOXML
/// documents, presentations, and workbooks.
///
/// **Note**: This requires the `ooxml` feature to be enabled.
#[cfg(feature = "ooxml")]
pub mod drawing {
    pub use litchi_drawingml::*;
}

/// OOXML format parser (modern .docx, .pptx files)
///
/// This module provides direct access to OOXML parsing functionality.
/// Most users should use the high-level [`Document`] and [`Presentation`]
/// APIs instead, which automatically handle format detection.
///
/// **Note**: This requires the `ooxml` feature to be enabled.
#[cfg(feature = "ooxml")]
pub mod ooxml {
    pub use litchi_ooxml::*;
    pub use litchi_ooxml_common::{custom, custom_xml, embedded, ribbon, web};
}

/// Runtime-neutral Microsoft Office cryptography.
///
/// The concrete format packages retain responsibility for locating encrypted
/// records; this facade exposes the canonical bounded cryptographic codecs and
/// OOXML encrypted-package service without a CFB type in their APIs.
#[cfg(feature = "ooxml_encryption")]
pub mod crypto {
    pub use litchi_crypto::*;
}

/// Trust-neutral Office signature authoring and verification.
///
/// The concise facade re-exports the canonical types without aliases. Format
/// packages provide the container-specific `signatures`, `sign`, `resign`, and
/// `unsign` operations.
#[cfg(feature = "sign")]
pub mod sign {
    pub use litchi_sign::*;
}

/// Formula module
///
/// This module provides functionality to parse and convert mathematical formulas between different formats.
///
/// **Note**: This requires the `formula` feature to be enabled.
#[cfg(feature = "formula")]
pub mod formula {
    pub use litchi_formula::*;
}

/// iWork Archive Format Support
///
/// Provides support for parsing Apple's iWork file formats
/// (Pages, Keynote, Numbers) which use the IWA (iWork Archive) format.
/// Use [`iwa::Document::open()`] to get started.
///
/// **Note**: This requires the `iwa` feature to be enabled.
#[cfg(feature = "iwa")]
pub mod iwa {
    pub use litchi_iwa::*;
}

/// OpenDocument Format (ODF) Support
///
/// Provides the layered OpenDocument family facades for `.odt`, `.ods`, and
/// `.odp`. Use [`odf::odt`], [`odf::ods`], or [`odf::odp`] to access the
/// format-specific APIs; the high-level [`Document`], [`Workbook`], and
/// [`Presentation`] facades auto-detect packaged files.
///
/// **Note**: This requires the `odf` feature to be enabled.
#[cfg(feature = "odf")]
pub mod odf {
    pub use litchi_odf_common as common;
    pub use litchi_odf_common::detect;
    pub use litchi_odp::{self as odp, Presentation, PresentationBuilder, Shape, Slide};
    pub use litchi_ods::{self as ods, MutableSpreadsheet, Spreadsheet, SpreadsheetBuilder};
    pub use litchi_odt::elements::text::Paragraph;
    pub use litchi_odt::{self as odt, Document, DocumentBuilder, MutableDocument};
}

/// RTF (Rich Text Format) Support
///
/// Provides high-performance parsing of RTF documents with support for RTF 1.9.1.
/// RTF documents are automatically integrated with the unified Document API.
/// Use [`Document::open()`] to parse RTF files.
///
/// **Note**: This requires the `rtf` feature to be enabled.
#[cfg(feature = "rtf")]
pub mod rtf {
    pub use litchi_rtf::*;
}

/// Shared font embedding and subsetting module
///
/// Provides functionality for font discovery, loading, and subsetting
/// to reduce the size of embedded fonts in documents.
///
/// **Note**: This requires the `fonts` feature to be enabled.
#[cfg(feature = "fonts")]
pub mod fonts {
    pub use litchi_fonts::*;
}

// Re-export high-level APIs
pub use common::{Error, Result};

#[cfg(any(
    feature = "doc",
    feature = "ooxml",
    feature = "rtf",
    feature = "odf",
    feature = "iwa"
))]
pub use document::{Document, DocumentElement};

#[cfg(any(feature = "ppt", feature = "ooxml", feature = "odf", feature = "iwa"))]
pub use presentation::Presentation;

#[cfg(any(feature = "xls", feature = "ooxml", feature = "odf", feature = "iwa"))]
pub use sheet::Workbook;

// Re-export commonly used types
pub use common::{FileFormat, Length, PlaceholderType, RGBColor, ShapeType};

#[cfg(any(
    feature = "doc",
    feature = "ppt",
    feature = "xls",
    feature = "ooxml",
    feature = "iwa",
    feature = "odf",
    feature = "rtf"
))]
pub use common::{detect_file_format, detect_file_format_from_bytes};
