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
//! ```ignore
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
//! ```ignore
//! use litchi::docx::Package;
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
//! ```ignore
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
//! ```ignore
//! use litchi::pptx::Package;
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
//! ```ignore
//! use litchi::xlsx::Workbook;
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
//! - `ppt` - Direct access to the legacy PowerPoint parser and writer
//! - `xls` - Direct access to legacy Excel BIFF parsers and writers
//! - `docx`, `pptx`, `xlsx`, and `xlsb` - Direct access to standalone OOXML owners
//!
//! Most users should use the high-level API and only access low-level modules
//! when format-specific features are needed.

#![forbid(unsafe_code)]

/// Common types, traits, and utilities shared across formats.
///
/// Re-export of the `litchi-core` crate under the concise `common` path.
///
/// Smart-detection items (`DetectedFormat`, `detect_format_smart`) live in
/// the umbrella's `detection_smart` module; this facade also exposes them
/// under `common::detection` beside the core detection vocabulary.
pub mod common {
    pub use litchi_core::*;

    // Re-export the smart-detection entry points beside the core vocabulary.
    #[cfg(any(
        feature = "doc",
        feature = "docx",
        feature = "ppt",
        feature = "pptx",
        feature = "xls",
        feature = "xlsx",
        feature = "xlsb",
        feature = "rtf",
        feature = "odt",
        feature = "ods",
        feature = "odp",
        feature = "pages",
        feature = "keynote",
        feature = "numbers"
    ))]
    pub use crate::detection_smart::{detect_file_format, detect_file_format_from_bytes};

    /// Detection re-exports — merges `litchi-core`'s signature detection with
    /// the umbrella's smart-detection entry points.
    pub mod detection {
        #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
        pub use crate::detection_smart::detect_format_smart_with_limits;
        #[cfg(any(
            feature = "doc",
            feature = "ppt",
            feature = "xls",
            feature = "docx",
            feature = "pptx",
            feature = "xlsx",
            feature = "xlsb",
            feature = "pages",
            feature = "keynote",
            feature = "numbers",
            feature = "odt",
            feature = "ods",
            feature = "odp",
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
    feature = "docx",
    feature = "ppt",
    feature = "pptx",
    feature = "xls",
    feature = "xlsx",
    feature = "xlsb",
    feature = "pages",
    feature = "keynote",
    feature = "numbers",
    feature = "odt",
    feature = "ods",
    feature = "odp",
    feature = "rtf"
))]
pub mod detection_smart;

#[cfg(feature = "yaml")]
mod metadata_ext;

#[cfg(feature = "yaml")]
pub use metadata_ext::MetadataYaml;

/// Unified Word document API
///
/// Provides format-agnostic interface for .doc, .docx, .rtf, and .odt files.
/// Use [`Document::open()`] to get started.
#[cfg(any(
    feature = "doc",
    feature = "docx",
    feature = "rtf",
    feature = "odt",
    feature = "pages"
))]
pub mod document;

/// Image processing and conversion module
///
/// Provides functionality to parse and convert Office Drawing formats
/// (EMF, WMF, PICT) to modern image standards (PNG, JPEG, WebP).
///
/// **Note**: This requires the `images` feature to be enabled.
#[cfg(feature = "images")]
pub mod images;

/// Unified PowerPoint presentation API
///
/// Provides format-agnostic interface for both .ppt and .pptx files.
/// Use [`Presentation::open()`] to get started.
///
/// **Note**: This requires at least one presentation-format feature to be enabled.
#[cfg(any(
    feature = "ppt",
    feature = "pptx",
    feature = "odp",
    feature = "keynote"
))]
pub mod presentation;

/// Unified Excel/Spreadsheet API (.xls, .xlsx, .xlsb, .ods, .numbers)
///
/// Requires the corresponding feature flags:
/// - `xls` for .xls
/// - `xlsx` for .xlsx
/// - `xlsb` for .xlsb
/// - `ods` for .ods
/// - `numbers` for .numbers
#[cfg(feature = "sheet")]
pub mod sheet;

/// Markdown conversion module
///
/// Provides functionality to convert Office documents and presentations to Markdown.
/// Use the [`markdown::ToMarkdown`] trait on Document or Presentation types.
#[cfg(feature = "markdown")]
pub mod markdown;

/// Compound File Binary (CFB) container primitives.
#[cfg(feature = "cfb")]
pub mod cfb {
    pub use litchi_cfb::*;
}

/// Shared OLE property-set and metadata primitives.
#[cfg(feature = "ole")]
pub mod ole {
    pub use litchi_ole_common::*;
}

/// Legacy Word binary parser and writer (`.doc`).
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
/// **Note**: This requires the `drawingml` feature to be enabled.
#[cfg(feature = "drawingml")]
pub mod drawing {
    pub use litchi_drawingml::*;
}

/// WordprocessingML (`.docx`) package and semantic APIs.
#[cfg(feature = "docx")]
pub mod docx {
    pub use litchi_docx::*;
}

/// PresentationML (`.pptx`) package and semantic APIs.
#[cfg(feature = "pptx")]
pub mod pptx {
    pub use litchi_pptx::*;
}

/// SpreadsheetML (`.xlsx`) package and semantic APIs.
#[cfg(feature = "xlsx")]
pub mod xlsx {
    pub use litchi_xlsx::*;
}

/// Binary SpreadsheetML (`.xlsb`) package and semantic APIs.
#[cfg(feature = "xlsb")]
pub mod xlsb {
    pub use litchi_xlsb::*;
}

/// Open Packaging Conventions package graph used by the standalone OOXML owners.
#[cfg(feature = "opc")]
pub mod opc {
    pub use litchi_opc::*;
}

/// Shared OOXML vocabulary and package services.
#[cfg(feature = "ooxml-common")]
pub mod ooxml_common {
    #[cfg(feature = "encryption")]
    pub use litchi_ooxml_common::package_encryption;
    #[cfg(feature = "vba-inspection")]
    pub use litchi_ooxml_common::vba;
    pub use litchi_ooxml_common::{
        Error, Keywords, Props, Result, XmlError, custom, custom_xml, embedded, external_link, mce,
        properties, relationships, ribbon, spreadsheet_xml_maps, web, xml, xml_name,
    };
}

#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
pub(crate) fn map_ooxml_error<E: std::fmt::Display>(error: E) -> litchi_core::Error {
    // The unified facade has no concrete OOXML error variant, but these
    // failures are format/graph validation failures rather than generic
    // application errors. Preserve that distinction for callers.
    litchi_core::Error::InvalidFormat(error.to_string())
}

/// Runtime-neutral Microsoft Office cryptography.
///
/// The concrete format packages retain responsibility for locating encrypted
/// records; this facade exposes the canonical bounded cryptographic codecs and
/// OOXML encrypted-package service without a CFB type in their APIs.
#[cfg(feature = "encryption")]
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

/// Apple Pages package and semantic APIs.
#[cfg(feature = "pages")]
pub mod pages {
    pub use litchi_pages::*;

    /// Archive-free, immutable Pages values.
    ///
    /// This namespace is the strict semantic reader surface. Package
    /// preservation and editing remain available from [`crate::pages::Package`] but
    /// are not part of this contract; no native archive, generated message, or
    /// compatibility projection is re-exported here. The checked source and
    /// semantic limit profiles are re-exported alongside
    /// [`DocumentReadOptions`] so a caller can keep the complete reader API in
    /// this namespace.
    pub mod semantic {
        pub use litchi_pages::{
            Body, Document, DocumentReadOptions, DocumentSourceLimitKind, DocumentSourceLimits,
            DocumentSourceLimitsError, DocumentStats, Error, IoKind, Position, ReadError,
            ReadLimitKind, Result, Root, Section, SectionSelector, SectionType, SelectorError,
            SelectorResult, SemanticLimitKind, SemanticLimits, SemanticLimitsError, TextPosition,
            TextSpan,
        };

        /// Result returned by the archive-free Pages source readers.
        ///
        /// This is separate from [`Result`], which is the leaf crate's
        /// semantic-construction result. Keeping the source error visible at
        /// this boundary preserves typed wrong-format, limit, and malformed
        /// input failures for callers.
        pub type ReadResult<T> = std::result::Result<T, ReadError>;

        /// Read one immutable Pages snapshot from a complete package or
        /// app-authored package directory.
        ///
        /// The returned document contains only archive-free semantic values;
        /// package members and native object identifiers remain below the
        /// package API. A source belonging to another iWork family is
        /// reported as [`ReadError::NotPages`].
        pub fn open(path: impl AsRef<std::path::Path>) -> ReadResult<Document> {
            Document::open(path)
        }

        /// Read one immutable Pages snapshot from package bytes.
        ///
        /// The source is detected and routed by the Pages owner, so callers
        /// do not need to inspect package markers or native identifiers.
        pub fn from_bytes(bytes: &[u8]) -> ReadResult<Document> {
            Document::from_bytes(bytes)
        }
    }
}

/// Apple Keynote package and semantic APIs.
#[cfg(feature = "keynote")]
pub mod keynote {
    pub use litchi_keynote::*;

    /// Archive-free, immutable Keynote values.
    ///
    /// This namespace intentionally excludes the package adapter and its
    /// editing transactions. Use [`crate::keynote::Package`] when exact package
    /// preservation or editing is required. The source and semantic limit
    /// profiles, selector errors, and immutable value errors are available
    /// here as well, so no package-layer type is needed to configure a read.
    pub mod semantic {
        pub use litchi_keynote::show::{Mode, Settings, Show, Size};
        pub use litchi_keynote::{
            Build, Document, DocumentIoKind, DocumentReadError, DocumentReadLimitKind,
            DocumentReadOptions, DocumentSemanticLimitKind, DocumentSemanticLimits,
            DocumentSemanticLimitsError, DocumentSourceLimitKind, DocumentSourceLimits,
            DocumentSourceLimitsError, DocumentStats, Error, Position, Result, Seconds, Slide,
            SlideSelector, SlideSelectorError, SlideSelectorResult, TextPosition, TextSpan,
            Transition,
        };

        /// Result returned by the archive-free Keynote source readers.
        ///
        /// This is separate from [`Result`], which remains the focused
        /// semantic-operation result. The source reader keeps Keynote's
        /// typed I/O, limit, malformed-input, allocation, and wrong-format
        /// errors available to callers.
        pub type ReadResult<T> = std::result::Result<T, DocumentReadError>;

        /// Read one immutable Keynote snapshot from a complete package or
        /// app-authored package directory.
        pub fn open(path: impl AsRef<std::path::Path>) -> ReadResult<Document> {
            Document::open(path)
        }

        /// Read one immutable Keynote snapshot from package bytes.
        ///
        /// Format detection and semantic projection stay owned by the
        /// Keynote reader; no native identifiers enter this facade.
        pub fn from_bytes(bytes: &[u8]) -> ReadResult<Document> {
            Document::from_bytes(bytes)
        }
    }
}

/// Apple Numbers package and semantic APIs.
#[cfg(feature = "numbers")]
pub mod numbers {
    // Keep the format facade explicit: a leaf-package glob would silently
    // publish any future package, generated, or physical helper module.
    // Package operations remain available by their established names, while
    // the archive-free reader surface below is deliberately allowlisted.
    pub use litchi_numbers::{
        AddressError, Cell, CellPosition, CellRange, ColumnDeletion, CoordinateError,
        DEFAULT_MAX_TEXT_BYTES, Dimension, Dimensions, Document, DocumentError, DocumentLimitKind,
        DocumentLimits, DocumentLimitsError, DocumentReadError, DocumentReadLimitKind,
        DocumentReadOptions, DocumentResult, DocumentSourceLimitKind, DocumentSourceLimits,
        DocumentSourceLimitsError, DocumentStats, Grid, GridBudget, InsertError, InsertResult,
        IoKind, MAX_MATERIALIZED_CELLS, MAX_OBJECTS, MAX_REFERENCES, MAX_SHEETS, MAX_TABLES,
        MergeReader, Package, PackageError, PackageLimits, PackagePayloadLimitKind,
        PackageReadOptions, PackageResourceError, PackageResult, PackageSemanticLimits,
        PackageSemanticLimitsError, PackageSemanticPath, Points, Position, Range, RowDeletion,
        SemanticLimitKind, Sheet, SheetBuilder, SheetSelector, Size, Table, TableBuilder,
        TableCellComment, TableCellCommentCommit, TableCellCommentDiagnostics,
        TableCellCommentEdit, TableCellCommentError, TableCellCommentLimitKind,
        TableCellCommentPatch, TableCellCommentPath, TableError, TableLockCommit,
        TableLockDiagnostics, TableLockEdit, TableLockError, TableLockLimitKind, TableLockPatch,
        TableMergesCommit, TableMergesDiagnostics, TableMergesEdit, TableMergesError,
        TableMergesLimitKind, TableMergesPatch, TableSelector, TableSelectorError, View,
        WriteError, compatibility_tables_from_bytes, compatibility_tables_from_bytes_with_options,
    };
    pub use litchi_numbers::{cell, document, formula, names, package, selector, sheet, table};

    /// Archive-free, immutable Numbers values.
    ///
    /// This namespace intentionally excludes the package adapter and the
    /// global compatibility-table projection. Use [`crate::numbers`] when exact
    /// package preservation or editing is required via [`crate::numbers::Package`].
    /// The checked source and semantic limit profiles and their errors remain
    /// in this namespace for a complete immutable-reader boundary.
    pub mod semantic {
        /// Archive-free typed Numbers cell values.
        ///
        /// These values describe what a caller can observe or stage at a
        /// semantic cell boundary. They contain no package objects, native
        /// identifiers, protobuf messages, or wire payloads.
        pub mod cell {
            pub use litchi_numbers::cell::{FiniteF64, FiniteF64Error, Type, Value};
        }

        /// Archive-free, dependency-free Numbers formula vocabulary.
        ///
        /// These are the small typed building blocks that do not retain a
        /// package snapshot. Formula expressions and source-bound formula
        /// handles remain in the focused package API.
        pub mod formula {
            pub use litchi_numbers::formula::{
                AxisReference, BinaryOperator, CachedValue, CellReference, Error as FormulaError,
            };
        }

        /// Archive-free Numbers table coordinates and cell presence values.
        ///
        /// The nested `cells` module mirrors the focused format crate so a
        /// caller can use typed A1 coordinates and presence-preserving values
        /// without importing the package adapter.
        pub mod table {
            pub use litchi_numbers::{
                AddressError, CellPosition, CellRange, CoordinateError, Dimensions, View,
            };

            pub mod cells {
                pub use litchi_numbers::table::cells::{State, Storage};
            }
        }

        pub use litchi_numbers::{
            Document, DocumentError as Error, DocumentLimitKind as LimitKind,
            DocumentLimits as Limits, DocumentLimitsError as LimitsError, DocumentReadError,
            DocumentReadLimitKind as ReadLimitKind, DocumentReadOptions, DocumentResult as Result,
            DocumentSourceLimitKind, DocumentSourceLimits, DocumentSourceLimitsError,
            DocumentStats, Sheet, SheetSelector, Table, TableSelector, TableSelectorError,
        };

        /// Result returned by the archive-free Numbers source readers.
        ///
        /// This is separate from [`Result`], which is the semantic document
        /// construction result. The source reader's typed error retains
        /// format routing and bounded-ingress diagnostics.
        pub type ReadResult<T> = std::result::Result<T, DocumentReadError>;

        /// Read one immutable Numbers snapshot from a complete package or
        /// app-authored package directory.
        pub fn open(path: impl AsRef<std::path::Path>) -> ReadResult<Document> {
            Document::open(path)
        }

        /// Read one immutable Numbers snapshot from package bytes.
        ///
        /// The Numbers owner performs family detection and semantic
        /// projection, so callers never need native identifiers or wire data.
        pub fn from_bytes(bytes: &[u8]) -> ReadResult<Document> {
            Document::from_bytes(bytes)
        }
    }
}

/// Format-neutral, bounded Apple iWork reading APIs.
///
/// This facade detects Pages, Keynote, and Numbers from one immutable package
/// snapshot, then publishes only archive-free semantic values. Enable the
/// aggregate `iwork` feature to use it.
#[cfg(feature = "iwork")]
pub mod iwork;

/// OpenDocument Presentation (`.odp`) package and semantic APIs.
#[cfg(feature = "odp")]
pub mod odp {
    pub use litchi_odp::*;
}

/// OpenDocument Spreadsheet (`.ods`) package and semantic APIs.
#[cfg(feature = "ods")]
pub mod ods {
    pub use litchi_ods::*;
}

/// OpenDocument Text (`.odt`) package and semantic APIs.
#[cfg(feature = "odt")]
pub mod odt {
    pub use litchi_odt::*;
}

/// Shared OpenDocument vocabulary and detection services.
#[cfg(feature = "odf-common")]
pub mod odf_common {
    pub use litchi_odf_common::*;
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

/// Spreadsheet formula evaluation primitives.
#[cfg(feature = "eval")]
pub mod eval {
    pub use litchi_eval::*;
}

// Re-export high-level APIs
pub use common::{Error, Result};

#[cfg(any(
    feature = "doc",
    feature = "docx",
    feature = "rtf",
    feature = "odt",
    feature = "pages"
))]
pub use document::{Document, DocumentElement};

#[cfg(any(
    feature = "ppt",
    feature = "pptx",
    feature = "odp",
    feature = "keynote"
))]
pub use presentation::Presentation;

#[cfg(feature = "pptx")]
pub use presentation::SlideDescriptor;

#[cfg(any(
    feature = "xls",
    feature = "xlsx",
    feature = "xlsb",
    feature = "ods",
    feature = "numbers"
))]
pub use sheet::Workbook;

// Re-export commonly used types
pub use common::{FileFormat, Length, PlaceholderType, RGBColor, ShapeType};

#[cfg(any(
    feature = "doc",
    feature = "docx",
    feature = "ppt",
    feature = "pptx",
    feature = "xls",
    feature = "xlsx",
    feature = "xlsb",
    feature = "pages",
    feature = "keynote",
    feature = "numbers",
    feature = "odt",
    feature = "ods",
    feature = "odp",
    feature = "rtf"
))]
pub use common::{detect_file_format, detect_file_format_from_bytes};

#[cfg(all(
    test,
    feature = "docx",
    feature = "pptx",
    feature = "xlsx",
    feature = "xlsb",
    feature = "opc",
    feature = "ooxml-common"
))]
mod standalone_ooxml_facade_tests {
    #[test]
    fn exposes_each_standalone_owner_without_an_outer_namespace() {
        let _: Option<super::docx::Package> = None;
        let _: Option<super::pptx::Package> = None;
        let _: Option<super::xlsx::Workbook> = None;
        let _: Option<super::xlsb::Workbook> = None;
        let _: Option<super::docx::ReadLimits> = None;
        let _: Option<super::pptx::ReadLimits> = None;
        let _: Option<super::xlsx::ReadLimits> = None;
        let _: Option<super::xlsb::ReadLimits> = None;
        let _: Option<super::opc::OpcPackage> = None;
        let _: Option<super::ooxml_common::custom::Props> = None;
    }
}

#[cfg(all(
    test,
    feature = "odp",
    feature = "ods",
    feature = "odt",
    feature = "odf-common"
))]
mod standalone_odf_facade_tests {
    #[test]
    fn exposes_each_standalone_owner_without_an_outer_namespace() {
        let _: Option<super::odp::Presentation> = None;
        let _: Option<super::ods::Spreadsheet> = None;
        let _: Option<super::odt::Document> = None;
        let _: Option<super::odf_common::detect::Format> = None;
    }
}
