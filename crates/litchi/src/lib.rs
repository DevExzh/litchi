//! Litchi: one Rust API for Microsoft Office, OpenDocument and Apple iWork
//! documents.
//!
//! # Opening a file
//!
//! [`open`] is the entry point. It detects the format from the content in
//! one pass and returns a [`File`]: a flat enum with one variant per format,
//! plus [`File::Locked`] for an encrypted input and [`File::Signed`] for a
//! signed one. [`open_bytes`] does the same for an in-memory input, and
//! [`open_with`] and [`open_bytes_with`] take [`OpenOptions`] for limits and
//! policy. The common case needs no options and no format knowledge.
//!
//! ```
//! # #[cfg(feature = "docx")]
//! # {
//! let path = concat!(
//!     env!("CARGO_MANIFEST_DIR"),
//!     "/../../test-data/ooxml/docx/documentProperties.docx"
//! );
//! let file = litchi::open(path)?;
//! assert_eq!(file.format(), Some(litchi::FileFormat::Docx));
//! assert_eq!(file.family(), Some(litchi::Family::Document));
//!
//! // Reads every format shares are methods on `File`.
//! assert_eq!(file.text()?.trim(), "Hello World!");
//! assert!(file.tolerance_report()?.is_clean());
//!
//! // Family reads go through the view of the file's family.
//! let document = file.document()?;
//! assert_eq!(document.paragraph_count()?, 1);
//! # }
//! # Ok::<(), litchi::Error>(())
//! ```
//!
//! A view refuses a file of another family, and names the format it found:
//!
//! ```
//! # #[cfg(all(feature = "docx", feature = "xlsx"))]
//! # {
//! # let path = concat!(
//! #     env!("CARGO_MANIFEST_DIR"),
//! #     "/../../test-data/ooxml/docx/documentProperties.docx"
//! # );
//! let file = litchi::open(path)?;
//! let error = file.workbook().unwrap_err();
//! assert_eq!(error.code(), litchi::Code::UnsupportedFormat);
//! assert_eq!(error.detected(), Some(litchi::FileFormat::Docx));
//! # }
//! # Ok::<(), litchi::Error>(())
//! ```
//!
//! Format-specific reads live on the handle a variant carries;
//! [`format`](mod@format) documents each handle:
//!
//! ```
//! # #[cfg(feature = "pptx")]
//! # {
//! use litchi::File;
//!
//! # let path = concat!(
//! #     env!("CARGO_MANIFEST_DIR"),
//! #     "/../../test-data/ooxml/pptx/SmartArt.pptx"
//! # );
//! let file = litchi::open(path)?;
//! if let File::Pptx(pptx) = &file {
//!     let (width, height) = pptx.slide_size()?;
//!     assert_eq!((width, height), (12_192_000, 6_858_000)); // EMU, 16:9
//!     assert_eq!(pptx.slide_catalog()?.len(), 1);
//! }
//! # }
//! # Ok::<(), litchi::Error>(())
//! ```
//!
//! With the `markdown` feature a document or presentation converts to
//! Markdown through `litchi::markdown::ToMarkdown`:
//!
//! ```
//! # #[cfg(all(feature = "docx", feature = "markdown"))]
//! # {
//! use litchi::markdown::ToMarkdown;
//!
//! # let path = concat!(
//! #     env!("CARGO_MANIFEST_DIR"),
//! #     "/../../test-data/ooxml/docx/documentProperties.docx"
//! # );
//! let file = litchi::open(path)?;
//! assert_eq!(file.to_markdown()?, "Hello World\\!\n\n");
//! # }
//! # Ok::<(), litchi::Error>(())
//! ```
//!
//! # Encrypted and signed files
//!
//! An encrypted input opens as [`File::Locked`]. Its reads refuse with
//! [`Code::DocumentProtected`] and a suggestion, until `Locked::unlock` (the
//! `encryption` feature) consumes the handle with the password. The password
//! never passes through [`OpenOptions`].
//!
//! ```
//! # #[cfg(all(feature = "doc", feature = "encryption"))]
//! # {
//! use litchi::{Code, EncryptionKind, File, FileFormat, Password};
//!
//! # let path = concat!(
//! #     env!("CARGO_MANIFEST_DIR"),
//! #     "/../../test-data/poi/test-data/document/password_password_cryptoapi.doc"
//! # );
//! let file = litchi::open(path)?;
//! let error = file.text().unwrap_err();
//! assert_eq!(error.code(), Code::DocumentProtected);
//! assert!(error.suggestion().is_some_and(|hint| hint.contains("Locked::unlock")));
//!
//! let file = match file {
//!     File::Locked(locked) => {
//!         assert_eq!(locked.format(), Some(FileFormat::Doc));
//!         assert_eq!(locked.encryption(), EncryptionKind::CryptoApi);
//!         locked.unlock(Password::new("password".to_owned()))?
//!     },
//!     file => file,
//! };
//! assert!(!file.text()?.is_empty());
//! # }
//! # Ok::<(), litchi::Error>(())
//! ```
//!
//! A signed OOXML or legacy binary input opens as [`File::Signed`]. It reads
//! like the file it wraps, and nothing is verified at open. Editing a signed
//! file requires consuming the signature first with [`Signed::unsign`], so
//! that no edit strips a signature silently.
//!
//! ```
//! # #[cfg(feature = "docx")]
//! # {
//! use litchi::{File, SignatureOrigin};
//!
//! # let path = concat!(
//! #     env!("CARGO_MANIFEST_DIR"),
//! #     "/../../test-data/poi/test-data/xmldsign/hello-world-signed.docx"
//! # );
//! let file = litchi::open(path)?;
//! assert!(file.is_signed());
//! assert_eq!(file.text()?.trim(), "Hello world");
//! if let File::Signed(signed) = file {
//!     assert_eq!(signed.origin(), SignatureOrigin::Opc);
//!     let unsigned: File = signed.unsign();
//!     assert!(!unsigned.is_signed());
//! }
//! # }
//! # Ok::<(), litchi::Error>(())
//! ```
//!
//! # Errors
//!
//! Every failure of the unified API is one [`Error`] with a stable [`Code`]
//! (OfficeCLI's code list plus Litchi's `unsupported_capability`,
//! `limit_exceeded` and `cancelled`), a message, an optional suggestion
//! ([`Error::suggestion`]), the format involved and the typed error of the
//! format crate that failed ([`Error::leaf`], also its
//! [`source`](std::error::Error::source)).
//!
//! ```
//! # #[cfg(feature = "docx")]
//! # {
//! let error = litchi::open_bytes(b"plain text".to_vec()).unwrap_err();
//! assert_eq!(error.code(), litchi::Code::UnsupportedFormat);
//! assert_eq!(error.code().as_str(), "unsupported_format");
//! assert_eq!(error.detected(), None);
//! # }
//! ```
//!
//! # Apple iWork
//!
//! Pages, Keynote and Numbers documents are read through the `iwork` module
//! (the `iwork` feature). [`open`] recognizes them and refuses them with
//! [`Code::UnsupportedFormat`] and a suggestion to use
//! `litchi::iwork::Document::open`.
//!
//! # Unified operations
//!
//! [`File`] provides typed reads, paths, queries and views. DOCX also supports
//! immutable paragraph/run edits, creation, stream publication and atomic path
//! saving. [`File::capabilities`] lists the supported verbs and contextual
//! gaps for the current handle. Batches and the other formats' unified write
//! handlers are still in progress; the format modules expose their leaf APIs.
//!
//! # Low-level modules
//!
//! Each format crate is re-exported under its own module for format-specific
//! reading and writing: `doc`, `ppt` and `xls` for the legacy binary formats,
//! `docx`, `pptx`, `xlsx` and `xlsb` for the OOXML owners, `odt`, `ods` and
//! `odp` for OpenDocument, and `rtf`. [`common`] holds the shared vocabulary
//! ([`common::ShapeType`], [`common::RGBColor`], [`common::Length`], ...).
//! Most users need only [`open`], and reach for these modules when they
//! write a document or need a format-specific feature.

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
        #[cfg(all(
            any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"),
            any(unix, windows)
        ))]
        pub use crate::detection_smart::{
            PreparedDetection, detect_and_prepare, detect_and_prepare_with_limits,
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

/// Format-neutral repair reports
///
/// See [`ToleranceReport`] and [`File::tolerance_report`].
#[cfg(any(
    feature = "doc",
    feature = "docx",
    feature = "rtf",
    feature = "odt",
    feature = "pages",
    feature = "ppt",
    feature = "pptx",
    feature = "odp",
    feature = "keynote",
    feature = "xls",
    feature = "xlsx",
    feature = "xlsb",
    feature = "ods",
    feature = "numbers"
))]
pub mod tolerance;

/// Image processing and conversion module
///
/// Provides functionality to parse and convert Office Drawing formats
/// (EMF, WMF, PICT) to modern image standards (PNG, JPEG, WebP).
///
/// **Note**: This requires the `images` feature to be enabled.
#[cfg(feature = "images")]
pub mod images;

/// Spreadsheet vocabulary: the generic sheet traits, text workbooks and
/// formula evaluation.
#[cfg(feature = "sheet")]
pub mod sheet;

/// Markdown conversion
///
/// [`markdown::ToMarkdown`] converts a [`File`] holding a document or a
/// presentation, and its family views, to Markdown.
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
        Error, Keywords, Props, Result, XmlError, custom, custom_xml, embedded, external_link,
        leniency, mce, properties, relationships, ribbon, spreadsheet_xml_maps, web, xml, xml_name,
    };
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
/// [`open`] reads RTF files as [`File::Rtf`]; this module is the RTF format
/// crate, for format-specific reading and writing.
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

// The unified operations API (ADR 0033, `litchi-ops`): `open` and `File`
// are the entry point. The openers wrap the `litchi-ops` ones with the
// facade's iWork classifier.
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
mod unified;

#[cfg(feature = "doc")]
pub use litchi_ops::DocOptions;
#[cfg(any(feature = "doc", feature = "docx", feature = "rtf", feature = "odt"))]
pub use litchi_ops::DocumentView;
#[cfg(all(
    feature = "encryption",
    any(
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
    )
))]
pub use litchi_ops::Password;
#[cfg(any(feature = "ppt", feature = "pptx", feature = "odp"))]
pub use litchi_ops::PresentationView;
#[cfg(any(feature = "xls", feature = "xlsx", feature = "xlsb", feature = "ods"))]
pub use litchi_ops::WorkbookView;
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
pub use litchi_ops::{
    Added, Applied, BatchCommit, BatchItem, BatchMode, Code, Conformance, CreateOptions,
    EncryptionKind, Error, Family, File, Flavor, IntoNode, IntoPlacement, IntoProps, IntoTarget,
    ItemResult, Key, Locked, NewNode, OpenOptions, Overwrite, Placement, PropValue, Props, Result,
    SaveOptions, SaveTo, Saved, SignatureOrigin, Signed, SourceMode, TargetArg, batch,
    capabilities, compat, create, create_with, error, format, new_node, node, path, placement,
    props, query, target, transaction, view,
};
#[cfg(all(
    any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"),
    any(unix, windows)
))]
pub use litchi_ops::{open_prepared, open_prepared_with};
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
pub use unified::{open, open_bytes, open_bytes_with, open_with};

// One entry of `format::pptx::Pptx::slide_catalog`, also named at the root.
#[cfg(feature = "pptx")]
pub use litchi_ops::format::pptx::SlideDescriptor;

#[cfg(any(
    feature = "doc",
    feature = "docx",
    feature = "rtf",
    feature = "odt",
    feature = "pages",
    feature = "ppt",
    feature = "pptx",
    feature = "odp",
    feature = "keynote",
    feature = "xls",
    feature = "xlsx",
    feature = "xlsb",
    feature = "ods",
    feature = "numbers"
))]
pub use tolerance::ToleranceReport;

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
    any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"),
    any(unix, windows)
))]
pub use common::detection::{
    PreparedDetection, detect_and_prepare, detect_and_prepare_with_limits,
};

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
        let _: Option<super::docx::SourceLimits> = None;
        let _: Option<super::pptx::SourceLimits> = None;
        let _: Option<super::xlsx::SourceLimits> = None;
        let _: Option<super::xlsb::SourceLimits> = None;
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
