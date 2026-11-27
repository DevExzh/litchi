//! Litchi: one Rust API for Microsoft Office, OpenDocument and Apple iWork
//! documents.
//!
//! Enable a Cargo feature such as `docx`, `xlsx`, or `pages` to open that
//! format, or `all-formats` for every supported format. Each enabled format
//! is available through the unified `File` API.
//!
#![cfg_attr(litchi_unified, doc = include_str!("unified_api.md"))]
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
    #[cfg(litchi_unified)]
    pub use crate::detection_smart::{
        SourceLimits, detect_file_format, detect_file_format_from_bytes,
    };

    /// Detection re-exports — merges `litchi-core`'s signature detection with
    /// the umbrella's smart-detection entry points.
    pub mod detection {
        #[cfg(litchi_unified)]
        pub use crate::detection_smart::{
            DetectedFormat, SourceLimits, detect_file_format, detect_file_format_from_bytes,
            detect_format_from_reader, detect_format_smart,
        };
        #[cfg(all(
            any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"),
            any(unix, windows)
        ))]
        pub use crate::detection_smart::{PreparedDetection, detect_and_prepare};
        pub use litchi_core::detection::*;
    }
}

// Smart format detection (depends on per-format crates; can't live in litchi-core).
#[cfg(litchi_unified)]
pub mod detection_smart;

#[cfg(feature = "yaml")]
mod metadata_ext;

#[cfg(feature = "yaml")]
pub use metadata_ext::MetadataYaml;

/// Format-neutral repair reports
///
/// See [`ToleranceReport`] and [`File::tolerance_report`].
#[cfg(litchi_unified)]
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
        Error, Result, XmlError, custom, custom_xml, embedded, external_link, leniency, mce,
        properties, relationships, ribbon, spreadsheet_xml_maps, web, xml, xml_name,
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

/// Math equations: conversion between OMML, MTEF, MathML and LaTeX.
///
/// **Note**: This requires the `equation` feature to be enabled.
#[cfg(feature = "equation")]
pub mod equation {
    pub use litchi_equation::*;
}

/// Apple Pages package and semantic APIs.
#[cfg(feature = "pages")]
pub mod pages {
    pub use litchi_pages::*;

    /// Immutable Pages values: the read-only `Document` projection and its
    /// sections, with the ingress limit profiles.
    ///
    /// No native archive, generated message, or compatibility projection is
    /// re-exported here.
    ///
    /// The namespace holds types only; it has no reader of its own. Open a
    /// Pages file with [`crate::open`] (a `File::Pages`), or with
    /// [`crate::pages::Package`] and its `document()` projection for
    /// format-specific work.
    pub mod semantic {
        pub use litchi_pages::{
            Body, Document, DocumentStats, Error, LimitKind, OpenOptions, PackageError, Position,
            Result, Root, Section, SectionSelector, SectionType, SelectorError, SelectorResult,
            SemanticLimitKind, SemanticLimits, SourceLimitKind, SourceLimits, TextPosition,
            TextSpan,
        };
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
    ///
    /// The namespace holds types only; it has no reader of its own. Open a
    /// Keynote file with [`crate::open`] (a `File::Keynote`), or with
    /// [`crate::keynote::Presentation::open`] for format-specific work.
    pub mod semantic {
        pub use litchi_keynote::presentation;
        pub use litchi_keynote::show::{Mode, Settings, Show, Size};
        pub use litchi_keynote::{
            Build, Error, LimitKind, OpenOptions, Position, Presentation, Result, Seconds,
            SemanticLimitKind, SemanticLimits, SemanticLimitsBuilder, Slide, SlideSelector,
            SlideSelectorError, SlideSelectorResult, SourceLimitKind, SourceLimits,
            SourceLimitsBuilder, TextPosition, TextSpan, Transition,
        };
    }
}

/// Apple Numbers package and semantic APIs.
#[cfg(feature = "numbers")]
pub mod numbers {
    // Keep the format facade explicit: a leaf-package glob would silently
    // publish any future package, generated, or physical helper module.
    // The package owner, its read-only `Workbook` projection, the ingress
    // profiles and the focused transaction modules keep their ADR 0034
    // names.
    pub use litchi_numbers::{
        AddressError, Area, At, Builder, Cell, ColumnDeletion, CoordinateError, Dimension,
        Dimensions, Error, Grid, GridBudget, Identity, InsertError, InsertResult, LimitKind,
        MAX_MATERIALIZED_CELLS, MAX_OBJECTS, MAX_REFERENCES, MAX_SHEETS, MAX_TABLES,
        MAX_TEXT_BYTES, MergeReader, OpenOptions, Package, PayloadLimitKind, Points, ResourceError,
        Result, RowDeletion, SemanticLimitKind, SemanticLimits, SemanticLimitsBuilder,
        SemanticPath, Sheet, SheetBuilder, SheetSelector, Size, SourceLimitKind, SourceLimits,
        SourceLimitsBuilder, Table, TableBuilder, TableError, TableSelector, TableSelectorError,
        View, Workbook, WriteError, compatibility_tables_from_bytes,
        compatibility_tables_from_bytes_with_options,
    };
    pub use litchi_numbers::{
        cell, create, formula, limits, names, package, selector, sheet, table, workbook,
    };

    /// Archive-free, immutable Numbers values.
    ///
    /// This namespace intentionally excludes the package adapter and the
    /// global compatibility-table projection. Use [`crate::numbers`] when exact
    /// package preservation or editing is required via [`crate::numbers::Package`].
    /// The checked source and semantic limit profiles and their errors remain
    /// in this namespace for a complete immutable-reader boundary.
    ///
    /// The namespace holds types only; it has no reader of its own. Open a
    /// Numbers file with [`crate::open`] (a `File::Numbers`), or with
    /// [`crate::numbers::Workbook::open`] for format-specific work.
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
                AxisReference, BinaryOperator, CachedValue, CellReference, Error,
            };
        }

        /// Archive-free Numbers table coordinates and cell presence values.
        ///
        /// The nested `cells` module mirrors the focused format crate so a
        /// caller can use typed A1 coordinates and presence-preserving values
        /// without importing the package adapter.
        pub mod table {
            pub use litchi_numbers::{AddressError, Area, At, CoordinateError, Dimensions, View};

            pub mod cells {
                pub use litchi_numbers::table::cells::{State, Storage};
            }
        }

        pub use litchi_numbers::workbook::{Error, ReadError, ReadLimitKind, Result, Stats};
        pub use litchi_numbers::{
            OpenOptions, SemanticLimitKind, SemanticLimits, Sheet, SheetSelector, SourceLimitKind,
            SourceLimits, Table, TableSelector, TableSelectorError, Workbook,
        };
    }
}

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

/// Unified formula evaluation over [`File`]: `File::evaluate` and
/// `File::recalculate`.
#[cfg(all(
    feature = "eval",
    any(feature = "xls", feature = "xlsx", feature = "xlsb")
))]
pub use litchi_ops::calc;

// The unified operations API (ADR 0033, `litchi-ops`): `open` and `File`
// are the entry point. The openers wrap the `litchi-ops` ones with the
// facade's iWork classifier.
#[cfg(litchi_unified)]
mod unified;

#[cfg(feature = "doc")]
pub use litchi_ops::DocOptions;
#[cfg(any(feature = "doc", feature = "docx", feature = "rtf", feature = "odt"))]
pub use litchi_ops::DocumentView;
#[cfg(all(feature = "encryption", litchi_unified))]
pub use litchi_ops::Password;
#[cfg(any(feature = "ppt", feature = "pptx", feature = "odp"))]
pub use litchi_ops::PresentationView;
#[cfg(any(feature = "xls", feature = "xlsx", feature = "xlsb", feature = "ods"))]
pub use litchi_ops::WorkbookView;
#[cfg(litchi_unified)]
pub use litchi_ops::{
    Added, Applied, BatchCommit, BatchItem, BatchMode, Code, Conformance, CreateOptions,
    EditOptions, EncryptionKind, Error, Family, File, Flavor, IntoNode, IntoPlacement, IntoProps,
    IntoTarget, ItemResult, Key, Locked, NewNode, OpenOptions, Overwrite, Placement, PropValue,
    Props, Result, SaveOptions, SaveTo, Saved, SignatureOrigin, Signed, SourceMode, TargetArg,
    batch, capabilities, compat, create, create_with, error, format, import, merge, new_node, node,
    path, placement, props, query, raw, refresh, schema, target, transaction, view,
};
#[cfg(all(
    any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"),
    any(unix, windows)
))]
pub use litchi_ops::{open_prepared, open_prepared_with};
#[cfg(litchi_unified)]
pub use unified::{open, open_bytes, open_bytes_with, open_with};

// One entry of `format::pptx::Pptx::slide_catalog`, also named at the root.
#[cfg(feature = "pptx")]
pub use litchi_ops::format::pptx::SlideDescriptor;

#[cfg(litchi_unified)]
pub use tolerance::ToleranceReport;

// Re-export commonly used types
pub use common::{FileFormat, Length, PlaceholderType, RGBColor, ShapeType};

#[cfg(litchi_unified)]
pub use common::{detect_file_format, detect_file_format_from_bytes};

#[cfg(all(
    any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"),
    any(unix, windows)
))]
pub use common::detection::{PreparedDetection, detect_and_prepare};

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
        let _: Option<super::ods::Workbook> = None;
        let _: Option<super::odt::Package> = None;
        let _: Option<super::ods::Package> = None;
        let _: Option<super::odp::Package> = None;
        let _: Option<super::odt::Document> = None;
        let _: Option<super::odf_common::detect::Format> = None;
    }
}

#[cfg(test)]
#[path = "../build/format_features.rs"]
mod format_features;

/// The build script derives `cfg(litchi_unified)` from the format features
/// in `Cargo.toml`; these tests pin what it derives.
#[cfg(test)]
mod format_feature_tests {
    use super::format_features::format_features;

    /// The 21 claimed formats of `docs/FEATURE_MATRIX.md`, by facade feature.
    const FORMATS: [&str; 21] = [
        "doc",
        "docx",
        "ppt",
        "pptx",
        "xls",
        "xlsx",
        "xlsb",
        "rtf",
        "odt",
        "ods",
        "odp",
        "oth",
        "odm",
        "odg",
        "odc",
        "odi",
        "odb",
        "odf-formula",
        "pages",
        "keynote",
        "numbers",
    ];

    #[test]
    fn every_claimed_format_feature_gates_the_unified_entry_point() {
        let mut derived = format_features(include_str!("../Cargo.toml"));
        derived.sort_unstable();
        let mut expected = FORMATS.map(str::to_owned).to_vec();
        expected.sort_unstable();
        assert_eq!(
            derived, expected,
            "each format feature must enable \"litchi-ops/<format>\"; a feature that does not \
             builds a facade without `litchi::open`"
        );
    }

    #[test]
    fn every_format_feature_names_a_litchi_ops_feature() {
        let ops = include_str!("../../litchi-ops/Cargo.toml");
        let declared = |name: &str| {
            ops.lines().any(|line| {
                line.strip_prefix(name)
                    .is_some_and(|rest| rest.trim_start().starts_with('='))
            })
        };
        for feature in FORMATS {
            assert!(declared(feature), "litchi-ops has no `{feature}` feature");
        }
    }

    #[test]
    fn parser_reads_multi_line_arrays_and_skips_weak_items_and_comments() {
        let manifest = r#"
[package]
name = "x" # "litchi-ops/not-a-feature"

[features]
default = []
# odd = ["litchi-ops/commented"]
one = ["dep:a", "litchi-ops/one"]
two = [
    "dep:b",
    # "litchi-ops/ignored",
    "litchi-ops/two",
]
weak = ["litchi-ops?/encryption"]
aggregate = ["one", "two"]

[dependencies]
three = ["litchi-ops/three"]
"#;
        assert_eq!(format_features(manifest), ["one", "two"]);
    }

    #[cfg(litchi_unified)]
    #[test]
    fn the_unified_entry_point_is_present() {
        let error = crate::open("/nonexistent/litchi-missing.docx").unwrap_err();
        assert_eq!(error.code(), crate::Code::FileNotFound);
    }
}
