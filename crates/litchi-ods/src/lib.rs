//! `OpenDocument` Spreadsheet (`.ods`) support.
//!
//! The crate is organized by responsibility: immutable spreadsheet vocabulary
//! in [`model`], XML codecs in [`codec`], the immutable [`Package`] owner and
//! its primary transaction in [`document`], construction in [`authoring`],
//! and the read-only [`Workbook`] projection in [`facade`].
//!
//! ```
//! use litchi_ods::{Builder, Package, Sheet};
//!
//! # fn main() -> Result<(), Box<dyn std::error::Error>> {
//! let mut builder = Builder::new();
//! builder.add_sheet(Sheet::new("Data")?)?;
//! let package: Package = builder.build()?;
//! let workbook = package.workbook()?;
//! assert_eq!(workbook.sheet_count(), 1);
//! # Ok(())
//! # }
//! ```

#![forbid(unsafe_code)]
// ODF element models and the established spreadsheet facade deliberately use
// specification-shaped names, ownership, and enum layouts. Changing them for
// generic API heuristics would be a breaking migration rather than a codec fix.
#![allow(
    clippy::large_enum_variant,
    reason = "the public ODS API and ODF vocabulary retain specification-shaped names and ownership"
)]
// ODF codecs are ordered by document traversal and keep schema cases explicit.
// These style exceptions retain that audit order while all correctness lints
// continue to be warning-denied.
#![allow(
    clippy::unnecessary_sort_by,
    reason = "ODF codec declarations follow document order and retain explicit schema branches for reviewability"
)]

pub mod advanced;
pub mod annotations;
pub mod authoring;
pub mod builder;
pub mod charts;
pub mod codec;
pub mod content_validation;
pub mod data_pilot;
pub mod data_style;
pub mod database_range;
pub mod definitions;
pub mod document;
pub mod drawing;
pub mod embedded;
pub mod facade;
pub mod flat;
pub mod media;
pub mod metadata;
pub mod metadata_graphs;
pub mod model;
mod open_parse;
pub(crate) mod package;
pub mod protection;
pub mod settings;
pub mod sheet_metadata;
pub mod source_backed;
pub mod source_features;
pub mod streaming;
pub mod styles;
pub mod worksheet;
mod xml_reference;

#[cfg(test)]
mod internal_tests;

pub use charts::Chart;
pub use document::{EditOptions, Package};
pub use drawing::{Frame, Part};
pub use embedded::{Kind, Object, Parameter, Root};
pub use facade::{
    Builder, CellSelector, FormulaChange, HyperlinkChange, HyperlinkOperation, MAX_CELL_SELECTORS,
    SheetCatalogEntry, SourceBackedSpreadsheet, SourceBackedSpreadsheetCatalog, SourceReadMetrics,
    Workbook,
};
pub use litchi_core::Metadata;
pub use litchi_odf_common::rdf;
pub use litchi_odf_common::{
    Error, LimitKind, OpenOptions, SemanticLimitKind, SemanticLimits, SemanticLimitsBuilder,
    SourceLimitKind, SourceLimits, SourceLimitsBuilder,
};
pub use media::Image;
pub use model::hyperlink::{Actuate, Link, Show};
pub use model::names;
pub use model::source::CellRange;
pub use model::tracked_changes;
pub use model::{dde, scenario};
pub use settings::{Iteration, IterationStatus, NullDate, Settings};
pub use source_features::{
    Drawing, DrawingKind, Hyperlink as SourceHyperlink, Limits as SourceFeatureLimits,
    Sheet as SourceSheet, Snapshot as SourceFeatures,
};
pub use streaming::{
    PublicationError, PublicationFailureKind, ScalarStreamReport, StreamingCell, StreamingError,
    StreamingLimits, XmlAuditLimits, stream_scalar_rows_to,
};
pub use worksheet::{Cell, CellValue, CellView, Merge, Row, Sheet};
